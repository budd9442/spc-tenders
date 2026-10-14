import os
import re
import json
from pathlib import Path
from typing import List, Dict, Any, Optional, Tuple
import pymupdf


class SPCTenderTableExtractor:
    """
    Extracts structured tender items and metadata from SPC/DHS tender documents.
    Features:
      - Dynamic column boundary calibration (A..F) based on actual page geometry.
      - 8-digit SR code anchor detection in Column B.
      - Table marker guards to prevent false-positive General Order Condition exclusions.
      - Sub-pixel word-level footer cutoff to prevent notes from leaking into item rows.
      - Multi-page continuation handling for multi-clause specifications.
      - Smart post-processing for delivery percentages, currency prefixes, and packing notes.
      - OCR fallback for scanned legacy tenders.
    """

    def __init__(self):
        # Fallback standard column boundaries calibrated for SPC/DHS tender templates
        self.default_col_bounds = [
            (0, 92),        # (A) Item No
            (92, 155),      # (B) SR Number (8-digit code)
            (155, 340),     # (C) Item Description / Specifications
            (340, 402),     # (D) Quantity
            (402, 485),     # (E) Delivery Schedule
            (485, 9999)     # (F) Bid Bond Value
        ]

    def extract_metadata(self, doc: pymupdf.Document) -> Dict[str, Any]:
        """Extracts key tender metadata from document header, text layers, and OCR."""
        metadata = {
            "procurement_no": None,
            "order_list_no": None,
            "date_of_issue": None,
            "closing_date": None,
            "bid_validity": None,
            "bid_bond_validity": None,
            "tender_fee": None,
            "page_count": len(doc)
        }

        full_text = ""
        for i in range(len(doc)):
            full_text += f"\n--- PAGE {i+1} ---\n" + doc[i].get_text()

        # If scanned document with minimal text layer, attempt OCR on key pages
        if len(full_text.strip()) < 100 and len(doc) > 0:
            for p_try in [0, min(1, len(doc)-1), min(9, len(doc)-1)]:
                try:
                    tp = doc[p_try].get_textpage_ocr(dpi=200)
                    full_text += f"\n--- PAGE {p_try+1} OCR ---\n" + tp.extractTEXT()
                except Exception:
                    pass

        # Normalize spacing around slashes (e.g. DHS/SA /NCB/26/27 -> DHS/SA/NCB/26/27)
        normalized_text = re.sub(
            r'DHS\s*/\s*([A-Za-z0-9_-]+)\s*/\s*([A-Za-z0-9_-]+)\s*/\s*([A-Za-z0-9_-]+)\s*/\s*(\d{2,4})',
            r'DHS/\1/\2/\3/\4',
            full_text,
            flags=re.IGNORECASE
        )

        # 1. Procurement / Tender Number
        proc_match = re.search(r'\b(DHS/[A-Z0-9/_-]+(?:/(?:27|2027|\d{2,4})))\b', normalized_text, re.IGNORECASE)
        if proc_match:
            metadata["procurement_no"] = proc_match.group(1).strip()
        else:
            proc_match2 = re.search(
                r'(?:PROCUREMENT\s*(?:NO|REFERENCE)|TENDER\s*NO)[\s.:/-]+(DHS/[A-Z0-9/_-]+)',
                normalized_text,
                re.IGNORECASE
            )
            if proc_match2:
                metadata["procurement_no"] = proc_match2.group(1).strip()

        # 2. MSD Order List Number
        order_match = re.search(r'\b(\d{4}/SPC/[A-Za-z0-9/_-]+)\b', normalized_text)
        if order_match:
            metadata["order_list_no"] = order_match.group(1).strip()

        # 3. Date of Issue
        issue_match = re.search(
            r'(?:DATE\s+OF\s+(?:BID\s+)?(?:INVITATION|ISSUE))[\s.:/-]+([0-9A-Za-z\s,/-]+?)(?=\s*(?:\n|CLOSING))',
            normalized_text,
            re.IGNORECASE
        )
        if issue_match:
            metadata["date_of_issue"] = issue_match.group(1).strip()

        # 4. Closing Date
        closing_match = re.search(
            r'CLOSING\s+(?:ON|DATE(?:\s*&\s*TIME)?)[\s.:/-]+([^\n\r]+?)(?=\s*(?:at\s*___|SRI|\n|$))',
            normalized_text,
            re.IGNORECASE
        )
        if closing_match:
            clean_closing = re.sub(r'\s+', ' ', closing_match.group(1)).strip()
            metadata["closing_date"] = clean_closing

        # 5. Bid Validity
        bid_val_match = re.search(r'Bid\s+valid\s+till\s*[:.\s-]+([^\n\r.]+)', normalized_text, re.IGNORECASE)
        if bid_val_match:
            metadata["bid_validity"] = bid_val_match.group(1).strip()

        # 6. Bid Bond Validity
        bond_val_match = re.search(r'Bid\s+Bond\s+valid\s+till\s*[:.\s-]+([^\n\r.]+)', normalized_text, re.IGNORECASE)
        if bond_val_match:
            metadata["bid_bond_validity"] = bond_val_match.group(1).strip()

        # 7. Tender Fee
        fee_match = re.search(r'fee\s+of\s+(LKR\s*[\d,]+(?:\s*/=)?(?:\s*\+\s*taxes)?)', normalized_text, re.IGNORECASE)
        if fee_match:
            metadata["tender_fee"] = fee_match.group(1).strip()

        return metadata

    def calibrate_columns(self, words: List[Tuple]) -> List[Tuple[float, float]]:
        """
        Dynamically calculates column boundaries based on actual header markers (A)...(F)
        or header keywords on this specific page.
        """
        hdr_D = [w for w in words if w[4].upper() in ['(D)', 'D', 'QTY', 'QUANTITY'] and 340 <= w[0] <= 430 and w[1] > 40]
        hdr_E = [w for w in words if w[4].upper() in ['(E)', 'E', 'DELIVERY'] and 400 <= w[0] <= 495 and w[1] > 40]
        hdr_F = [w for w in words if (w[4].upper() in ['(F)', 'F'] or 'BOND' in w[4].upper()) and w[0] >= 480 and w[1] > 40]

        if hdr_D and hdr_E:
            d_x = min(w[0] for w in hdr_D)
            e_x = min(w[0] for w in hdr_E)
            f_x = min(w[0] for w in hdr_F) if hdr_F else max(e_x + 50.0, 490.0)

            bound_cd = min(350.0, max(330.0, d_x - 22.0))
            bound_de = max(375.0, min(455.0, (d_x + e_x) / 2.0))
            bound_ef = max(440.0, min(520.0, (e_x + f_x) / 2.0))
            return [
                (0, 92),
                (92, 155),
                (155, bound_cd),
                (bound_cd, bound_de),
                (bound_de, bound_ef),
                (bound_ef, 9999)
            ]
        return self.default_col_bounds

    def extract_items_table(self, doc: pymupdf.Document) -> List[Dict[str, Any]]:
        """
        Extracts all tender items across all pages belonging to the Schedule of Requirements.
        Handles scanned PDFs with OCR fallback, multi-page continuation, and column calibration.
        """
        items = []
        last_item = None

        # Check if entire document is scanned (no text layer on initial pages)
        is_scanned_doc = False
        words_check = sum(len(doc[i].get_text('words')) for i in range(min(3, len(doc))))
        if words_check < 40 and len(doc) > 0:
            is_scanned_doc = True

        for page_idx in range(len(doc)):
            page = doc[page_idx]

            # If document is scanned, run OCR on first 12 pages
            if is_scanned_doc:
                if page_idx >= 12:
                    break
                try:
                    tp = page.get_textpage_ocr(dpi=200)
                    words = tp.extractWORDS()
                except Exception:
                    words = []
            else:
                words = page.get_text('words')

            if not words:
                continue

            # Identify SR anchors in Column B (x roughly 45 to 165)
            sr_anchors = []
            for w in words:
                if 45 <= w[0] <= 165:
                    if re.match(r'^\d{8}$', w[4]):
                        sr_anchors.append((w, w[4]))
                    elif is_scanned_doc and re.search(r'\d{8}$', w[4]):
                        sr_anchors.append((w, re.search(r'(\d{8})$', w[4]).group(1)))

            # Exclude GOC (General Order Conditions) pages UNLESS page has clear table markers
            page_full_text = page.get_text().lower()
            has_table_marker = any(kw in page_full_text for kw in ['order list', 'item no', 'sr no', 'bid no', '(a)', '(b)'])
            if not has_table_marker:
                page_text_lower = ' '.join(w[4].lower() for w in words[:60])
                if 'general order conditions' in page_text_lower or 'part a-general order' in page_text_lower:
                    sr_anchors = []

            if sr_anchors:
                sr_anchors.sort(key=lambda a: a[0][1])

                # Calibrate column bounds
                col_bounds = self.calibrate_columns(words)
                (b_a, b_b, b_c, b_d, b_e, b_f) = col_bounds

                # Determine table footer cutoff y
                footer_y = page.rect.height
                for b in page.get_text('blocks'):
                    if b[1] > sr_anchors[-1][0][1]:
                        b_lower = b[4].lower()
                        if any(kw in b_lower for kw in [
                            'sufficient quantity', 'bid valid', 'when the required value',
                            'non-refundable fee', 'procurement no', 'procurement reference',
                            'common conditions', 'evaluation as tender samples', 'pleaserefer'
                        ]):
                            if b[1] < footer_y:
                                footer_y = b[1]

                # Word-level fallback when footer is merged into the last row's block
                for w in words:
                    if w[1] > sr_anchors[-1][0][1] + 10:
                        w_lower = w[4].lower()
                        if any(kw in w_lower for kw in [
                            'sufficient', 'representative', 'evaluation', 'validity', 'non-refundable'
                        ]):
                            if w[1] - 4 < footer_y:
                                footer_y = w[1] - 4

                for i, (anchor_w, sr_code) in enumerate(sr_anchors):
                    y_start = anchor_w[1] - 8
                    y_end = sr_anchors[i + 1][0][1] - 8 if i + 1 < len(sr_anchors) else footer_y
                    row_words = [w for w in words if y_start <= w[1] < y_end]
                    row_words.sort(key=lambda w: (round(w[1] / 6) * 6, w[0]))

                    cols = {
                        'item_no': [],
                        'desc': [],
                        'qty': [],
                        'delivery': [],
                        'bid_bond': []
                    }

                    for w in row_words:
                        x0 = w[0]
                        w_str = w[4]
                        if w_str.upper() in ['ITEM', 'NO', 'NO.', '(A)', '(B)', '(C)', '(D)', '(E)', '(F)', '|']:
                            continue
                        if re.sub(r'[^0-9]', '', w_str) == sr_code:
                            continue

                        if x0 < b_a[1]:
                            cols['item_no'].append(w_str)
                        elif b_b[0] <= x0 < b_b[1]:
                            pass  # SR column
                        elif b_c[0] <= x0 < b_c[1]:
                            cols['desc'].append(w_str)
                        elif b_d[0] <= x0 < b_d[1]:
                            cols['qty'].append(w_str)
                        elif b_e[0] <= x0 < b_e[1]:
                            cols['delivery'].append(w_str)
                        else:
                            cols['bid_bond'].append(w_str)

                    raw_item_no = ' '.join(cols['item_no']).strip()
                    m_item = re.search(r'\b(\d{1,3})\b', raw_item_no)
                    clean_item_no = m_item.group(1) if m_item else str(len(items) + 1)
                    try:
                        if int(clean_item_no) > len(items) + 15:
                            clean_item_no = str(len(items) + 1)
                    except ValueError:
                        pass

                    raw_desc = ' '.join(cols['desc']).strip()
                    raw_qty = ' '.join(cols['qty']).strip()
                    raw_del = ' '.join(cols['delivery']).strip()
                    raw_bond = ' '.join(cols['bid_bond']).strip()

                    # Smart Quantity & Delivery Schedule separation
                    del_matches = re.findall(r'(\d+(?:%|\s*Nos|\s*Pack|\s*Vials)?\s*[-–]\s*\d{2}/\d{4})', raw_qty)
                    if del_matches:
                        for dm in del_matches:
                            raw_qty = raw_qty.replace(dm, '').strip()
                            raw_del = (dm + ' ' + raw_del).strip()

                    # Embedded percentage delivery in qty (e.g. "100 100% Nos" -> qty="100 Nos", del="100% ...")
                    pct_embedded = re.search(r'\b(\d+%\s*[-–]?)\b', raw_qty)
                    if pct_embedded:
                        matched_pct = pct_embedded.group(1).strip()
                        raw_qty = re.sub(r'\s+', ' ', raw_qty.replace(matched_pct, '')).strip()
                        raw_del = (matched_pct + ' ' + raw_del).strip()

                    # Smart Delivery & Bid Bond separation (if currency symbol leaked into delivery)
                    bond_match = re.search(r'(LKR\s*[\d,]+(?:\s*/=)?|USD\s*[\d,]+(?:\s*/=)?)', raw_del)
                    if bond_match and not raw_bond:
                        raw_bond = bond_match.group(1)
                        raw_del = raw_del.replace(raw_bond, '').strip()

                    curr_tail = re.search(r'\b(LKR|USD|EUR)\s*$', raw_del)
                    if curr_tail:
                        curr = curr_tail.group(1)
                        raw_del = raw_del[:curr_tail.start()].strip()
                        if raw_bond:
                            if not re.search(r'\b(LKR|USD|EUR)\b', raw_bond):
                                raw_bond = f"{curr} {raw_bond}".strip()
                        else:
                            raw_bond = curr

                    # Check if delivery text leaked into bid bond
                    del_keywords = ['lot', 'sets', 'months', 'after', 'immediately', 'weeks', 'days']
                    if raw_bond and any(kw in raw_bond.lower() for kw in del_keywords) and not re.search(r'\b(lkr|usd|eur|\d{3,})\b', raw_bond.lower()):
                        raw_del = (raw_del + ' ' + raw_bond).strip()
                        raw_bond = None

                    # Clean trailing packing note from item no, delivery, or quantity
                    for field_val in [raw_item_no, raw_qty, raw_del]:
                        if 'Packing' in field_val:
                            pack_note = re.search(r'(Packing\s*:\s*[^\n]+)', field_val)
                            if pack_note:
                                raw_desc += ' ' + pack_note.group(1)
                                if field_val == raw_del:
                                    raw_del = raw_del.replace(pack_note.group(1), '').strip()
                                elif field_val == raw_qty:
                                    raw_qty = raw_qty.replace(pack_note.group(1), '').strip()

                    clean_bond = raw_bond if raw_bond and not re.match(r'^[-_|\s]+$', raw_bond) else None

                    item_entry = {
                        "item_no": clean_item_no,
                        "sr_no": sr_code,
                        "description": raw_desc,
                        "quantity": raw_qty,
                        "delivery_schedule": raw_del,
                        "bid_bond": clean_bond,
                        "source_page": page_idx + 1
                    }
                    items.append(item_entry)
                    last_item = item_entry

            elif last_item is not None and page_idx < 12:
                # Multi-page continuation of the previous item description on the immediate next page
                cont_blocks = []
                for b in page.get_text('blocks'):
                    b_text = b[4].strip()
                    b_lower = b_text.lower()
                    if any(kw in b_lower for kw in [
                        'when the required value', 'bid valid till', 'non-refundable fee',
                        'please refer', 'common conditions', 'conditions for supply',
                        'part a-general order', 'general order conditions'
                    ]):
                        break
                    # Ignore standalone page numbers
                    if re.match(r'^\d{1,2}$', b_text):
                        continue
                    cont_blocks.append(b_text)

                if cont_blocks:
                    last_item["description"] += " " + " ".join(cont_blocks)
                last_item = None  # Continuation is only allowed for the immediate next page

        return items

    def process_pdf(self, pdf_path: str) -> Dict[str, Any]:
        """Processes a single PDF file and returns its structured tender record."""
        try:
            doc = pymupdf.open(pdf_path)
            meta = self.extract_metadata(doc)
            items = self.extract_items_table(doc)

            return {
                "file_path": os.path.abspath(pdf_path),
                "file_name": os.path.basename(pdf_path),
                "metadata": meta,
                "item_count": len(items),
                "items": items,
                "status": "SUCCESS"
            }
        except Exception as e:
            return {
                "file_path": os.path.abspath(pdf_path),
                "file_name": os.path.basename(pdf_path),
                "status": "ERROR",
                "error": str(e)
            }
