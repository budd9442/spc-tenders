import os
import re
import time
import urllib.request
import urllib.parse
from bs4 import BeautifulSoup
from concurrent.futures import ThreadPoolExecutor
from typing import List, Dict, Any, Optional

SEARCH_URL = "https://www.spc.lk/tender-display.php"
BASE_URL = "https://www.spc.lk/"


def fetch_page_tenders(query_id: str, offset: int, max_retries: int = 3) -> List[Dict[str, Any]]:
    """Fetches and parses a single pagination page for a given query and offset."""
    data = urllib.parse.urlencode({
        'tenderType': '2',
        'tenderID': query_id,
        'offset': str(offset)
    }).encode('utf-8')

    req = urllib.request.Request(
        SEARCH_URL,
        data=data,
        headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'}
    )

    for attempt in range(max_retries):
        try:
            with urllib.request.urlopen(req, timeout=25) as resp:
                html = resp.read().decode('utf-8', errors='ignore')
            break
        except Exception as e:
            if attempt == max_retries - 1:
                print(f"[WARN] Failed fetching offset {offset} for query '{query_id}': {e}")
                return []
            time.sleep(2 * (attempt + 1))

    soup = BeautifulSoup(html, 'html.parser')
    tbl = soup.find('table', {'id': 'tender-display'})
    if not tbl:
        return []

    items = []
    for row in tbl.find_all('tr')[1:]:
        cols = [td.get_text(strip=True) for td in row.find_all('td')]
        a = row.find('a')
        if len(cols) > 2 and a and a.get('href'):
            raw_tno = cols[2]
            tno = re.sub(r'\s+', '', raw_tno)
            # Match DHS/SS or DHS/SA tenders for years 27 and 2027
            if re.match(r'^DHS/(?:SS|SA)/[^/]+/[^/]+/(?:27|2027)$', tno, re.I):
                href = a['href']
                full_href = urllib.parse.urljoin(BASE_URL, href)
                qs = urllib.parse.parse_qs(urllib.parse.urlparse(full_href).query)
                doc_api_url = qs.get('url', [None])[0] or full_href
                doc_id = qs.get('id', [None])[0] or ''
                safe_fn = tno.replace('/', '_') + '.pdf'

                category = "DHS/SS" if "DHS/SS" in tno.upper() else "DHS/SA"
                proc_type = "ICB" if "ICB" in tno.upper() else "NCB"

                items.append({
                    'tender_no': tno,
                    'original_tender_no': raw_tno,
                    'category': category,
                    'procurement_type': proc_type,
                    'status': cols[1] if len(cols) > 1 else 'Open',
                    'closing_date': cols[3] if len(cols) > 3 else '',
                    'view_url': full_href,
                    'document_url': doc_api_url,
                    'document_id': doc_id,
                    'filename': safe_fn
                })

    return items


def discover_all_online_tenders(query_ids: Optional[List[str]] = None) -> List[Dict[str, Any]]:
    """
    Queries backend search across both '/27' and '2027' tender IDs to capture
    all 2-digit and 4-digit year notations without manual browser interaction.
    """
    if query_ids is None:
        query_ids = ['/27', '2027']

    all_tenders = {}

    for qid in query_ids:
        # Determine total count for query
        data = urllib.parse.urlencode({'tenderType': '2', 'tenderID': qid, 'offset': '0'}).encode('utf-8')
        req = urllib.request.Request(
            SEARCH_URL,
            data=data,
            headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'}
        )
        try:
            with urllib.request.urlopen(req, timeout=20) as resp:
                html = resp.read().decode('utf-8', errors='ignore')
        except Exception as e:
            print(f"[ERROR] Failed initial discovery for query '{qid}': {e}")
            continue

        m = re.search(r'to\s+(\d+)\s+tenders', html)
        total = int(m.group(1)) if m else 900
        offsets = list(range(0, total, 100))
        print(f"Query '{qid}': found ~{total} tenders across {len(offsets)} offsets.")

        with ThreadPoolExecutor(max_workers=5) as pool:
            futures = [pool.submit(fetch_page_tenders, qid, off) for off in offsets]
            for f in futures:
                for item in f.result():
                    all_tenders[item['tender_no']] = item

    print(f"Total unique matching tenders discovered online: {len(all_tenders)}")
    return list(all_tenders.values())


def download_pdf(url: str, output_path: str, max_retries: int = 3) -> bool:
    """Downloads a tender PDF with exponential backoff and verification."""
    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    req = urllib.request.Request(
        url,
        headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'}
    )

    for attempt in range(max_retries):
        try:
            with urllib.request.urlopen(req, timeout=30) as resp:
                data = resp.read()
            if len(data) > 1000 and data[:4] == b'%PDF':
                with open(output_path, 'wb') as f:
                    f.write(data)
                return True
            elif len(data) > 500:
                # Sometimes server returns PDF with small leading whitespace
                with open(output_path, 'wb') as f:
                    f.write(data)
                return True
            else:
                print(f"[WARN] Downloaded data too small ({len(data)} bytes) for {url}")
        except Exception as e:
            if attempt == max_retries - 1:
                print(f"[ERROR] Failed downloading {url}: {e}")
                return False
            time.sleep(2 * (attempt + 1))

    return False
