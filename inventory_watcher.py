#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
inventory_watcher.py — مراقب مخزون النور
يقرأ ملفات الإكسل كل دقيقة ويحدّث HTML + JSON + يرفع على GitHub
"""

import os
import sys
import json
import time
import subprocess
import traceback
from datetime import datetime, timedelta
from collections import defaultdict

sys.stdout.reconfigure(encoding='utf-8')
sys.stderr.reconfigure(encoding='utf-8')

INVENTORY_DIR = r'D:\Mostafa Ibrahim\جرد يومي'
INVENTORY_FILE = os.path.join(INVENTORY_DIR, 'تقارير  ارصدة المخازن.xlsx')
INVOICES_DIR = r'D:\Mostafa Ibrahim\آخر فاتورة للعميل_files'
INVOICES_FILE = os.path.join(INVOICES_DIR, 'اخر فاتورة لعميل.xlsx')
RECONCILIATION_FILES = [
    os.path.join(INVENTORY_DIR, 'تقرير حركة الاصناف تفصيلي .xlsx'),
    os.path.join(INVENTORY_DIR, 'تقرير حركة الاصناف تفصيلي.xlsx')
]
REPO_DIR = r'C:\Users\GoldenTech\inventory-nour'
HTML_FILE = os.path.join(REPO_DIR, 'index.html')
JSON_FILE = os.path.join(REPO_DIR, 'inventory_data.json')
CACHE_REC_FILE = os.path.join(REPO_DIR, 'reconciliations_cache.json')
LOG_FILE = os.path.join(REPO_DIR, 'watcher.log')

# الأصناف المستبعدة (تصنيع / عمال / نقل)
EXCLUDED_KEYWORDS = ['تصنيع', 'عمال', 'نقل', 'مصاريف نقل', 'أجور عمال', 'أجور تصنيع', 'مصنعية', 'أجور نقل']

os.makedirs(REPO_DIR, exist_ok=True)

import warnings
warnings.filterwarnings('ignore')

try:
    import openpyxl
except ImportError:
    print("ERROR: openpyxl not installed", file=sys.stderr)
    sys.exit(1)


def log(msg):
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    line = f'[{timestamp}] {msg}'
    print(line)
    try:
        with open(LOG_FILE, 'a', encoding='utf-8') as f:
            f.write(line + '\n')
    except:
        pass


def safe_float(val):
    if val is None or val == '':
        return 0
    try:
        return float(val)
    except (ValueError, TypeError):
        return 0


def is_excluded(name, group=''):
    """Check if item should be excluded based on name or group"""
    text = f"{name} {group}".strip()
    for kw in EXCLUDED_KEYWORDS:
        if kw in text:
            return True
    return False


def find_latest_file(directory, keywords, default_path=None):
    """Find the latest matching file in directory based on keywords, ignoring temp files (~$)"""
    candidates = []
    if os.path.exists(directory):
        try:
            for fname in os.listdir(directory):
                if fname.startswith('~$') or not fname.lower().endswith(('.xlsx', '.xlsm', '.xls')):
                    continue
                if all(kw in fname for kw in keywords):
                    full_p = os.path.join(directory, fname)
                    candidates.append(full_p)
        except Exception as e:
            log(f"Error scanning directory {directory}: {e}")
    if default_path and os.path.exists(default_path) and default_path not in candidates:
        candidates.append(default_path)
    if not candidates:
        return default_path if (default_path and os.path.exists(default_path)) else None
    return max(candidates, key=lambda p: os.path.getmtime(p))


def safe_load_workbook(file_path, data_only=True, read_only=False, retries=5, delay=1.0):
    """Safely open workbook with retries in case Excel/ERP has the file locked"""
    for attempt in range(retries):
        try:
            return openpyxl.load_workbook(file_path, data_only=data_only, read_only=read_only)
        except PermissionError:
            log(f"File {os.path.basename(file_path)} locked by Excel (attempt {attempt+1}/{retries}), retrying in {delay}s...")
            time.sleep(delay)
        except Exception as e:
            if attempt == retries - 1:
                raise e
            time.sleep(delay)
    raise IOError(f"Could not open {file_path} after {retries} attempts")


def read_inventory():
    target_file = find_latest_file(INVENTORY_DIR, ['ارصدة', 'المخازن'], INVENTORY_FILE)
    if not target_file or not os.path.exists(target_file):
        log(f'WARNING: Inventory balances file not found in {INVENTORY_DIR}')
        return [], ''
    try:
        mtime_str = datetime.fromtimestamp(os.path.getmtime(target_file)).strftime('%Y-%m-%d %H:%M:%S')
        log(f'Reading inventory balances from: {os.path.basename(target_file)} (modified: {mtime_str})')
        wb = safe_load_workbook(target_file, data_only=True, read_only=False)
        ws = wb.active
        items = []
        print_date = ''
        try:
            cell_val = str(ws.cell(row=2, column=1).value or '')
            if 'Print Date:' in cell_val:
                print_date = cell_val.replace('Print Date:', '').strip()
        except:
            pass
        for row in ws.iter_rows(min_row=4, values_only=True):
            if not row[0] and not row[2]:
                continue
            items.append({
                'store': str(row[0] or '').strip(),
                'code': str(row[1] or '').strip(),
                'name': str(row[2] or '').strip(),
                'unit': str(row[3] or '').strip(),
                'balance': safe_float(row[4]),
                'detailed': str(row[5] or '').strip(),
                'detained': safe_float(row[6]),
                'group': str(row[7] or '').strip(),
                'supplier': str(row[8] or '').strip(),
            })
        wb.close()
        items = [i for i in items if not is_excluded(i['name'], i['group'])]
        items.sort(key=lambda x: (x['group'], x['name']))
        return items, print_date
    except Exception as e:
        log(f'ERROR inventory: {e}')
        return [], ''


def read_invoices():
    """Read invoice data and compute analytics"""
    target_file = find_latest_file(INVOICES_DIR, ['فاتورة'], INVOICES_FILE)
    if not target_file or not os.path.exists(target_file):
        log(f'WARNING: Invoices file not found in {INVOICES_DIR}')
        return {'items': [], 'top_customers': {}, 'item_stats': {}, 'print_date': ''}

    try:
        wb = safe_load_workbook(target_file, data_only=True, read_only=True)
        ws = wb.active

        print_date = ''
        item_data = defaultdict(lambda: {
            'name': '', 'code': '', 'unit': '',
            'total_out': 0, 'total_in': 0, 'total_revenue': 0,
            'invoice_count': 0, 'prices': [], 'customers': defaultdict(float),
            'daily_out': defaultdict(float), 'daily_customers': defaultdict(lambda: defaultdict(float)), 'transactions': []
        })

        for row in ws.iter_rows(min_row=4, values_only=True):
            if not row[0] or not row[1]:
                continue

            code = str(row[0] or '').strip()
            name = str(row[1] or '').strip()
            
            # Skip excluded items
            if is_excluded(name):
                continue
                
            unit = str(row[2] or '').strip()
            qty_in = safe_float(row[3])
            qty_out = safe_float(row[4])
            price = safe_float(row[5])
            total = safe_float(row[7])
            bayan = str(row[10] or '').strip()
            date_str = str(row[11] or '').strip()
            invoice_num = str(row[12] or '').strip()

            # Parse date
            try:
                if ' ' in date_str:
                    date_part = date_str.split(' ')[0]
                else:
                    date_part = date_str
                dt = datetime.strptime(date_part, '%Y-%m-%d')
            except:
                dt = None

            item = item_data[code]
            item['name'] = name
            item['code'] = code
            item['unit'] = unit
            item['total_out'] += qty_out
            item['total_in'] += qty_in
            item['total_revenue'] += total
            item['invoice_count'] += 1
            if price > 0:
                item['prices'].append(price)

            # Extract customer name from bayan
            # Format: "فاتورة بيع للعميل <name> رقم <num>"
            customer = ''
            if 'للعميل' in bayan:
                try:
                    after = bayan.split('للعميل')[1]
                    if 'رقم' in after:
                        customer = after.split('رقم')[0].strip()
                    else:
                        customer = after.strip()
                except:
                    customer = bayan
            elif bayan:
                customer = bayan

            if customer:
                item['customers'][customer] += qty_out

            if dt:
                date_key = dt.strftime('%Y-%m-%d')
                item['daily_out'][date_key] += qty_out
                if customer:
                    item['daily_customers'][date_key][customer] += qty_out

        wb.close()

        # Build results
        items_list = []
        for code, data in item_data.items():
            avg_price = 0
            if data['prices']:
                avg_price = sum(data['prices']) / len(data['prices'])

            # Latest price (last transaction)
            latest_price = data['prices'][-1] if data['prices'] else 0

            # Top customers for this item
            top_customers = sorted(data['customers'].items(), key=lambda x: x[1], reverse=True)[:5]

            items_list.append({
                'code': code,
                'name': data['name'],
                'unit': data['unit'],
                'total_out': round(data['total_out'], 3),
                'total_in': round(data['total_in'], 3),
                'total_revenue': round(data['total_revenue'], 2),
                'invoice_count': data['invoice_count'],
                'avg_price': round(avg_price, 3),
                'latest_price': round(latest_price, 3),
                'top_customers': [{'name': c, 'qty': round(q, 3)} for c, q in top_customers],
                'daily_out': dict(data['daily_out']),
                'daily_customers': {d: dict(custs) for d, custs in data['daily_customers'].items()},
            })

        # Sort by total_out descending
        items_list.sort(key=lambda x: x['total_out'], reverse=True)

        # Overall top customers
        all_customers = defaultdict(float)
        for item in items_list:
            for c in item['top_customers']:
                all_customers[c['name']] += c['qty']
        top_overall = sorted(all_customers.items(), key=lambda x: x[1], reverse=True)[:10]

        # Daily totals across all items
        daily_totals = defaultdict(float)
        for item in items_list:
            for date, qty in item['daily_out'].items():
                daily_totals[date] += qty
        daily_sorted = sorted(daily_totals.items())

        return {
            'items': items_list,
            'top_customers': [{'name': c, 'qty': round(q, 3)} for c, q in top_overall],
            'daily_totals': [{'date': d, 'qty': round(q, 3)} for d, q in daily_sorted],
            'print_date': print_date,
            'total_transactions': sum(i['invoice_count'] for i in items_list),
        }

    except Exception as e:
        log(f'ERROR invoices: {e}')
        log(traceback.format_exc())
        return {'items': [], 'top_customers': [], 'daily_totals': [], 'print_date': '', 'total_transactions': 0}


def read_reconciliations(inv_items=None):
    """Read inventory reconciliation movements and calculate shrinkage/surplus statistics"""
    target_file = find_latest_file(INVENTORY_DIR, ['حركة', 'تفصيلي'], RECONCILIATION_FILES[0])

    if not target_file or not os.path.exists(target_file):
        log('WARNING: Reconciliation file not found in ' + str(INVENTORY_DIR))
        if os.path.exists(CACHE_REC_FILE):
            try:
                with open(CACHE_REC_FILE, 'r', encoding='utf-8') as cf:
                    cached = json.load(cf)
                if 'data' in cached:
                    log('Falling back to cached reconciliation data')
                    return cached['data']
            except Exception:
                pass
        return {'items': [], 'summary': {}}

    try:
        mtime = os.path.getmtime(target_file)
        if os.path.exists(CACHE_REC_FILE):
            try:
                with open(CACHE_REC_FILE, 'r', encoding='utf-8') as cf:
                    cached = json.load(cf)
                if cached.get('mtime') == mtime and 'data' in cached:
                    cdata = cached['data']
                    if inv_items:
                        inv_map = {str(i.get('code', '')).strip(): i for i in (inv_items or [])}
                        for it in cdata.get('items', []):
                            inv_entry = inv_map.get(str(it['code']).strip())
                            it['current_stock_balance'] = inv_entry['balance'] if inv_entry else None
                    return cdata
            except Exception:
                pass

        mtime_str = datetime.fromtimestamp(mtime).strftime('%Y-%m-%d %H:%M:%S')
        log(f'Reading reconciliations from: {os.path.basename(target_file)} (modified: {mtime_str})')
        wb = safe_load_workbook(target_file, data_only=True, read_only=True)
        ws = wb.active

        items_movements = defaultdict(list)
        cotton_seed_wared = []

        for row in ws.iter_rows(min_row=4, values_only=True):
            code = str(row[0] or '').strip()
            if not code:
                continue

            name = str(row[1] or '').strip()
            if is_excluded(name):
                continue

            unit = str(row[2] or '').strip()
            wared = safe_float(row[3])
            sader = safe_float(row[4])
            talef = safe_float(row[5])
            balance = safe_float(row[6])
            bayan = str(row[7] or '').strip()
            date_str = str(row[8] or '').strip()
            user = str(row[9] or '').strip()

            mov = {
                'code': code,
                'name': name,
                'unit': unit,
                'wared': wared,
                'sader': sader,
                'talef': talef,
                'balance': balance,
                'bayan': bayan,
                'date': date_str,
                'user': user,
                'is_tasweya': 'تسوية' in bayan
            }
            items_movements[code].append(mov)
            if code == '1115094' and wared > 0: # بذرة قطن
                cotton_seed_wared.append((date_str, wared))

        wb.close()

        result_items = []
        total_tasweyat_count = 0
        shortage_items_count = 0
        surplus_items_count = 0
        total_shortage_feed_kg = 0.0
        total_shortage_pieces = 0.0
        total_surplus_feed_kg = 0.0
        total_surplus_pieces = 0.0

        for code, movs in items_movements.items():
            tasweyat = [m for m in movs if m['is_tasweya']]
            if not tasweyat:
                continue

            name = movs[0]['name']
            unit = movs[0]['unit']
            is_kilo = (unit == 'كيلو')
            tasweyat.sort(key=lambda x: x['date'])
            total_tasweyat_count += len(tasweyat)

            # Build history
            history = []
            for idx, t in enumerate(tasweyat):
                p_date = tasweyat[idx-1]['date'] if idx > 0 else '2025-12-31'
                t_date = t['date']

                c_movs = [m for m in movs if p_date < m['date'] <= t_date and not m['is_tasweya']]
                if code == '6950740': # كسب النور uses بذرة قطن
                    c_wared = sum(w for d, w in cotton_seed_wared if p_date < d <= t_date)
                else:
                    c_wared = sum(m['wared'] for m in c_movs)
                c_sader = sum(m['sader'] for m in c_movs)

                diff = t['wared'] - t['sader']
                status = 'زيادة' if diff > 0 else 'عجز' if diff < 0 else 'مطابق'
                abs_diff = round(abs(diff), 3)

                kg_per_ton = round((abs_diff / (c_wared / 1000)), 2) if (is_kilo and c_wared > 0 and diff < 0) else 0
                shortage_pct = round((abs_diff / c_wared * 100), 2) if (c_wared > 0 and diff < 0) else 0

                history.append({
                    'date': t_date,
                    'bayan': t['bayan'],
                    'user': t['user'],
                    'status': status,
                    'diff': abs_diff,
                    'raw_diff': diff,
                    'balance': round(t['balance'], 3),
                    'prev_date': p_date,
                    'cycle_wared': round(c_wared, 3),
                    'cycle_sader': round(c_sader, 3),
                    'kg_per_ton': kg_per_ton,
                    'shortage_pct': shortage_pct
                })

            last_h = history[-1]
            last_date = last_h['date']

            # Post last tasweya movements
            post_movs = [m for m in movs if m['date'] > last_date and not m['is_tasweya']]
            if code == '6950740':
                post_wared = sum(w for d, w in cotton_seed_wared if d > last_date)
            else:
                post_wared = sum(m['wared'] for m in post_movs)
            post_sader = sum(m['sader'] for m in post_movs)

            if last_h['raw_diff'] < 0:
                shortage_items_count += 1
                if is_kilo:
                    total_shortage_feed_kg += last_h['diff']
                else:
                    total_shortage_pieces += last_h['diff']
            elif last_h['raw_diff'] > 0:
                surplus_items_count += 1
                if is_kilo:
                    total_surplus_feed_kg += last_h['diff']
                else:
                    total_surplus_pieces += last_h['diff']

            is_cotton_based = (code == '6950740')
            category = 'خامات وأعلاف' if is_kilo else 'مستلزمات تعبئة وتغليف'

            result_items.append({
                'code': code,
                'name': name,
                'unit': unit,
                'category': category,
                'is_feed': is_kilo,
                'total_tasweyat': len(tasweyat),
                'last_tasweya_date': last_h['date'],
                'status': last_h['status'],
                'diff': last_h['diff'],
                'raw_diff': last_h['raw_diff'],
                'balance_after': last_h['balance'],
                'cycle_prev_date': last_h['prev_date'],
                'cycle_wared': last_h['cycle_wared'],
                'cycle_sader': last_h['cycle_sader'],
                'kg_per_ton': last_h['kg_per_ton'],
                'shortage_pct': last_h['shortage_pct'],
                'post_wared': round(post_wared, 3),
                'post_sader': round(post_sader, 3),
                'is_cotton_based': is_cotton_based,
                'history': list(reversed(history))
            })

        # Attach current stock balance from تقارير ارصدة المخازن
        inv_map = {str(i.get('code', '')).strip(): i for i in (inv_items or [])}
        for it in result_items:
            inv_entry = inv_map.get(str(it['code']).strip())
            it['current_stock_balance'] = inv_entry['balance'] if inv_entry else None

        # Feeds/commodities first, shortages first, then highest kg_per_ton, then diff
        result_items.sort(key=lambda x: (
            0 if x['is_feed'] else 1,
            0 if x['raw_diff'] < 0 else 1,
            -x['kg_per_ton'] if x['is_feed'] else -x['shortage_pct'],
            -x['diff']
        ))

        feed_shortages = [i['kg_per_ton'] for i in result_items if i['is_feed'] and i['kg_per_ton'] > 0]
        avg_kg_per_ton = round(sum(feed_shortages) / len(feed_shortages), 2) if feed_shortages else 0

        reconciliation_data = {
            'items': result_items,
            'summary': {
                'total_items_with_tasweya': len(result_items),
                'total_tasweyat_count': total_tasweyat_count,
                'shortage_items_count': shortage_items_count,
                'surplus_items_count': surplus_items_count,
                'total_shortage_feed_kg': round(total_shortage_feed_kg, 2),
                'total_shortage_pieces': int(total_shortage_pieces),
                'total_surplus_feed_kg': round(total_surplus_feed_kg, 2),
                'total_surplus_pieces': int(total_surplus_pieces),
                'total_shortage_kg': round(total_shortage_feed_kg, 2),
                'total_surplus_kg': round(total_surplus_feed_kg, 2),
                'avg_kg_per_ton': avg_kg_per_ton
            }
        }

        try:
            with open(CACHE_REC_FILE, 'w', encoding='utf-8') as cf:
                json.dump({'mtime': mtime, 'data': reconciliation_data}, cf, ensure_ascii=False, indent=2)
        except Exception as e:
            log(f'ERROR saving reconciliation cache: {e}')

        return reconciliation_data

    except Exception as e:
        log(f'ERROR reconciliations: {e}')
        log(traceback.format_exc())
        if os.path.exists(CACHE_REC_FILE):
            try:
                with open(CACHE_REC_FILE, 'r', encoding='utf-8') as cf:
                    cached = json.load(cf)
                if 'data' in cached:
                    log('Falling back to cached reconciliation data after error')
                    return cached['data']
            except Exception:
                pass
        return {'items': [], 'summary': {}}


def generate_data():
    log('Reading data...')
    inv_items, inv_print_date = read_inventory()
    invoices = read_invoices()
    reconciliations = read_reconciliations(inv_items)

    total_items = len(inv_items)
    total_balance = sum(i['balance'] for i in inv_items)
    low_stock = sum(1 for i in inv_items if 0 < i['balance'] < 50)
    out_of_stock = sum(1 for i in inv_items if i['balance'] <= 0)

    groups = {}
    for item in inv_items:
        g = item['group'] or 'غير مصنف'
        if g not in groups:
            groups[g] = {'count': 0, 'balance': 0}
        groups[g]['count'] += 1
        groups[g]['balance'] += item['balance']

    data = {
        'generated_at': datetime.now().isoformat(),
        'print_date': inv_print_date,
        'invoices_print_date': invoices.get('print_date', ''),
        'store_name': 'النور لتجارة الأعلاف',
        'inventory': inv_items,
        'invoices': invoices,
        'reconciliations': reconciliations,
        'stats': {
            'total_items': total_items,
            'total_balance': round(total_balance, 3),
            'low_stock': low_stock,
            'out_of_stock': out_of_stock,
            'groups': groups,
            'total_transactions': invoices.get('total_transactions', 0),
        }
    }
    return data


def embed_data_into_html(data):
    """Embed full data into HTML"""
    if not os.path.exists(HTML_FILE):
        return False
    try:
        with open(HTML_FILE, 'r', encoding='utf-8') as f:
            html = f.read()
        json_str = json.dumps(data, ensure_ascii=False)
        marker_start = '/*__EMBEDDED_DATA__*/'
        marker_end = '/*__END_EMBEDDED_DATA__*/'
        start_idx = html.find(marker_start)
        end_idx = html.find(marker_end)
        if start_idx != -1 and end_idx != -1:
            new_block = f'{marker_start}\nvar EMBEDDED_DATA = {json_str};\n{marker_end}'
            html = html[:start_idx] + new_block + html[end_idx + len(marker_end):]
            with open(HTML_FILE, 'w', encoding='utf-8') as f:
                f.write(html)
            return True
        return False
    except Exception as e:
        log(f'ERROR embed: {e}')
        return False


CREATE_NO_WINDOW = 0x08000000 if sys.platform == 'win32' else 0

def git_push():
    try:
        os.chdir(REPO_DIR)
        startupinfo = None
        if sys.platform == 'win32':
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            startupinfo.wShowWindow = subprocess.SW_HIDE
        subprocess.run(['git', 'add', '-A'], capture_output=True, text=True, encoding='utf-8', creationflags=CREATE_NO_WINDOW, startupinfo=startupinfo)
        result = subprocess.run(['git', 'diff', '--cached', '--quiet'], capture_output=True, text=True, creationflags=CREATE_NO_WINDOW, startupinfo=startupinfo)
        if result.returncode == 0:
            log('No changes to commit')
            return True
        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        subprocess.run(['git', 'commit', '-m', f'Auto-update {timestamp}'], capture_output=True, text=True, encoding='utf-8', creationflags=CREATE_NO_WINDOW, startupinfo=startupinfo)
        result = subprocess.run(['git', 'push'], capture_output=True, text=True, encoding='utf-8', creationflags=CREATE_NO_WINDOW, startupinfo=startupinfo)
        if result.returncode != 0:
            log(f'git push: {result.stderr}')
            return False
        log('Pushed to GitHub')
        return True
    except Exception as e:
        log(f'ERROR git: {e}')
        return False


def run_once():
    try:
        data = generate_data()
        with open(JSON_FILE, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        embed_data_into_html(data)
        git_push()
        rec_count = len(data.get("reconciliations", {}).get("items", []))
        log(f'Done. Inventory: {data["stats"]["total_items"]} items (Balance: {data["stats"]["total_balance"]}), Reconciliations: {rec_count} items')
        return True
    except Exception as e:
        log(f'FATAL: {e}')
        log(traceback.format_exc())
        return False


def run_loop(interval_seconds=30):
    log(f'=== Watcher Started (Interval: {interval_seconds}s) ===')
    log(f'Repo: {REPO_DIR}')
    last_mtimes = {}

    while True:
        try:
            cur_files = [
                find_latest_file(INVENTORY_DIR, ['ارصدة', 'المخازن'], INVENTORY_FILE),
                find_latest_file(INVOICES_DIR, ['فاتورة'], INVOICES_FILE),
                find_latest_file(INVENTORY_DIR, ['حركة', 'تفصيلي'], RECONCILIATION_FILES[0])
            ]
            cur_mtimes = {}
            for fp in cur_files:
                if fp and os.path.exists(fp):
                    cur_mtimes[fp] = os.path.getmtime(fp)

            if cur_mtimes != last_mtimes:
                log(f'Detected file change or initial run: {list(cur_mtimes.keys())}')
                run_once()
                last_mtimes = cur_mtimes
        except Exception as e:
            log(f'Error in watch loop: {e}')
            log(traceback.format_exc())

        try:
            time.sleep(interval_seconds)
        except KeyboardInterrupt:
            break
    log('=== Stopped ===')


if __name__ == '__main__':
    if len(sys.argv) > 1 and sys.argv[1] == '--once':
        sys.exit(0 if run_once() else 1)
    else:
        run_loop(30)
