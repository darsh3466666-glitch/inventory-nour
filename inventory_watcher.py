#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
inventory_watcher.py — مراقب مخزون النور
يقرأ ملف "تقارير  ارصدة المخازن.xlsx" كل دقيقة
يحدّث HTML + JSON + يرفع على GitHub تلقائياً
"""

import os
import sys
import json
import time
import subprocess
import traceback
from datetime import datetime

sys.stdout.reconfigure(encoding='utf-8')
sys.stderr.reconfigure(encoding='utf-8')

INVENTORY_DIR = r'D:\Mostafa Ibrahim\جرد يومي'
INVENTORY_FILE = os.path.join(INVENTORY_DIR, 'تقارير  ارصدة المخازن.xlsx')
REPO_DIR = r'C:\Users\GoldenTech\inventory-nour'
HTML_FILE = os.path.join(REPO_DIR, 'index.html')
JSON_FILE = os.path.join(REPO_DIR, 'inventory_data.json')
LOG_FILE = os.path.join(REPO_DIR, 'watcher.log')

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


def read_inventory():
    if not os.path.exists(INVENTORY_FILE):
        log(f'WARNING: File not found: {INVENTORY_FILE}')
        return [], ''

    try:
        wb = openpyxl.load_workbook(INVENTORY_FILE, data_only=True)
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
            store = str(row[0] or '').strip()
            code = str(row[1] or '').strip()
            name = str(row[2] or '').strip()
            unit = str(row[3] or '').strip()
            balance = safe_float(row[4])
            detailed = str(row[5] or '').strip()
            detained = safe_float(row[6])
            group = str(row[7] or '').strip()
            supplier = str(row[8] or '').strip()
            if not name:
                continue
            items.append({
                'store': store, 'code': code, 'name': name, 'unit': unit,
                'balance': balance, 'detailed': detailed, 'detained': detained,
                'group': group, 'supplier': supplier,
            })

        wb.close()
        items.sort(key=lambda x: (x['group'], x['name']))
        return items, print_date
    except Exception as e:
        log(f'ERROR: {e}')
        log(traceback.format_exc())
        return [], ''


def generate_data():
    log('Reading inventory data...')
    items, print_date = read_inventory()

    total_items = len(items)
    total_balance = sum(i['balance'] for i in items)
    low_stock = sum(1 for i in items if 0 < i['balance'] < 50)
    out_of_stock = sum(1 for i in items if i['balance'] <= 0)

    groups = {}
    for item in items:
        g = item['group'] or 'غير مصنف'
        if g not in groups:
            groups[g] = {'count': 0, 'balance': 0}
        groups[g]['count'] += 1
        groups[g]['balance'] += item['balance']

    data = {
        'generated_at': datetime.now().isoformat(),
        'print_date': print_date,
        'store_name': 'النور لتجارة الأعلاف',
        'inventory': items,
        'stats': {
            'total_items': total_items,
            'total_balance': round(total_balance, 3),
            'low_stock': low_stock,
            'out_of_stock': out_of_stock,
            'groups': groups,
        }
    }
    return data


def embed_data_into_html(data):
    """Embed JSON data into the HTML file"""
    if not os.path.exists(HTML_FILE):
        log(f'HTML file not found: {HTML_FILE}')
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
        else:
            log('WARNING: Embedded data markers not found in HTML')
            return False
    except Exception as e:
        log(f'ERROR embedding HTML: {e}')
        return False


def git_push():
    """Push changes to GitHub"""
    try:
        os.chdir(REPO_DIR)

        # Add all changes
        result = subprocess.run(['git', 'add', '-A'], capture_output=True, text=True, encoding='utf-8')
        if result.returncode != 0:
            log(f'git add failed: {result.stderr}')
            return False

        # Check if there are changes to commit
        result = subprocess.run(['git', 'diff', '--cached', '--quiet'], capture_output=True, text=True)
        if result.returncode == 0:
            log('No changes to commit')
            return True

        # Commit
        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        result = subprocess.run(['git', 'commit', '-m', f'Auto-update inventory {timestamp}'], capture_output=True, text=True, encoding='utf-8')
        if result.returncode != 0:
            log(f'git commit failed: {result.stderr}')
            return False

        # Push
        result = subprocess.run(['git', 'push'], capture_output=True, text=True, encoding='utf-8')
        if result.returncode != 0:
            log(f'git push failed: {result.stderr}')
            return False

        log('Pushed to GitHub successfully')
        return True
    except Exception as e:
        log(f'ERROR git push: {e}')
        return False


def run_once():
    try:
        data = generate_data()

        # Write JSON
        with open(JSON_FILE, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)

        # Embed into HTML
        embed_data_into_html(data)

        # Push to GitHub
        git_push()

        log(f'Update complete. Items: {data["stats"]["total_items"]}, Balance: {data["stats"]["total_balance"]}')
        return True
    except Exception as e:
        log(f'FATAL: {e}')
        log(traceback.format_exc())
        return False


def run_loop(interval_seconds=60):
    log(f'=== Inventory Watcher Started ===')
    log(f'Source: {INVENTORY_FILE}')
    log(f'Repo: {REPO_DIR}')
    log(f'Interval: {interval_seconds}s')

    while True:
        run_once()
        log(f'Next update in {interval_seconds} seconds...')
        try:
            time.sleep(interval_seconds)
        except KeyboardInterrupt:
            log('Stopping...')
            break

    log('=== Inventory Watcher Stopped ===')


if __name__ == '__main__':
    if len(sys.argv) > 1 and sys.argv[1] == '--once':
        success = run_once()
        sys.exit(0 if success else 1)
    else:
        run_loop(60)
