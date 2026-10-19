"""Recompute answers.json from the raw AK Furs exports.

Usage: python solve.py data/ > answers.json
Answers go to stdout; sanity counts go to stderr.
"""
import csv
import datetime as dt
import glob
import json
import os
import re
import sys
from collections import Counter

MONTHS = {'янв': 1, 'фев': 2, 'мар': 3, 'апр': 4, 'мая': 5, 'май': 5, 'июн': 6,
          'июл': 7, 'авг': 8, 'сен': 9, 'окт': 10, 'ноя': 11, 'дек': 12}


def log(*a):
    print(*a, file=sys.stderr)


def find(data_dir, prefix):
    hits = sorted(glob.glob(os.path.join(data_dir, prefix + '*')))
    if not hits:
        sys.exit(f'no file {prefix}* in {data_dir}')
    return hits[0]


def read_text(path):
    raw = open(path, 'rb').read()
    try:
        return raw.decode('utf-8-sig')
    except UnicodeDecodeError:
        return raw.decode('cp1251')  # bank export is Windows-1251


def order_no(s):
    """A-00075, А-75 (Cyrillic), a-75, A75 -> 75."""
    m = re.search(r'\d+', s)
    return int(m.group()) if m else None


def money(s):
    """'2,48 млн', '245,000.00', '140000,00', '2 200 000', '74.000' -> float."""
    s = s.strip().replace('\xa0', ' ').replace('₽', '').replace('руб.', '').replace('руб', '').strip()
    mult = 1
    if 'млн' in s:
        mult, s = 1_000_000, s.replace('млн', '').strip()
    s = s.replace(' ', '')
    if re.fullmatch(r'-?\d{1,3}(,\d{3})+(\.\d+)?', s):      # 245,000.00
        s = s.replace(',', '')
    elif re.fullmatch(r'-?\d{1,3}(\.\d{3})+', s):           # 74.000 (thousands dot)
        s = s.replace('.', '')
    else:                                                   # 140000,00 / 2,48
        s = s.replace(',', '.')
    return float(s) * mult


def parse_date(s):
    s = s.strip()
    if re.fullmatch(r'\d{5}', s):                           # Excel serial
        return dt.date(1899, 12, 30) + dt.timedelta(days=int(s))
    if re.fullmatch(r'\d{4}-\d{2}-\d{2}', s):
        return dt.date.fromisoformat(s)
    m = re.fullmatch(r'(\d{1,2})[./](\d{1,2})[./](\d{4})', s)  # day first: day values go up to 31
    if m:
        return dt.date(int(m[3]), int(m[2]), int(m[1]))
    m = re.fullmatch(r'(\d{1,2})\s+([а-я]+)\.?\s+(\d{4})', s.lower())
    if m and m[2][:3] in MONTHS:
        return dt.date(int(m[3]), MONTHS[m[2][:3]], int(m[1]))
    raise ValueError(f'unknown date format: {s!r}')


def status(s):
    s = s.strip().lower().replace('ё', 'е')
    if s in ('оплачен', 'оплачено', 'paid'):
        return 'Оплачен'
    if s in ('отменен', 'отменено', 'cancelled', 'canceled'):
        return 'Отменён'
    if s in ('возврат', 'refund'):
        return 'Возврат'
    return 'В работе'


def load_lines(path):
    """Dedup on (order number, line number), keeping the latest 'Выгружено'."""
    rows = list(csv.DictReader(read_text(path).splitlines(), delimiter=';'))
    latest = {}
    for r in rows:
        key = (order_no(r['Номер заказа']), int(r['№ строки']))
        if key not in latest or r['Выгружено'].strip() > latest[key]['Выгружено'].strip():
            latest[key] = r
    lines = []
    for (order, line), r in latest.items():
        lines.append({
            'order': order, 'line': line,
            'created': parse_date(r['Дата создания']),
            'manager_raw': r['Менеджер'],
            'sku': r['Артикул'].strip(),
            'qty': int(r['Кол-во']),
            'price': money(r['Цена за ед.']),
            'discount_raw': r['Скидка'].strip(),
            'status': status(r['Статус']),
        })
    return rows, lines


def main(data_dir):
    raw_rows, lines = load_lines(find(data_dir, '01_'))
    orders = {l['order'] for l in lines}
    log(f'Q1: raw rows {len(raw_rows)}, unique lines {len(lines)}, '
        f'dropped duplicates {len(raw_rows) - len(lines)}, unique orders {len(orders)} '
        f'(numbers {min(orders)}..{max(orders)})')
    log('Q1: statuses per line', dict(Counter(l['status'] for l in lines)))

    answers = json.load(open(find(data_dir, 'answers_template')))
    answers.update({
        'q1_unique_orders': len(orders),
        'q1_unique_order_lines': len(lines),
        'run_command': 'python solve.py data/ > answers.json',
    })
    json.dump(answers, sys.stdout, ensure_ascii=False, indent=1)
    print()


if __name__ == '__main__':
    main(sys.argv[1] if len(sys.argv) > 1 else 'data/')
