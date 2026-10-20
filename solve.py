"""Recompute answers.json from the raw AK Furs exports.

Usage: python solve.py data/ > answers.json
Answers go to stdout; sanity counts go to stderr.
"""
import csv
import datetime as dt
import glob
import hashlib
import json
import os
import re
import subprocess
import sys
import time
import zipfile
from collections import Counter
from concurrent.futures import ThreadPoolExecutor

import openpyxl

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


# Order reference in a payment purpose: "A877", "А-00391", "зак. 528", "заказ 775", "№ 00991".
# The letter must not follow another letter, so "Оплата 75" is not read as order 75.
ORDER_REF = re.compile(r'(?<![^\W\d_])(?:[AАaа]-?|№\s*|зак(?:аз[а-я]*)?\.?\s*)0*(\d+)', re.I)


def load_bank(path):
    rows = list(csv.reader(read_text(path).splitlines(), delimiter=';'))
    head = next(i for i, r in enumerate(rows) if r and r[0].strip() == 'Дата операции')
    cols = [c.strip() for c in rows[head]]
    return [dict(zip(cols, r)) for r in rows[head + 1:] if any(c.strip() for c in r)]


def write_check(name, header, rows):
    """Full lists for checking by eye go to checks/, the console gets a summary."""
    os.makedirs('checks', exist_ok=True)
    with open(os.path.join('checks', name), 'w', encoding='utf-8-sig', newline='') as f:
        w = csv.writer(f, delimiter=';')
        w.writerow(header)
        w.writerows(rows)
    return os.path.join('checks', name)


def rub(x):
    x = round(x, 2)
    return int(x) if x == int(x) else x


def discount(l):
    """'N%' is a percent of Цена × Кол-во; a plain number is roubles off the whole line."""
    gross, d = l['price'] * l['qty'], l['discount_raw']
    return gross * money(d[:-1]) / 100 if d.endswith('%') else money(d) if d else 0.0


def q2(lines, bank, notes):
    orders = {l['order'] for l in lines}
    statuses = {}
    for l in lines:
        statuses.setdefault(l['order'], []).append(l['status'])
    paid_any = {o for o, st in statuses.items() if 'Оплачен' in st}  # any non-cancelled line Оплачен
    paid_all = {o for o, st in statuses.items() if set(st) == {'Оплачен'}}
    mixed = sum(1 for st in statuses.values() if len(set(st)) > 1)
    log(f'Q2: orders with mixed line statuses {mixed}; paid orders: any-line rule {len(paid_any)}, '
        f'all-lines rule {len(paid_all)}, differ by {len(paid_any ^ paid_all)}')

    incoming = [b for b in bank if b['Тип'].strip() == 'Поступление']
    log(f'Q2: bank rows {len(bank)}, incoming {len(incoming)}, skipped outgoing {len(bank) - len(incoming)}')
    paid_orders, unident, unread = set(), [], []
    for b in incoming:
        purpose = b['Назначение платежа']
        m = ORDER_REF.search(purpose)
        if m and int(m.group(1)) in orders:
            paid_orders.add(int(m.group(1)))
        else:
            unident.append(b)
        rest = purpose[:m.start()] + ' ' + purpose[m.end():] if m else purpose
        other = re.findall(r'[+\d][\d\s().:-]*\d', rest)
        if other:
            unread.append([b['Дата операции'], b['Сумма'], purpose, m.group(1) if m else '', ' | '.join(o.strip() for o in other)])

    def answer(paid):
        return sorted(o for o in paid if o not in paid_orders)

    q2a, q2a_all = answer(paid_any), answer(paid_all)
    total = sum(money(b['Сумма']) for b in unident)
    log(f'Q2a: paid orders without any incoming payment: {len(q2a)} -> {q2a}')
    notes.append('Q2a: заказ «Оплачен», если хотя бы одна его неотменённая строка «Оплачен».')
    if q2a != q2a_all:
        log(f'Q2a: with the all-lines rule the answer would be {len(q2a_all)} -> {q2a_all}')
        notes.append(f'Q2a: если требовать «Оплачен» у всех строк, ответ другой: {", ".join(map(str, q2a_all))}.')
    else:
        notes.append('Q2a: правило «все строки Оплачен» даёт тот же ответ.')
    total_int = round(total)
    if abs(total - total_int) > 0.005:
        log(f'Q2b WARNING: sum {total:.2f} has kopecks, rounded to {total_int} for the int answer')
    log(f'Q2b: unidentified incoming payments {len(unident)}, sum {total_int}')
    totals, created = Counter(), {}
    for l in lines:
        totals[l['order']] += l['price'] * l['qty'] - discount(l)
        created[l['order']] = l['created']
    for b in unident:
        same = [o for o in q2a if abs(totals[o] - money(b['Сумма'])) < 1]  # amount equals a Q2a order total
        log(f'   {b["Дата операции"]}  {b["Сумма"]:>14}  {b["Назначение платежа"]}'
            + (f'   <-- amount equals total of Q2a order {same}' if same else ''))
        for o in same:
            pay_day = parse_date(b['Дата операции'])
            notes.append(f'Q2: неопознанный платёж {b["Сумма"]} ₽ от {b["Дата операции"]} («{b["Назначение платежа"]}») '
                         f'по сумме совпадает с заказом {o} из Q2a, но по правилу ТЗ остаётся неопознанным'
                         + (f', тем более заказ создан позже, {created[o]:%d.%m.%Y}.' if created[o] > pay_day else '.'))
    by_shape = Counter(re.sub(r'\d+', '9', r[2]) for r in unread)
    path = write_check('q2_unread_numbers.csv', ['дата', 'сумма', 'назначение', 'прочитан номер заказа', 'прочие числа'], unread)
    log(f'Q2: {len(unread)} purposes contain numbers not read as an order number (full list: {path}):')
    for shape, n in by_shape.most_common():
        log(f'   {n:4}  {shape}')
    notes += ['Q2: платёж опознан, если в назначении есть номер существующего заказа, в том числе отменённого.',
              'Q2: номер в «Счёт № N» и «по счету №N» считаю номером заказа: суммы таких платежей совпадают с суммами заказов.',
              'Q2: телефоны и даты в назначении платежа не считаю номерами заказов.',
              'Q2: списания (возвраты) не считаю поступлениями.']
    return {
        'q2a_paid_orders_without_payment': q2a,
        'q2b_unidentified_payments_count': len(unident),
        'q2b_unidentified_payments_sum_rub': total_int,
    }


def sheet(wb, name, first_col):
    """Rows of a sheet as dicts, starting from the row whose first cell is first_col."""
    rows = list(wb[name].iter_rows(values_only=True))
    head = next(i for i, r in enumerate(rows) if r and str(r[0]).strip() == first_col)
    cols = [str(c).strip() if c is not None else '' for c in rows[head]]
    return [dict(zip(cols, r)) for r in rows[head + 1:] if r and r[0] is not None]


def as_date(v):
    if v is None or v == '':
        return None
    return v.date() if isinstance(v, dt.datetime) else v if isinstance(v, dt.date) else parse_date(str(v))


def name_key(s):
    return re.sub(r'[\s.]', '', s.lower().replace('ё', 'е'))


def q3(lines, wb, notes):
    renames = {r['Старый артикул'].strip(): r['Новый артикул'].strip() for r in sheet(wb, 'Переименования', 'Старый артикул')}
    prices = {}
    for r in sheet(wb, 'Прайс_история', 'Артикул'):
        prices.setdefault(r['Артикул'].strip(), []).append(
            (as_date(r['Действует с']), as_date(r['Действует по']), float(r['Минимально допустимая цена, ₽'])))
    managers, fired = {}, {}
    for r in sheet(wb, 'Менеджеры', 'Каноническое имя'):
        if as_date(r.get('Дата увольнения')):
            fired[r['Каноническое имя'].strip()] = as_date(r['Дата увольнения'])
        for v in [r['Каноническое имя']] + str(r['Варианты написания в CRM'] or '').split('|'):
            managers[name_key(v)] = r['Каноническое имя'].strip()

    created = {}
    for l in lines:
        created.setdefault(l['order'], set()).add(l['created'])
    odd = {o: sorted(d) for o, d in created.items() if len(d) > 1}
    log(f'Q3: orders whose lines have different creation dates: {len(odd)} {odd if odd else ""}')
    created = {o: min(d) for o, d in created.items()}  # shortcut: earliest date if lines disagree (none do today)

    unknown_mgr, no_period, viol = Counter(), Counter(), []
    renamed = 0
    for l in lines:
        if l['status'] == 'Отменён':
            continue
        sku = l['sku']
        while sku in renames:
            sku = renames[sku]
        renamed += sku != l['sku']
        day = created[l['order']]
        period = next(((a, b, p) for a, b, p in prices.get(sku, []) if a <= day and (b is None or day <= b)), None)
        if period is None:
            no_period[sku] += 1
            continue
        d = l['discount_raw']
        disc = discount(l)
        unit = (l['price'] * l['qty'] - disc) / l['qty']
        if unit < period[2] - 1e-6:
            mgr = managers.get(name_key(l['manager_raw']))
            if mgr is None:
                unknown_mgr[l['manager_raw']] += 1
                mgr = l['manager_raw'].strip()
            viol.append([l['order'], l['line'], day, l['sku'], sku, l['qty'], l['price'], d, rub(disc),
                         round(unit, 2), period[2], period[0], period[1] or '', mgr])
    if no_period:
        log(f'Q3 WARNING: lines with no price period for their SKU/date, skipped: {dict(no_period)}')
    if unknown_mgr:
        log(f'Q3 WARNING: manager spellings not in the reference, kept as is: {dict(unknown_mgr)}')
    counts = Counter(v[-1] for v in viol)
    by_mgr = dict(sorted(counts.items()))
    zero = sorted(set(managers.values()) - set(counts))
    header = ['заказ', 'строка', 'дата заказа', 'артикул в CRM', 'артикул в прайсе', 'кол-во', 'цена за ед.',
              'скидка', 'скидка, руб', 'цена за ед. после скидки', 'мин. цена', 'период с', 'период по', 'менеджер']
    path = write_check('q3_violations.csv', header, viol)
    log(f'Q3: non-cancelled lines checked {sum(l["status"] != "Отменён" for l in lines)}, '
        f'old SKUs mapped to new {renamed}; violations {len(viol)} by manager {by_mgr} (all: {path})')
    for v in viol[:5]:
        log('   ' + '; '.join(f'{h}={x}' for h, x in zip(header, v)))
    notes += ['Q3: старый артикул всегда заменяю на новый, независимо от даты перехода: в «Прайс_история» есть только новые.',
              'Q3: дата заказа одна на заказ; границы периодов цен включительно, пустое «Действует по» — без конца.',
              'Q3: цена за единицу после скидки = (Цена × Кол-во − скидка) / Кол-во; нарушение — строго ниже минимальной.',
              'Q3: исключены только строки «Отменён»; «Возврат» и «В работе» учитываются.',
              ]
    for m, day in sorted(fired.items()):
        before = sum(1 for v in viol if v[-1] == m and v[2] <= day)
        after = sum(1 for v in viol if v[-1] == m and v[2] > day)
        log(f'Q3: {m} left {day}: violations on orders created before/on {before}, after {after}')
        notes.append(f'Q3: нарушения уволившегося менеджера ({m}, {day:%d.%m.%Y}) считаю на него: '
                     f'{before} по заказам до увольнения, {after} — после.')
    notes += [f'Q3: {m} — 0 нарушений, поэтому в разбивку не включена.' for m in zero]
    return {'q3_price_violation_lines_total': len(viol), 'q3_price_violation_lines_by_manager': by_mgr}


def load_chats(path):
    """{chat_id: [(role, text), ...]} read straight from the zip; role is 'manager' or 'client'."""
    chats = {}
    with zipfile.ZipFile(path) as z:
        for name in sorted(z.namelist()):
            if not name.endswith('.txt') or '__MACOSX' in name:
                continue
            text = z.read(name).decode('utf-8-sig')
            head, *body = text.splitlines()
            m = re.search(r'C\d+', head)
            cid = m.group() if m else os.path.splitext(os.path.basename(name))[0]
            msgs = []
            for line in body:
                m = re.match(r'\[[^\]]*\]\s*([^:]+):\s*(.*)', line)
                if m:
                    msgs.append(['manager' if m[1].strip().lower().startswith('менедж') else 'client', m[2]])
                elif line.strip() and msgs:
                    msgs[-1][1] += '\n' + line  # message continued on the next line
            chats[cid] = [tuple(x) for x in msgs]
    return chats


# Regex detector, used to check the LLM and as its fallback.
PRICE_RE = re.compile(r'\d[\d\s.,]*\s*(?:тыс|млн|₽|руб|р\.)|\b\d{1,3}(?:[ .\xa0]\d{3})+\b|\b\d{5,}\b|тысяч|миллион', re.I)
INVITE_RE = re.compile(r'примерк|приезжа|подъех|приглаша|запиш|встреч|заезжа|приходи|жд[её]м вас', re.I)
# Showroom address / opening hours: not an invitation, but notes report how the answer changes if it were.
ADDRESS_RE = re.compile(r'находи(?:тся|мся)|адрес|метро|работаем|часы работы|ежедневно', re.I)


def regex_labels(msgs):
    return {'price': [i for i, (r, t) in enumerate(msgs, 1) if r == 'manager' and PRICE_RE.search(t)],
            'invite': [i for i, (r, t) in enumerate(msgs, 1) if r == 'manager' and INVITE_RE.search(t)]}


LLM_PROMPT = '''Ниже переписка менеджера ателье меховых изделий с клиентом. Сообщения пронумерованы.
Для каждого сообщения МЕНЕДЖЕРА определи две метки.

price — менеджер называет клиенту стоимость изделия в деньгах: цифрами или прописью («53 000 ₽», «2,4 млн»,
«двести тысяч рублей»), включая «от 300 тыс.», начальную, итоговую или скидочную цену, а также когда менеджер
повторяет или подтверждает сумму, названную клиентом («да, 450 тысяч»).
НЕ price: проценты (предоплата 30%), число платежей, время, часы работы, даты, размеры, адрес.

invite — менеджер приглашает клиента приехать: на встречу, примерку, в шоурум или ателье; предлагает записать;
спрашивает, когда клиент сможет приехать.
НЕ invite: только адрес или часы работы без приглашения, рассказ об ателье.

Сообщения клиента не размечай. Верни номера сообщений для каждой метки (пустой список, если нет).

<chat>
{chat}
</chat>'''
LLM_SCHEMA = json.dumps({'type': 'object', 'required': ['price', 'invite'], 'properties': {
    'price': {'type': 'array', 'items': {'type': 'integer'}},
    'invite': {'type': 'array', 'items': {'type': 'integer'}}}})
LLM_CMD = ['claude', '-p', '--model', 'sonnet', '--output-format', 'json', '--json-schema', LLM_SCHEMA,
           '--tools', '', '--strict-mcp-config', '--no-session-persistence',
           '--settings', '{"disableAllHooks": true}',
           '--system-prompt', 'Ты размечаешь сообщения переписки. Отвечай только по JSON-схеме.']
CACHE = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'llm_cache.json')


def llm_labels(msgs):
    chat = '\n'.join(f'{i}. {"Менеджер" if r == "manager" else "Клиент"}: {t}' for i, (r, t) in enumerate(msgs, 1))
    out = subprocess.run(LLM_CMD, input=LLM_PROMPT.format(chat=chat), capture_output=True, text=True, timeout=180)
    if out.returncode:
        raise RuntimeError(f'claude exit {out.returncode}: {(out.stderr or out.stdout)[-300:]}')
    res = json.loads(out.stdout)
    labels = res.get('structured_output') or json.loads(res['result'])
    return {k: sorted({int(i) for i in labels[k]}) for k in ('price', 'invite')}


def all_llm_labels(chats):
    """LLM labels per chat: cached by chat-text hash, else 8 parallel `claude -p` calls, else regex fallback."""
    try:
        cache = json.load(open(CACHE, encoding='utf-8'))
    except FileNotFoundError:
        cache = {}
    key = {cid: hashlib.sha256((LLM_PROMPT + repr(m)).encode()).hexdigest() for cid, m in chats.items()}
    todo = [cid for cid in chats if key[cid] not in cache]
    labels, failed, errors = {cid: cache[key[cid]] for cid in chats if key[cid] in cache}, [], set()
    log(f'Q4: {len(chats)} chats, {len(labels)} from LLM cache, {len(todo)} to send to `claude -p`')
    if todo:
        start = time.time()
        with ThreadPoolExecutor(max_workers=8) as pool:
            futures = {cid: pool.submit(llm_labels, chats[cid]) for cid in todo}
            for cid, f in futures.items():
                try:
                    labels[cid] = cache[key[cid]] = f.result()
                except Exception as e:  # CLI missing, timeout, bad JSON: use regex for this chat
                    failed.append(cid)
                    labels[cid] = regex_labels(chats[cid])
                    if str(e) not in errors:  # same error for every chat when the CLI is missing: say it once
                        errors.add(str(e))
                        log(f'Q4 WARNING: LLM failed for {cid}, using regex detector: {e}')
        log(f'Q4: LLM calls took {time.time() - start:.0f}s for {len(todo)} chats')
        json.dump(cache, open(CACHE, 'w', encoding='utf-8'), indent=0, sort_keys=True)
    if failed:
        log(f'Q4 WARNING: {len(failed)} chats fell back to regex: {failed}')
    return labels, failed


def meeting_before_price(msgs, labels):
    """First manager invitation comes before the first manager price (or no price at all)."""
    manager = {i for i, (r, _) in enumerate(msgs, 1) if r == 'manager'}
    price = min((i for i in labels['price'] if i in manager), default=None)
    invite = min((i for i in labels['invite'] if i in manager), default=None)
    return invite is not None and (price is None or invite < price)


def show_chat(cid, msgs, labels):
    out = [f'--- {cid}: meeting before price = {meeting_before_price(msgs, labels)}']
    for i, (r, t) in enumerate(msgs, 1):
        tags = ''.join(f' [{k}]' for k in ('price', 'invite') if i in labels[k])
        out.append(f'  {i:2}. {"М" if r == "manager" else "К"}: {t}{tags}')
    return '\n'.join(out)


def chat_no(cid):
    return int(re.sub(r'\D', '', cid) or 0)


def q4(chats, notes):
    llm, failed = all_llm_labels(chats)
    flagged = sorted((cid for cid in chats if meeting_before_price(chats[cid], llm[cid])), key=chat_no)
    rx = {cid: regex_labels(m) for cid, m in chats.items()}
    rx_flagged = sorted(cid for cid in chats if meeting_before_price(chats[cid], rx[cid]))
    label_diff = [cid for cid in chats if llm[cid] != rx[cid] and cid not in failed]
    verdict_diff = sorted(set(flagged) ^ set(rx_flagged))
    log(f'Q4: meeting before price (LLM): {len(flagged)} chats; regex detector: {len(rx_flagged)}')
    log(f'Q4: LLM vs regex: labels differ in {len(label_diff)} chats, final answer differs in {len(verdict_diff)}: {verdict_diff}')
    details = []
    for cid in label_diff:
        details.append(show_chat(cid, chats[cid], llm[cid]) + f'\n  regex: {rx[cid]}  llm: {llm[cid]}')
        log(f'   {cid}: llm {llm[cid]}  regex {rx[cid]}' + ('  <-- answer differs' if cid in verdict_diff else ''))
    os.makedirs('checks', exist_ok=True)
    open('checks/q4_disagreements.txt', 'w', encoding='utf-8').write('\n\n'.join(details) + '\n')
    sample = flagged[:3] + [c for c in chats if c not in flagged][:2]
    log('Q4: 5 chats to check by eye (М = менеджер, К = клиент):')
    for cid in sample:
        log(show_chat(cid, chats[cid], llm[cid]))
    no_price = [cid for cid in flagged
                if not any(i in llm[cid]['price'] for i, (r, _) in enumerate(chats[cid], 1) if r == 'manager')]
    notes += ['Q4: LLM (claude -p, Sonnet) только отмечает сообщения менеджера с ценой и с приглашением, порядок проверяет код.',
              'Q4: чат в ответе, если первое приглашение менеджера раньше первой названной им цены.',
              'Q4: если менеджер пригласил, а цену так и не назвал'
              + (f' (как в {", ".join(no_price)})' if no_price else '') + ', чат тоже в ответе.',
              'Q4: чат без приглашения и чат, где цена и приглашение в одном сообщении, нарушением не считаю.',
              'Q4: суммы, названные клиентом, ценой не считаю; подтверждение такой суммы менеджером и «от N» — считаю.',
              f'Q4: проверка независимым regex-детектором: разметка расходится в {len(label_diff)} из '
              f'{len(chats) - len(failed)} чатов, итоговый ответ — в {len(verdict_diff)}'
              + (f' ({", ".join(verdict_diff)}).' if verdict_diff else '.')]
    address = {cid: [i for i, (r, t) in enumerate(m, 1) if r == 'manager' and ADDRESS_RE.search(t)]
               for cid, m in chats.items()}
    alt = [cid for cid in chats if meeting_before_price(
        chats[cid], {'price': llm[cid]['price'], 'invite': llm[cid]['invite'] + address[cid]})]
    log(f'Q4: address/hours messages in {sum(map(bool, address.values()))} chats; '
        f'counted as invites the answer would be {len(alt)} chats instead of {len(flagged)}')
    notes.append(f'Q4: сообщение с адресом и часами работы шоурума приглашением не считаю; '
                 f'если считать — ответ {len(alt)} чатов, а не {len(flagged)}.')
    if failed:
        notes.append(f'Q4: для {len(failed)} чатов LLM была недоступна, использован regex-детектор.')
    return {'q4_chats_meeting_before_price': flagged}


def main(data_dir):
    notes = []
    raw_rows, lines = load_lines(find(data_dir, '01_'))
    orders = {l['order'] for l in lines}
    log(f'Q1: raw rows {len(raw_rows)}, unique lines {len(lines)}, '
        f'dropped duplicates {len(raw_rows) - len(lines)}, unique orders {len(orders)} '
        f'(numbers {min(orders)}..{max(orders)})')
    log('Q1: statuses per line', dict(Counter(l['status'] for l in lines)))
    notes += ['Q1: номер заказа — цифры из поля, буква и ведущие нули отбрасываются.',
              'Q1: из дублей пары «заказ + № строки» беру строку с самым поздним «Выгружено».',
              'Q1: даты вида 11/02/2025 читаю как ДД/ММ, потому что первое число доходит до 31.',
              'Q1: NEW, «Ждём оплату» и прочие статусы считаю «В работе».',
              'Q1: статус беру только из поля «Статус», комментарий («клиент вернул») не учитываю.']

    answers = json.load(open(find(data_dir, 'answers_template')))
    answers.update({
        'q1_unique_orders': len(orders),
        'q1_unique_order_lines': len(lines),
    })
    answers.update(q2(lines, load_bank(find(data_dir, '02_')), notes))
    answers.update(q3(lines, openpyxl.load_workbook(find(data_dir, '03_'), read_only=True), notes))
    answers.update(q4(load_chats(find(data_dir, '04_')), notes))
    answers['run_command'] = 'python solve.py data/ > answers.json'
    answers['notes'] = ' '.join(notes)
    json.dump(answers, sys.stdout, ensure_ascii=False, indent=1)
    print()


if __name__ == '__main__':
    main(sys.argv[1] if len(sys.argv) > 1 else 'data/')
