"""Parser self-checks: python test_solve.py"""
import solve

assert solve.order_no('A-00075') == solve.order_no('А-75') == solve.order_no(' a-75') == solve.order_no('A75') == 75
assert solve.money('2,48 млн') == 2_480_000
assert solve.money('245,000.00') == solve.money('245000,00') == solve.money('245 000') == solve.money('245.000') == 245_000
assert str(solve.parse_date('45752')) == '2025-04-05'
assert str(solve.parse_date('11/02/2025')) == str(solve.parse_date('11.02.2025')) == '2025-02-11'
assert str(solve.parse_date('23 сен 2026')) == '2026-09-23'
assert [solve.status(s) for s in ('Оплачен ', 'PAID', 'отменен', 'REFUND', 'NEW', 'Ждём оплату')] == \
    ['Оплачен', 'Оплачен', 'Отменён', 'Возврат', 'В работе', 'В работе']
print('ok')
