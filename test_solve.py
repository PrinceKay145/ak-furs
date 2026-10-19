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

# Q4 regex detector edge cases from the brief and the review
for text, price, invite in [
    ('да, 450 тысяч', True, False),            # manager confirms the client's amount
    ('от 300 тыс.', True, False),
    ('Стоимость этой модели — 208.000 руб.', True, False),
    ('Для вас итоговая — два миллиона двести тысяч рублей', True, False),
    ('Возможна предоплата 30%, остальное при получении.', False, False),
    ('Можем оплатить в 2-3 платежа.', False, False),
    ('Наш шоурум находится у метро Фрунзенская, работаем ежедневно с 10 до 21.', False, False),
    ('Давайте я запишу вас на встречу в шоурум, там и подберём. Завтра в 15:00 удобно?', False, True),
    ('Когда сможете подъехать на примерку?', False, True),
]:
    got = solve.regex_labels([('manager', text)])
    assert (got['price'] == [1], got['invite'] == [1]) == (price, invite), (text, got)
assert solve.regex_labels([('client', 'У конкурентов видела за 600 000 руб.')]) == {'price': [], 'invite': []}
assert solve.meeting_before_price([('manager', 'a'), ('manager', 'b')], {'price': [2], 'invite': [1]})
assert solve.meeting_before_price([('manager', 'a')], {'price': [], 'invite': [1]})        # never named a price
assert not solve.meeting_before_price([('manager', 'a')], {'price': [1], 'invite': [1]})   # same message
assert solve.meeting_before_price([('client', 'a'), ('manager', 'b')], {'price': [1], 'invite': [2]})  # client's price ignored
print('q4 detector ok')
