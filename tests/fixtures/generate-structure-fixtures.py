"""Recreate the independent structural workbook fixtures with openpyxl 3.1.5."""
from pathlib import Path
from openpyxl import Workbook

out = Path(__file__).resolve().parent
rows = [('Rent', 100), ('Wages', 80), ('Cloud', 60), ('Travel', 40), ('Misc', 20)]

def save(name, items, formula):
    wb = Workbook()
    ws = wb.active
    ws.title = 'Budget'
    ws.append(['Item', 'Amount'])
    for item in items:
        ws.append(item)
    ws.append(['Total', formula])
    wb.save(out / name)

save('budget-before.xlsx', rows, '=SUM(B2:B6)')
for name, index in [('head', 0), ('middle', 2), ('tail', 5)]:
    added = rows.copy()
    added.insert(index, ('Insurance', 120))
    save(f'insert-{name}-correct.xlsx', added, '=SUM(B2:B7)')
    save(f'insert-{name}-wrong.xlsx', added, '=SUM(B3:B7)' if index == 0 else '=SUM(B2:B6)')
    removed = rows.copy()
    removed.pop(min(index, 4))
    save(f'delete-{name}-correct.xlsx', removed, '=SUM(B2:B5)')
    save(f'delete-{name}-wrong.xlsx', removed, '=SUM(B2:B6)')
