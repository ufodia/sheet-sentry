# Workbook fixture provenance

`q3-budget-openpyxl.xlsx` is a no-edit openpyxl 3.1.5 load/save of `examples/q3-budget.xlsx`; openpyxl removes formula caches. The `budget-before`, `insert-*`, and `delete-*` workbooks are independent openpyxl 3.1.5 outputs from `generate-structure-fixtures.py`. They include correct and incorrect totals after insertion and deletion at the start, middle, and end of a five-item list. Tests consume the checked-in files without requiring Python.
