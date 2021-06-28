
import pandas as pd
from openpyxl import load_workbook

wb = load_workbook('vendas-combustiveis-m3.xlsx')
ws = wb.worksheets[0]

pivot = ws._pivots[0]

pivot.cache.refreshOnLoad = True

pivot_sheet = pivot.cache.cacheSource.worksheetSource.sheet
pivot_ref = pivot.cache.cacheSource.worksheetSource.ref

# Deste ponto em diante, de posse da tabela-source e das celulas-referencia,
# os dados seriam recuperados e formatados num DataFrame, usando pandas.
# Nao desenvolvi este item.
