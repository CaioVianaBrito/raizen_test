
from openpyxl import load_workbook
from openpyxl.worksheet import datavalidation
from openpyxl.worksheet.datavalidation import DataValidation

### opening
wbook = load_workbook('vendas-combustiveis-m3.xlsx')
wsheet = wbook.worksheets[0]

# data-validation object
dv = wsheet.data_validations.dataValidation

# dados da 1a tabela dinamica a ser coletada
data1 = [[cell.value for cell in row] for row in wsheet['$B$53:$W$65']]

# dados da 2a tabela dinamica a ser coletada
data2 = [[cell.value for cell in row] for row in wsheet['$B$132:$J$145']]

# As variaveis 'data1' e 'data2' acima recebem os dados das tabelas dinamicas que devem ser coletadas.
# Estes dados deveriam ser convertidos em um DataFrame para utilizacao conforme o enunciado do exercicio.
# Nao desenvolvi este item.
