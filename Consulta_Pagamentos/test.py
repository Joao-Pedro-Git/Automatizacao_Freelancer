# app.py para testar se os dados estão indo para planilha 

import openpyxl

# Abrir a planilha de fechamento
planilha_fechamento = openpyxl.Workbook("Sua Planilha de Test !AQUI!")
pagina_fechamento = planilha_fechamento.active

# Adicionar alguns dados de exemplo
pagina_fechamento.append(['Nome', 'Valor', 'CPF', 'Vencimento', 'Status', 'Data Pagamento', 'Método Pagamento'])
pagina_fechamento.append(['Cliente A', 100, '12345678900', '2024-08-01', 'em dia', '2024-07-20', 'Cartão'])
pagina_fechamento.append(['Cliente B', 150, '98765432100', '2024-07-15', 'pendente', '', ''])

# Salvar a planilha de fechamento
planilha_fechamento.save('planilha_report.xlsx')

print("Dados adicionados com sucesso à planilha de fechamento.")
