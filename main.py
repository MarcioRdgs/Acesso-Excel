import openpyxl

# Caminho para o arquivo Excel
caminho_arquivo_excel = 'C:\Users\mrcio\OneDrive\Documentos\teste acesso python.xlsx'

# Carregando a planilha
planilha = openpyxl.load_workbook(C:\Users\mrcio\OneDrive\Documentos\teste acesso python.xlsx)

# Selecionando a folha de trabalho ativa (pode ser necessário ajustar o nome da folha)
folha = planilha.active

# Acessando e modificando a célula A1
celula_a1 = folha['A1']
celula_a1.value = 'Texto de teste py'

# Salvando as alterações de volta no arquivo
planilha.save(caminho_arquivo_excel)

# Fechando a planilha
planilha.close()
