import os
import pandas as pd

# Lê a tabela do Excel
df = pd.read_excel(r'C:\Users\psace\Downloads\teste-certificados\certificados.xlsx')

# Diretório onde estão os arquivos PDF
diretorio_pdf = r'C:\Users\psace\Downloads\teste-certificados\1200/'

# Itera pelas linhas da tabela do Excel
for index, row in df.iterrows():
    nome_arquivo_pdf = f'certificadospagina_{index + 1}.pdf'  # Nome do arquivo PDF
    novo_nome_arquivo_pdf = f'{row["Nome"]}.pdf'  # Novo nome com base na tabela do Excel

    caminho_original = os.path.join(diretorio_pdf, nome_arquivo_pdf)
    caminho_novo = os.path.join(diretorio_pdf, novo_nome_arquivo_pdf)

    if not os.path.exists(caminho_novo):
        try:
            os.rename(caminho_original, caminho_novo)
            print(f"Arquivo {nome_arquivo_pdf} renomeado para {novo_nome_arquivo_pdf}")
        except FileNotFoundError:
            print(f"Arquivo {nome_arquivo_pdf} não encontrado.")
    else:
        print(f"Arquivo {novo_nome_arquivo_pdf} já existe. Não foi renomeado.")

print("Concluído!")