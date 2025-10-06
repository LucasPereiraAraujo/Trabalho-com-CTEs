import os
import re
import pandas as pd

# Caminho para a pasta com os arquivos
caminho_pasta = r'C:\Users\lucas.araujo\Downloads\Relação - Faturas e codigos'

resumo = []

# Contadores
total_arquivos = 0
arquivos_processados = 0
arquivos_ignorados = 0
arquivos_com_erro = 0

# Percorre todos os arquivos .xlsx na pasta
for arquivo in os.listdir(caminho_pasta):
    if not arquivo.endswith('.xlsx'):
        continue

    total_arquivos += 1
    caminho_arquivo = os.path.join(caminho_pasta, arquivo)

    # Extrai número do nome do arquivo (pode conter texto extra)
    match_valor = re.match(r"^(\d+[.,]?\d*)", arquivo)
    if not match_valor:
        print(f"[IGNORADO] Nome inválido (sem valor numérico no início): {arquivo}")
        arquivos_ignorados += 1
        continue

    try:
        valor_fatura = float(match_valor.group(1).replace(',', '.'))
    except ValueError:
        print(f"[IGNORADO] Valor inválido no nome do arquivo: {arquivo}")
        arquivos_ignorados += 1
        continue

    try:
        # Lista os nomes das abas
        abas = pd.ExcelFile(caminho_arquivo).sheet_names
        if len(abas) < 2:
            print(f"[ERRO] Arquivo com menos de duas abas: {arquivo}")
            arquivos_com_erro += 1
            continue

        # Lê a segunda aba
        df = pd.read_excel(caminho_arquivo, sheet_name=1)

        # Valida colunas
        colunas_esperadas = ['Data Import.', 'Documento', 'Cod. Fornecedor', 'Chave Ct-e', 'Total']
        if not all(col in df.columns for col in colunas_esperadas):
            print(f"[ERRO] Colunas faltando em: {arquivo}")
            arquivos_com_erro += 1
            continue

        # Conversões
        df['Total'] = pd.to_numeric(df['Total'], errors='coerce').fillna(0)
        df['Data Import.'] = pd.to_datetime(df['Data Import.'], errors='coerce')

        # Agrupa por código
        agrupado = df.groupby('Cod. Fornecedor')['Total'].sum().reset_index()
        data_importacao = df['Data Import.'].min()
        soma_total = df['Total'].sum()

        for _, linha in agrupado.iterrows():
            resumo.append({
                'Nome Fatura': match_valor.group(1),
                'Data Importação': data_importacao.date() if pd.notnull(data_importacao) else None,
                'Código Fornecedor': linha['Cod. Fornecedor'],
                'Total por Código': linha['Total'],
                'Soma Total': soma_total,
                'Valor Esperado': valor_fatura,
                'Diferença': round(soma_total - valor_fatura, 2)
            })

        arquivos_processados += 1

    except Exception as e:
        print(f"[ERRO] Falha ao processar {arquivo}: {e}")
        arquivos_com_erro += 1

# Exporta para Excel
df_resumo = pd.DataFrame(resumo)
df_resumo.to_excel('resumo_faturas.xlsx', index=False)

# Relatório final
print("\n=== RESUMO DE EXECUÇÃO ===")
print(f"Arquivos totais encontrados:         {total_arquivos}")
print(f"Arquivos processados com sucesso:   {arquivos_processados}")
print(f"Arquivos ignorados (nome inválido): {arquivos_ignorados}")
print(f"Arquivos com erro (estrutura):      {arquivos_com_erro}")
print(f"Linhas geradas no Excel final:      {len(df_resumo)}")
print("✅ Arquivo 'resumo_faturas.xlsx' salvo com sucesso.")
