import os
import shutil
import pandas as pd

# Pastas
pasta_origem = r'C:\Users\lucas.araujo\Desktop\XML 2Q'
pasta_erro = r'C:\Users\lucas.araujo\Desktop\COM_ERRO_DUPLICADO2Q'
pasta_ok = r'C:\Users\lucas.araujo\Desktop\SEM_ERRO2Q'
os.makedirs(pasta_erro, exist_ok=True)
os.makedirs(pasta_ok, exist_ok=True)

# Listas
arquivos_com_erro = []
arquivos_ok = []

# Verifica cada XML
for nome_arquivo in os.listdir(pasta_origem):
    if nome_arquivo.lower().endswith('.xml'):
        caminho = os.path.join(pasta_origem, nome_arquivo)
        try:
            with open(caminho, 'r', encoding='utf-8') as f:
                conteudo = f.read()

            if conteudo.count('<cteProc') > 1:
                # Com erro
                arquivos_com_erro.append(nome_arquivo)
                shutil.copy(caminho, os.path.join(pasta_erro, nome_arquivo))
            else:
                # Sem erro
                arquivos_ok.append(nome_arquivo)
                shutil.copy(caminho, os.path.join(pasta_ok, nome_arquivo))

        except Exception as e:
            print(f"⚠️ Erro ao processar {nome_arquivo}: {e}")

# Salva planilha dos arquivos com erro
if arquivos_com_erro:
    df_erro = pd.DataFrame(arquivos_com_erro, columns=['Arquivos com erro'])
    caminho_planilha = r'C:\Users\lucas.araujo\Desktop\CTEs_com_erro.xlsx'
    df_erro.to_excel(caminho_planilha, index=False)
    print(f"\n✅ {len(arquivos_com_erro)} arquivos com erro copiados para: {pasta_erro}")
    print(f"📝 Planilha criada: {caminho_planilha}")
else:
    print("✅ Nenhum erro encontrado!")

print(f"✅ {len(arquivos_ok)} arquivos sem erro copiados para: {pasta_ok}")
