import os
import pandas as pd
import xml.etree.ElementTree as ET

# Caminho da pasta com os XMLs dos CTe
caminho_pasta = 'C:/Users/'

# Caminho de saída para o Excel
caminho_excel = 'C:/Users/'

# Lista para armazenar resultados
resultados = []

# Percorre os arquivos da pasta
for nome_arquivo in sorted(os.listdir(caminho_pasta)):
    if nome_arquivo.lower().endswith('.xml'):
        caminho_xml = os.path.join(caminho_pasta, nome_arquivo)

        chave_acesso = ''
        valor_frete = None

        try:
            tree = ET.parse(caminho_xml)
            root = tree.getroot()

            # Detecta namespace automaticamente
            ns = {'ns': root.tag.split('}')[0].strip('{')}

            # Extrai a chave de acesso
            infCte = root.find(".//ns:infCte", ns)
            if infCte is not None and 'Id' in infCte.attrib:
                chave_acesso = infCte.attrib['Id'].replace('CTe', '')
            else:
                chave_acesso = 'Chave não encontrada'

            # Extrai o valor do frete (vTPrest)
            valor = root.find(".//ns:vTPrest", ns)
            if valor is not None:
                try:
                    valor_frete = float(valor.text)
                except ValueError:
                    valor_frete = None

            resultados.append({
                "Arquivo": nome_arquivo,
                "Chave de Acesso": chave_acesso,
                "Valor do Frete (vTPrest)": valor_frete
            })

        except Exception as e:
            resultados.append({
                "Arquivo": nome_arquivo,
                "Chave de Acesso": f"Erro: {str(e)}",
                "Valor do Frete (vTPrest)": None
            })

# Gera o Excel com os resultados
df = pd.DataFrame(resultados)
df.to_excel(caminho_excel, index=False)

print(f"\n✅ Extração concluída. Resultados salvos em:\n{caminho_excel}")
