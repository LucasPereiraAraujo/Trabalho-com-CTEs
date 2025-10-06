import os
import xmlschema
import pandas as pd
import xml.etree.ElementTree as ET

# Caminho da pasta com os XMLs dos CTe
caminho_pasta = 'C:/Users/lucas.araujo/Desktop/Total Fatura'

# Caminho para o arquivo XSD da versão 4.00
caminho_schema = 'C:/Users/lucas.araujo/Desktop/PL_CTe_400_NT2024.002_1.05/procCTe_v4.00.xsd'

# Caminho de saída para o Excel
caminho_excel = 'C:/Users/lucas.araujo/Desktop/Validacao Total menores2.xlsx'

# Carrega o schema
try:
    schema = xmlschema.XMLSchema(caminho_schema)
except Exception as e:
    print(f"Erro ao carregar o schema XSD:\n{e}")
    exit(1)

# Lista para armazenar resultados
resultados = []

# Valida cada XML na pasta
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

            infCte = root.find(".//ns:infCte", ns)
            if infCte is not None and 'Id' in infCte.attrib:
                chave_acesso = infCte.attrib['Id'].replace('CTe', '')

            # Buscar o valor do frete (vTPrest)
            vTPrest = infCte.find(".//ns:vTPrest", ns) if infCte is not None else None
            if vTPrest is not None:
                try:
                    valor_frete = float(vTPrest.text)
                except ValueError:
                    valor_frete = None
        except Exception as e:
            chave_acesso = f"Erro ao ler chave: {str(e)}"

        # Validação do XML
        try:
            if schema.is_valid(caminho_xml):
                resultados.append({
                    "Arquivo": nome_arquivo,
                    "Chave de Acesso": chave_acesso,
                    "Valor do Frete": valor_frete,
                    "Status": "Válido",
                    "Erro": ""
                })
            else:
                try:
                    schema.validate(caminho_xml)
                except xmlschema.XMLSchemaValidationError as e:
                    resultados.append({
                        "Arquivo": nome_arquivo,
                        "Chave de Acesso": chave_acesso,
                        "Valor do Frete": valor_frete,
                        "Status": "Erro",
                        "Erro": str(e)
                    })
        except Exception as e:
            resultados.append({
                "Arquivo": nome_arquivo,
                "Chave de Acesso": chave_acesso,
                "Valor do Frete": valor_frete,
                "Status": "Erro",
                "Erro": f"Erro inesperado: {str(e)}"
            })

# Gera o Excel com os resultados
df = pd.DataFrame(resultados)
df.to_excel(caminho_excel, index=False)

print(f"\n✅ Validação concluída. Resultados salvos em:\n{caminho_excel}")
