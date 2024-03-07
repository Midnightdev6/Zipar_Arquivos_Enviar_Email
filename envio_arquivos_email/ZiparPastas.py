import os
import zipfile

pastas_desejadas = ["Contabilidade", "Financeiro", "Fiscal"]

def iterar_dentro_empreendimento(caminho,ano_mes):
    caminho_produtos = os.path.join(caminho, ano_mes)
    if os.path.exists(caminho_produtos):
        ## lista as pastas 
        for pasta in os.listdir(caminho_produtos):
           
            produto = os.path.join(caminho_produtos, pasta)
            if os.path.isdir(produto) and pasta in pastas_desejadas:
                zipar_pasta(produto)


def zipar_pasta(pasta):
    nome_zip = f"{pasta}.zip"
    try:
        if os.path.exists(nome_zip):
            os.remove(nome_zip)
        with zipfile.ZipFile(nome_zip, 'w') as zipf:
            for raiz, _, arquivos in os.walk(pasta):
                for arquivo in arquivos:
                    caminho_completo = os.path.join(raiz, arquivo)
                    try:
                        zipf.write(caminho_completo, os.path.relpath(caminho_completo, pasta))
                    except FileNotFoundError as e:
                        print(f"Arquivo não encontrado: {caminho_completo}")
                        continue
    except Exception as e:
        print(f"Erro ao zipar pasta {pasta}: {e}")
        