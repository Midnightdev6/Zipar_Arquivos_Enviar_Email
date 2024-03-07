from genericpath import isdir
import os

caminho_empreendimentos ="C:\\Users\\vitor.moreira\\OneDrive - Grupo Trinus Co\\Documentos - PRESTAÇÃO DE CONTAS BACKOFFICE"
ano_mes = "2024\\01 - Janeiro"

def excluir_zip(caminho1,caminho2):
    for empreendimento in os.listdir(caminho1):
       
        if os.path.isdir(caminho1):
            caminho_empreendimento = os.path.join(caminho1, empreendimento)
            caminho_completo = os.path.join(caminho_empreendimento, caminho2)
            print(os.listdir(caminho_completo))


            for pasta in os.listdir(caminho_completo):
                
                if pasta.endswith(".zip"):
                    print(pasta.endswith(".zip"))
                    print(f"{caminho_completo}\\{pasta}")
                    os.rmdir(f"{caminho_completo}\\{pasta}")

             


    
    


excluir_zip(caminho_empreendimentos, ano_mes)