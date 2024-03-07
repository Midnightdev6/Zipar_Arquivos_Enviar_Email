import os
import pandas as pd
from zmq import NULL
import EnviarEmail
import ZiparPastas
import base64

def file_to_base64(file_path):
    with open(file_path, "rb") as file:
        encoded_string = base64.b64encode(file.read()).decode('utf-8')
    return encoded_string

excel= r"C:\\Users\\vitor.moreira\\Documents\\backoffice\\Endereços de pastas - PRESTAÇÃO DE CONTAS BACKOFFICE 1 1 1 (2).xlsx"
df =pd.read_excel(excel,sheet_name="Sheet", header=0)

pasta_raiz = "C:\\Users\\vitor.moreira\\OneDrive - Grupo Trinus Co\\Documentos - PRESTAÇÃO DE CONTAS BACKOFFICE"
ano_mes = "2024\\01 - Janeiro"

for empreendimentos in os.listdir(pasta_raiz):

    # print(empreendimentos)
    caminho_empreendimento = os.path.join(pasta_raiz, empreendimentos)
    # print(caminho_empreendimento)
    if os.path.isdir(caminho_empreendimento):
        ZiparPastas.iterar_dentro_empreendimento(caminho_empreendimento,ano_mes)
    # print(caminho_empreendimento)
        
            ######### Enviar E-mail#################        
        
    for index,row in df.iterrows():  
        # print(row['Empreendimento GI'])     
        if  row['Empreendimento GI'] == empreendimentos:
            print(row['Empreendimento GI'])
            print(empreendimentos)

            ######## Fazer caminho até os zips####### var produto
            arquivo1_path = f"{caminho_empreendimento}\\{ano_mes}\\Contabilidade.zip"
            arquivo2_path = f"{caminho_empreendimento}\\{ano_mes}\\Financeiro.zip"
            arquivo3_path = f"{caminho_empreendimento}\\{ano_mes}\\Fiscal.zip"
            print(arquivo1_path)
            # Converter os arquivos para Base64
            try:
                # Converter os arquivos para Base64
                arquivo1_base64 = file_to_base64(arquivo1_path)
                arquivo2_base64 = file_to_base64(arquivo2_path)
                arquivo3_base64 = file_to_base64(arquivo3_path)
                
                # Verificar o conteúdo dos arquivos base64
                print("Contabilidade.zip:", arquivo1_base64[:100])  
                print("Financeiro.zip:", arquivo2_base64[:100])  
                print("Fiscal.zip:", arquivo3_base64[:100])  

            ######### Provisoriamente, a pasta fiscal foi removida, ela está no codigo comentado  arquivo3_path e arquivo3_base64 #########################################
            except Exception as e:
                print("Erro durante a conversão ou anexação dos arquivos:", e)

            sender_email = "vitor.moreira@trinusco.com.br"
            sender_password = "B2luzinho"
            receiver_email = f"{row['e-mail (separados por vírgula)']}"
            emails_list = receiver_email.split(',')
            subject = empreendimentos
            message = """
                Prezados(as),

                Segue em anexo alguns documentos, referentes ao mês de Janeiro/2024, gerados pelas seguintes áreas do backoffice:

                - Contabilidade - Balancete
                - Financeiro - Extratos bancários, Contas Pagas, Adiantamento a Fornecedor

                Caso surjam dúvidas, por favor, entre em contato conosco:
                - E-mail: centralatendimentob2b@trinusco.com.br
                - WhatsApp: Clique aqui para nos contatar via https://wa.me/556231577462

                Atenciosamente,
                Vitor Ramos Moreira
                """



            attachments = [
            {"filename":arquivo1_path, "data": arquivo1_base64},
            {"filename": arquivo2_path, "data": arquivo2_base64},
            # {"filename": arquivo3_path, "data": arquivo3_base64}
        ]


            # EnviarEmail.send_email(sender_email, sender_password, receiver_email, subject, message, attachments)
            # Itere sobre a lista de e-mails e envie o e-mail para cada destinatário
            for receiver_email in emails_list:
                # Verifique se o e-mail não é 'nan'
                if isinstance(receiver_email, str) and receiver_email != "" and receiver_email.strip().lower() != 'nan':
                # if isinstance(receiver_email, str) and receiver_email.strip().lower() == type(str):
                    EnviarEmail.send_email(sender_email, sender_password, receiver_email.strip(), subject, message, attachments)
                else:
                    print(receiver_email)
                    print(emails_list)
                    continue    
            ###################### Remove os zips criados #######################    
                # if os.path.exists(arquivo1_path):
                #     os.remove(arquivo1_path)
                    
                # if os.path.exists(arquivo2_path):
                #     os.remove(arquivo2_path)   

                # if os.path.exists(arquivo3_path):
                #     os.remove(arquivo3_path)  

            # for receiver_email in emails_list:
            #     EnviarEmail.send_email(sender_email, sender_password, receiver_email.strip(), subject, message, attachments)
