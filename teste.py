import os
import fitz # biblioteca PyMuPDF
from time import sleep

print("### BEM VINDO ###\n\n### ESTE BOT TEM A FUNÇÃO DE MESCLAR ARQUIVOS PDF LOCALMENTE ###")

print("\n-> OPÇÕES DE DIRETÓRIO: ")
print("1 - INFORMAR O CAMINHO DO DIRETÓRIO")
print("2 - INSERIR OS ARQUIVOS NO DIRETÓRIO 'MERGER' CRIADO AUTOMATICAMENTE NO DISCO C:/")

opcao = input("INFORME A OPÇÃO (1) OU (2): ")

while opcao not in ["1", "2"]:
    print("\nOPÇÃO INVÁLIDA!\n")
    print("-> OPÇÕES DE DIRETÓRIO: ")
    print("1 - INFORMAR O CAMINHO DO DIRETÓRIO")
    print("2 - INSERIR OS ARQUIVOS NO DIRETÓRIO 'MERGER' CRIADO AUTOMATICAMENTE NO DISCO C:/ \n")
    opcao = input("INFORME A OPÇÃO (1) OU (2): \n")

# RODANDO COM DIRETÓRIO INFORMADO PELO USUÁRIO
if opcao == "1":
    diretorio = input("\nINFORME O CAMINHO COMPLETO DO DIRETÓRIO: ")

    while True:
        try:
            os.chdir(diretorio)
            print("- DIRETÓRIO LOCALIZADO COM SUCESSO")
            break
        # except:
        #     print("\nCAMINHO INVÁLIDO, EXECUTE O BOT NOVAMENTE E FORNEÇA UM CAMINHO VÁLIDO!")
        except FileNotFoundError:
            print("ERRO: O DIRETÓRIO INFORMADO NÃO FOI ENCONTRADO")
            diretorio = input("\nINFORME O CAMINHO COMPLETO DO DIRETÓRIO: ")
        except NotADirectoryError:
            print("ERRO: O CAMINHO INFORMADO NÃO É UM DIRETÓRIO VÁLIDO")
            diretorio = input("\nINFORME O CAMINHO COMPLETO DO DIRETÓRIO: ")
        except PermissionError:
            print("ERRO: PERMISSÃO NEGADA PARA ACESSAR ESSE DIRETÓRIO")
            diretorio = input("\nINFORME O CAMINHO COMPLETO DO DIRETÓRIO: ")

    files_list = os.listdir()

    if len(files_list) > 0:
        # Criar um novo documento vazio para mesclar os PDFs
        merger = fitz.open()

        for file in files_list:
            if ".pdf" in file:
                pdf_atual = fitz.open(file)
                merger.insert_pdf(pdf_atual)
                pdf_atual.close()

        merger.save("PDF_MESCLADO.PDF")
        merger.close()

        print(f"- ARQUIVO MESCLADO COM SUCESSO\n- DISPONÍVEL NO DIRETÓRIO -> {diretorio} \n- ESTE TERMINAL SERÁ FECHADO AUTOMATICAMENTE EM 10 SEGUNDOS...")
        sleep(10)
        exit()

    else:
        print("- O DIRETÓRIO INFORMADO ESTÁ VAZIO.\n- EXECUTE O BOT NOVAMENTE.\n- ESTE TERMINAL SERÁ FECHADO AUTOMATICAMENTE EM 10 SEGUNDOS...")
        sleep(10)
        exit()