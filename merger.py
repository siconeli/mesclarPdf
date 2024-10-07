import fitz # biblioteca PyMuPDF
import os
from time import sleep

def mesclarPdf(files_list, file_name):
    try:
        merger = fitz.open()

        for file in files_list:
            if ".pdf" in file:
                pdf_atual = fitz.open(file)
                merger.insert_pdf(pdf_atual)
                pdf_atual.close()

        merger.save(f"{file_name}.PDF")
        merger.close()

        return True

    except:
        print("ERRO NA MESCLAGEM DOS ARQUIVOS!")

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
    opcao = input("INFORME A OPÇÃO (1) OU (2): ")

# RODANDO COM DIRETÓRIO INFORMADO PELO USUÁRIO
if opcao == "1":
    diretorio = input("\nINFORME O CAMINHO COMPLETO DO DIRETÓRIO: ")

    while True:
        try:
            os.chdir(diretorio)
            print("- DIRETÓRIO LOCALIZADO COM SUCESSO")
            break

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

    while len(files_list) == 0:
        files_list = os.listdir()
        sleep(2)
        print("\nADICIONE ARQUIVOS NO DIRETÓRIO...")

    if len(files_list) > 0:
        file_name = input("\nINFORME O NOME DO ARQUIVO FINAL: ")

        if mesclarPdf(files_list, file_name):
            diretorio = "C:\merger"
            print("\n------------------------------------------------------------------")
            print("- ARQUIVO MESCLADO COM SUCESSO")
            print(f"\n- DISPONÍVEL NO DIRETÓRIO -> {diretorio}")
            print("\n- ESTE TERMINAL SERÁ FECHADO AUTOMATICAMENTE EM 10 SEGUNDOS...")
            print("------------------------------------------------------------------")
            sleep(10)
            exit()

# RODANDO COM DIRETÓRIO FIXO NO C:/
if opcao == "2":
    while True:
        c_dir = "C:/"

        try:
            os.chdir(c_dir)
            break

        except FileNotFoundError:
            print(f"ERRO: O DIRETÓRIO NÃO FOI ENCONTRADO -> {c_dir}")
            sleep(10)
            exit()

        except NotADirectoryError:
            print(f"ERRO: O CAMINHO INFORMADO NÃO É UM DIRETÓRIO VÁLIDO -> {c_dir}")
            sleep(10)
            exit()

        except PermissionError:
            print(f"ERRO: PERMISSÃO NEGADA PARA ACESSAR ESSE DIRETÓRIO -> {c_dir}")
            sleep(10)
            exit()

    if(os.path.isdir("merger")):
        os.chdir("merger")
    else:
        os.mkdir("merger")
        print("- DIRETÓRIO MERGER CRIADO COM SUCESSO!")

    files_list = os.listdir()

    while len(files_list) == 0:
        files_list = os.listdir()
        sleep(2)
        print("\nADICIONE ARQUIVOS NO DIRETÓRIO...")

    if len(files_list) > 0:
        file_name = input("\nINFORME O NOME DO ARQUIVO FINAL: ")

        if mesclarPdf(files_list, file_name):
            diretorio = "C:\merger"
            print("\n------------------------------------------------------------------")
            print("- ARQUIVO MESCLADO COM SUCESSO")
            print(f"\n- DISPONÍVEL NO DIRETÓRIO -> {diretorio}")
            print("\n- ESTE TERMINAL SERÁ FECHADO AUTOMATICAMENTE EM 10 SEGUNDOS...")
            print("------------------------------------------------------------------")
            sleep(10)
            exit()



