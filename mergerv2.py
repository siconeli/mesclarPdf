import fitz # biblioteca PyMuPDF -> para mesclar arquivos pdf
import os
from time import sleep
from docx2pdf import convert # biblioteca docx2pdf -> pip install docx2pdf -> converte docx para pdf
# Usando o pyinstaller -> para criar o file.exe da aplicação -> Comand: pyinstaller "nome da aplicação .py"

import shutil

def converterDocx(files_list):
    for file in files_list:
        try:
            if ".docx" in file: #.doc
                try:
                    print(f"(RETORNO): CONVERTENDO O ARQUIVO -> {file}")  
                    final_file = file.replace(".docx", "")
                    convert(file, f"{final_file}.pdf")
                    os.remove(file)

                except Exception as e:
                    print(f"(ERRO): ERRO NA CONVERSÃO DO ARQUIVO DOCX -> {e}")
        except:
            print(f"\n(ERRO): ERRO NA CONVERSÃO DOS ARQUIVOS WORD PARA PDF!")

def mesclarPdf(files_list, file_name, diretorio):
    try:
        merger = fitz.open()

        for file in files_list:
            if ".pdf" in file:
                pdf_atual = fitz.open(file)
                merger.insert_pdf(pdf_atual)
                pdf_atual.close()
                os.remove(file)

        merger.save(f"{diretorio}\{file_name}.PDF")
        merger.close()

        return True
    except Exception as e:
        print(f"\n(ERRO): ERRO NA MESCLAGEM DOS ARQUIVOS -> {e}")

print("### BEM-VINDO AO MERGER ###")
print("\n### OPÇÃO 1 - CONVERTE ARQUIVOS DOCX PARA PDF E MESCLA TODOS OS ARQUIVOS EM UM ÚNICO PDF ###")
print("\n### OPÇÃO 2 - CONVERTER ARQUIVOS DOCX PARA PDF ###")
print("\n### OPÇÃO 3 - ENCERRA A APLICAÇÃO ###")
print("\n### ESCOLHA A FUNCIONALIDADE DESEJADA DE ACORDO COM O NÚMERO ###")

encerrar = "0"

while encerrar != "3":
    print("\n[ MENU ] ")
    print("1 - MESCLAR ARQUIVOS PDF")
    print("2 - CONVERTER ARQUIVOS DOCX PARA PDF")
    print("3 - ENCERRAR MERGER")

    opcao = input("\n(MERGER): INFORME A OPÇÃO (1), (2) OU (3): ")

    if opcao == "3":    
        print("\n(RETORNO): ENCERRANDO O MERGER...")
        sleep(2)
        exit()

    while opcao not in ["1", "2", "3"]:
        print("\n(RETORNO): OPÇÃO INVÁLIDA!")
        print("\n[ MENU ] ")
        print("1 - MESCLAR ARQUIVOS PDF")
        print("2 - CONVERTER ARQUIVOS DOCX PARA PDF")
        print("3 - ENCERRAR MERGER")
        opcao = input("\n(MERGER): INFORME A OPÇÃO (1), (2) OU (3): ")
        
        if opcao == "3":    
            print("\n(RETORNO): ENCERRANDO O MERGER...")
            sleep(2)
            exit()

    if opcao == "1":
        diretorio = input("\n(MERGER): INFORME O CAMINHO COMPLETO DO DIRETÓRIO: ")

        while True:
            try:
                os.chdir(diretorio)
                print("\n(RETORNO): DIRETÓRIO LOCALIZADO COM SUCESSO")
                break

            except FileNotFoundError:
                print("\n(ERRO): O DIRETÓRIO INFORMADO NÃO FOI ENCONTRADO")
                diretorio = input("\n(MERGER): INFORME O CAMINHO COMPLETO DO DIRETÓRIO: ")

            except NotADirectoryError:
                print("\n(ERRO): O CAMINHO INFORMADO NÃO É UM DIRETÓRIO VÁLIDO")
                diretorio = input("\n(MERGER): INFORME O CAMINHO COMPLETO DO DIRETÓRIO: ")

            except PermissionError:
                print("\n(ERRO): PERMISSÃO NEGADA PARA ACESSAR ESSE DIRETÓRIO")
                diretorio = input("\n(MERGER): INFORME O CAMINHO COMPLETO DO DIRETÓRIO: ")

        files_list = os.listdir()

        if len(files_list) == 0:
            print("\n(RETORNO): DIRETÓRIO VAZIO...")
            print("\n(RETORNO): ADICIONE ARQUIVOS NO DIRETÓRIO...")
            sleep(2)
            os.startfile(diretorio)

        while len(files_list) == 0:
            files_list = os.listdir()
            sleep(2)
            print("\n(RETORNO): ADICIONE ARQUIVOS NO DIRETÓRIO...")

        if not os.path.exists("ArquivoMesclado"):
            os.mkdir("ArquivoMesclado")
            print("Diretório criado com sucesso!")


        origem = os.getcwd()
        arquivos_mesclados = os.path.join(origem, "ArquivoMesclado")

        for file in files_list:
            caminho_origem = os.path.join(origem, file)
            caminho_destino = os.path.join(arquivos_mesclados, file)

            if os.path.isfile(caminho_origem):
                shutil.copy2(caminho_origem, caminho_destino)

        os.chdir(arquivos_mesclados)

        files_list = os.listdir()

        if len(files_list) > 0:
            validation = False
            for file in files_list:
                if ".docx" in file:
                    validation = True       

            if validation:
                print("\n(RETORNO): CONVERTENDO ARQUIVOS WORD PARA PDF...\n")
                converterDocx(files_list)
            
            os.chdir("C:/")

            os.chdir(arquivos_mesclados)
            
            files_list = os.listdir()
            files_list.sort()

            file_name = input("\n(MERGER): INFORME O NOME DO ARQUIVO FINAL: ")
            file_name = file_name.upper()

            print("\n(RETORNO): MESCLANDO ARQUIVOS...")

            if mesclarPdf(files_list, file_name, arquivos_mesclados):
                print("\n------------------------------------------------------------------")
                print("(RETORNO): ARQUIVO MESCLADO COM SUCESSO")
                print(f"\n(RETORNO): DISPONÍVEL NO DIRETÓRIO -> {arquivos_mesclados}")
                sleep(2)
                os.startfile(arquivos_mesclados)
                sleep(2)
                print("------------------------------------------------------------------")


    if opcao == "2":
        diretorio = input("\n(MERGER): INFORME O CAMINHO COMPLETO DO DIRETÓRIO: ")

        while True:
            try:
                os.chdir(diretorio)
                print("\n(RETORNO): DIRETÓRIO LOCALIZADO COM SUCESSO")
                break

            except FileNotFoundError:
                print("\n(ERRO): O DIRETÓRIO INFORMADO NÃO FOI ENCONTRADO")
                diretorio = input("\n(MERGER): INFORME O CAMINHO COMPLETO DO DIRETÓRIO: ")

            except NotADirectoryError:
                print("\n(ERRO): O CAMINHO INFORMADO NÃO É UM DIRETÓRIO VÁLIDO")
                diretorio = input("\n(MERGER): INFORME O CAMINHO COMPLETO DO DIRETÓRIO: ")

            except PermissionError:
                print("\n(ERRO): PERMISSÃO NEGADA PARA ACESSAR ESSE DIRETÓRIO")
                diretorio = input("\n(MERGER): INFORME O CAMINHO COMPLETO DO DIRETÓRIO: ")

        files_list = os.listdir()

        if len(files_list) == 0:
            print("\n(RETORNO): DIRETÓRIO VAZIO...")
            print("\n(RETORNO): ADICIONE ARQUIVOS NO DIRETÓRIO...")
            sleep(2)
            os.startfile(diretorio)

        while len(files_list) == 0:
            files_list = os.listdir()
            sleep(2)
            print("\n(RETORNO): ADICIONE ARQUIVOS NO DIRETÓRIO...")

        if not os.path.exists("ArquivosConvertidos"):
            os.mkdir("ArquivosConvertidos")

        origem = os.getcwd()
        arquivos_convertidos = os.path.join(origem, "ArquivosConvertidos")

        for file in files_list:
            caminho_origem = os.path.join(origem, file)
            caminho_destino = os.path.join(arquivos_convertidos, file)

            if os.path.isfile(caminho_origem):
                shutil.copy2(caminho_origem, caminho_destino)

        os.chdir(arquivos_convertidos)

        files_list = os.listdir()

        if len(files_list) > 0:
            validation = False
            for file in files_list:
                if ".docx" in file:
                    validation = True       

            if validation:
                print("\n(RETORNO): CONVERTENDO ARQUIVOS WORD PARA PDF...\n")
                converterDocx(files_list)

        os.chdir(arquivos_convertidos)































