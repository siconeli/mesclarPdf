import fitz # biblioteca PyMuPDF -> para mesclar arquivos pdf
import os
from time import sleep
from docx2pdf import convert # biblioteca docx2pdf -> pip install docx2pdf -> converte docx para pdf

# Obrigatório ter o Microsoft Word instalado no computador hospedeiro
def converterDocx(files_list):
    for file in files_list:
        try:
            if ".docx" in file: #.doc
                try:
                    print(f"(RETORNO): CONVERTENDO O ARQUIVO -> {file}")  
                    final_file = file.replace(".docx", "")
                    convert(file, f"{final_file}.pdf")
                except Exception as e:
                    print(f"(ERRO): ERRO NA CONVERSÃO DO ARQUIVO DOCX -> {e}")
        except:
            print(f"\n(ERRO): ERRO NA CONVERSÃO DOS ARQUIVOS WORD PARA PDF!")

def mesclarPdf(files_list, file_name, diretorio):
    try:
        diretorio_final = "DOCSMERGER"
        if not os.path.exists(diretorio_final):
            os.mkdir(diretorio_final)

        merger = fitz.open()

        for file in files_list:
            if ".pdf" in file:
                pdf_atual = fitz.open(file)
                merger.insert_pdf(pdf_atual)
                pdf_atual.close()

        merger.save(f"{diretorio}\{diretorio_final}\{file_name}.PDF")
        merger.close()

        return True

    except:
        print("\n(ERRO): ERRO NA MESCLAGEM DOS ARQUIVOS, VERIFIQUE OS ARQUIVOS E TENTE NOVAMENTE!")

print("### BEM-VINDO AO MERGER ###")
print("\n### O MERGER TEM A FUNÇÃO DE CONVERTER ARQUIVOS DOCX PARA PDF, EM SEGUIDA MESCLAR E GERAR UN ÚNICO ARQUIVO ###")
print("\n### ESCOLHA AS OPÇÕES DISPONÍVEIS NO MENU ###")

encerrar = "0"

while encerrar != "3":
    print("\n[ MENU - OPÇÕES DE DIRETÓRIO ] ")
    print("1 - INFORMAR O CAMINHO DO DIRETÓRIO")
    print("2 - INSERIR ARQUIVOS NO DIRETÓRIO 'MERGER' CRIADO AUTOMATICAMENTE NO DISCO C:/")
    print("3 - ENCERRAR MERGER")

    opcao = input("\n(MERGER): INFORME A OPÇÃO (1), (2) OU (3): ")

    if opcao == "3":    
        print("\n(RETORNO): ENCERRANDO O MERGER...")
        sleep(3)
        exit()

    while opcao not in ["1", "2", "3"]:
        print("\n(RETORNO): OPÇÃO INVÁLIDA!")
        print("\nMENU - OPÇÕES DE DIRETÓRIO: ")
        print("1 - INFORMAR O CAMINHO DO DIRETÓRIO")
        print("2 - INSERIR OS ARQUIVOS NO DIRETÓRIO 'MERGER' CRIADO AUTOMATICAMENTE NO DISCO C:/")
        print("3 - ENCERRAR MERGER")
        opcao = input("\n(MERGER): INFORME A OPÇÃO (1), (2) OU (3): ")
        
        if opcao == "3":    
            print("\n(RETORNO): ENCERRANDO O MERGER...")
            sleep(3)
            exit()


    # RODANDO COM DIRETÓRIO INFORMADO PELO USUÁRIO
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
            print("\n(RETORNO): ABRINDO DIRETÓRIO...")
            sleep(2)
            os.startfile(diretorio)

        while len(files_list) == 0:
            files_list = os.listdir()
            sleep(2)
            print("\n(RETORNO): ADICIONE ARQUIVOS NO DIRETÓRIO...")

        if len(files_list) > 0:
            file_name = input("\n(MERGER): INFORME O NOME DO ARQUIVO FINAL: ")

            print("\n(RETORNO): MESCLANDO ARQUIVOS...")

            if mesclarPdf(files_list, file_name):
                print("\n------------------------------------------------------------------")
                print("(RETORNO): ARQUIVO MESCLADO COM SUCESSO")
                print(f"\n(RETORNO): DISPONÍVEL NO DIRETÓRIO -> {diretorio}")
                print("\n(RETORNO): ABRINDO DIRETÓRIO...")
                sleep(2)
                os.startfile(diretorio)
                sleep(2)
                print("------------------------------------------------------------------")

    # RODANDO COM DIRETÓRIO FIXO NO C:/
    if opcao == "2":
        while True:
            c_dir = "C:/"

            try:
                os.chdir(c_dir)
                break

            except FileNotFoundError:
                print(f"\n(ERRO): O DIRETÓRIO NÃO FOI ENCONTRADO -> {c_dir}")
                sleep(10)
                exit()

            except NotADirectoryError:
                print(f"\n(ERRO): O CAMINHO INFORMADO NÃO É UM DIRETÓRIO VÁLIDO -> {c_dir}")
                sleep(10)
                exit()

            except PermissionError:
                print(f"\n(ERRO): PERMISSÃO NEGADA PARA ACESSAR ESSE DIRETÓRIO -> {c_dir}")
                sleep(10)
                exit()

        if(os.path.isdir("MERGER")):
            os.chdir("MERGER")
            files_list = os.listdir()

            if len(files_list) == 0:
                print("\n(RETORNO): DIRETÓRIO VAZIO...")
                print("\n(RETORNO): ADICIONE ARQUIVOS NO DIRETÓRIO...")
                print("\n(RETORNO): ABRINDO DIRETÓRIO...")
                sleep(2)
                os.startfile("C:\merger")
        else:
            os.mkdir("MERGER")
            os.chdir("MERGER")
            print("\n(RETORNO): DIRETÓRIO MERGER CRIADO COM SUCESSO!")
            print("\n(RETORNO): ADICIONE ARQUIVOS NO DIRETÓRIO...")
            sleep(2)
            os.startfile("C:\MERGER")

        files_list = os.listdir()

        while len(files_list) == 0:
            files_list = os.listdir()
            sleep(2)
            print("\n(RETORNO): ADICIONE ARQUIVOS NO DIRETÓRIO...")

        if len(files_list) > 0:
            validation = False

            # APLICAR LÓGICA PARA VERIFICAR SE EXISTE ARQUIVO WORD NA PASTA!
            for file in files_list:
                if ".docx" in file:
                    validation = True       

            if validation:
                print("\n(RETORNO): CONVERTENDO ARQUIVOS WORD PARA PDF...\n")
                converterDocx(files_list)

            file_name = input("\n(MERGER): INFORME O NOME DO ARQUIVO FINAL: ")

            print("\n(RETORNO): MESCLANDO ARQUIVOS...")

            files_list = os.listdir()

            diretorio = "C:\MERGER"
            if mesclarPdf(files_list, file_name, diretorio):
                print("\n------------------------------------------------------------------")
                print("(RETORNO): ARQUIVO MESCLADO COM SUCESSO")
                print(f"\n(RETORNO): DISPONÍVEL NO DIRETÓRIO -> {diretorio}\DOCSMERGER")
                print("\n(RETORNO): ABRINDO DIRETÓRIO...")
                sleep(2)
                os.startfile("C:\MERGER")
                print("------------------------------------------------------------------")

