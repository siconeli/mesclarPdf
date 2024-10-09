import fitz # biblioteca PyMuPDF -> para mesclar arquivos pdf
import os
from time import sleep
from docx2pdf import convert # biblioteca docx2pdf -> pip install docx2pdf -> converte docx para pdf
import shutil
import subprocess
# Utilizando o pyinstaller -> para criar o file.exe da aplicação -> Comand: pyinstaller "nome da aplicação .py"
# Utilizando o GhostScript -> para comprimir arquivos PDF -> Fazer o download em: https://www.ghostscript.com/releases/gsdnld.html -> Versão: Ghostscript 10.04.0 para Windows (64 bits) Lançamento do Ghostscript AGPL

def converterDocx(files_list):
    for file in files_list:
        if ".docx" in file:
            try:
                print(f"(RETORNO): CONVERTENDO O ARQUIVO -> {file}")  
                final_file = file.replace(".docx", "")
                convert(file, f"{final_file}.pdf")
                os.remove(file)
            except Exception as e:
                print(f"(ERRO): ERRO NA CONVERSÃO DO ARQUIVO -> {file} -> {e}")
                print("\n(RETORNO): O MERGER SERÁ FECHADO EM 60 SEGUNDOS, MOSTRE A MENSAGEM DE ERRO AO RESPONSÁVEL...")
                sleep(60)
                exit()
                
    return True
    
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
        print("\n(RETORNO): O MERGER SERÁ FECHADO EM 60 SEGUNDOS, MOSTRE A MENSAGEM DE ERRO AO RESPONSÁVEL...")
        sleep(60)
        exit()


# Utilizando GhostScript -> Para comprimir arquivos PDF -> fazer o download em: https://www.ghostscript.com/releases/gsdnld.html -> Versão: Ghostscript 10.04.0 para Windows (64 bits) Lançamento do Ghostscript AGPL
def comprimirPdf(input_pdf, output_pdf, quality="screen"):
    try:
        with open("C:\merger\gs_path.txt", "r", encoding="utf-8") as gs_path:
            gs_path = gs_path.readline()

    except Exception as e:
        print(f"\n(ERRO): TXT COM O CAMINHO DO GHOSTSCRIPT NÃO LOCALIZADO -> {e}")
        print("\n(RETORNO): O MERGER SERÁ FECHADO EM 60 SEGUNDOS, MOSTRE A MENSAGEM DE ERRO AO RESPONSÁVEL...")
        sleep(60)
        exit()

    try:
        subprocess.run([
            gs_path, '-sDEVICE=pdfwrite', '-dCompatibilityLevel=1.4',
            f'-dPDFSETTINGS=/{quality}',
            '-dNOPAUSE', '-dQUIET', '-dBATCH',
            f'-sOutputFile={output_pdf}', input_pdf
        ], check=True)
        
        return True
    except Exception as e:
        print(f"\n(ERRO): O COMPUTADOR HOSPEDEIRO NÃO POSSUI O GHOSTSCRIPT INSTALADO OU O SEU DIRETÓRIO NÃO FOI INFORMADO CORRETAMENTE... ENTRE EM CONTATO COM O RESPONSÁVEL PELO MERGER -> Erro: {e}")
        print("\n(RETORNO): O MERGER SERÁ FECHADO EM 60 SEGUNDOS, MOSTRE A MENSAGEM DE ERRO AO RESPONSÁVEL...")
        sleep(60)
        exit()

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
                print("\n(RETORNO): ARQUIVO MESCLADO COM SUCESSO")

                comprimir = input("\n(MERGER): COMPRIMIR ARQUIVO? (S) OU (N): ")
                comprimir = comprimir.upper()

                while comprimir not in ["S", "N"]:
                    comprimir = input("\n(MERGER): COMPRIMIR ARQUIVO? (S) OU (N): ")
                    comprimir = comprimir.upper()

                if comprimir == "S":
                    os.chdir(arquivos_mesclados)

                    files_list = os.listdir()

                    for file in files_list:
                        output_file = f"c.{file}"
                        if comprimirPdf(file, output_file, quality="screen"):
                            os.remove(file)
                            print("\n(RETORNO): ARQUIVO COMPRIMIDO COM SUCESSO")

                print(f"\n(RETORNO): DISPONÍVEL NO DIRETÓRIO -> {arquivos_mesclados}")
                sleep(2)
                os.startfile(arquivos_mesclados)
                sleep(2)

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
                if converterDocx(files_list):
                    print("\n(RETORNO): ARQUIVOS CONVERTIDOS COM SUCESSO")

                    comprimir = input("\n(MERGER): COMPRIMIR ARQUIVO? (S) OU (N): ")
                    comprimir = comprimir.upper()

                    while comprimir not in ["S", "N"]:
                        comprimir = input("\n(MERGER): COMPRIMIR ARQUIVO? (S) OU (N): ")
                        comprimir = comprimir.upper()

                    if comprimir == "S":
                        os.chdir(arquivos_convertidos)

                        files_list = os.listdir()

                        for file in files_list:
                            output_file = f"c.{file}"
                            if comprimirPdf(file, output_file, quality="screen"):
                                os.remove(file)
                                print(f"\n(RETORNO): ARQUIVO {file} COMPRIMIDO COM SUCESSO")
               
                    print(f"\n(RETORNO): DISPONÍVEL NO DIRETÓRIO -> {arquivos_convertidos}")
                    sleep(2)
                    os.startfile(arquivos_convertidos)




























