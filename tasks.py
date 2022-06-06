from openpyxl import load_workbook
import pyautogui
import time
import pyperclip
import os


planilha = load_workbook("Relatorio.xlsx")

def count_rows(sheet_index: int) -> int:
    """
    Retorna a quantidade de linhas da tabela de acordo com o índice fornecido.

    sheet_index = 0 -> Capa
    sheet_index = 1 -> Dados coletados
    """

    table = planilha.worksheets[sheet_index]

    cont = 1
    
    while True:
        if table[f"A{str(cont)}"].value != None:
            cont+=1
        else:
            n_rows = int(cont - 3)
            break
    return n_rows


def open_model(path: str):

    """
    Executa o formulário onde os dados/informações serão inseridos.
    Utilize a variável 'word_model' localizada no módulo paths.py como parâmetro.

    Software -> Word
    """

    pyautogui.hotkey("winleft", "r")    #Atalho 'Executar'
    time.sleep(1)
    pyautogui.write(path)
    time.sleep(0.5)
    pyautogui.press("enter")

    while True:
        try:
            time.sleep(0.2)
            pyautogui.click('images\\click_icons\\word_icons\\selecionar_pagina_modelo\\selecionar_pagina_modelo.png')      #Word - Ponto de referência
            break
        except TypeError:
            try:
                pyautogui.click('images\\click_icons\\word_icons\\selecionar_pagina_modelo\\selecionar_pagina_modelo_2.png')      #Word - Ponto de referência
                break
            except TypeError:
                try:
                    pyautogui.click('images\\click_icons\\word_icons\\selecionar_pagina_modelo\\selecionar_pagina_modelo_3.png')      #Word - Ponto de referência
                    break
                except TypeError:
                    try:
                        pyautogui.click('images\\click_icons\\word_icons\\selecionar_pagina_modelo\\selecionar_pagina_modelo_4.png')      #Word - Ponto de referência
                        break
                    except TypeError:
                        try:
                            pyautogui.click('images\\click_icons\\word_icons\\selecionar_pagina_modelo\\selecionar_pagina_modelo_5.png')      #Word - Ponto de referência
                            break
                        except TypeError:
                                print("Procurando 'Modelo - Word'...")

    time.sleep(0.5)
    pyautogui.hotkey('winleft', 'up')   #Atalho 'Tela cheia'

    while True:
        try:
            pyautogui.click('images\\click_icons\\word_icons\\word_exibir\\word_exibir.png')    #Word - Exibir
            break
        except TypeError:
            try:
                pyautogui.click('images\\click_icons\\word_icons\\word_exibir\\word_exibir_2.png')    #Word - Exibir
                break
            except TypeError:
                try:
                    pyautogui.click('images\\click_icons\\word_icons\\word_exibir\\word_exibir_3.png')    #Word - Exibir
                    break
                except TypeError:
                    try:
                        pyautogui.click('images\\click_icons\\word_icons\\word_exibir\\word_exibir_4.png')    #Word - Exibir
                        break
                    except TypeError:
                        print("Procurando ícone 'Exibir'...")

    while True:
        try:
            pyautogui.click('images\\click_icons\\word_icons\\word_view_uma_pagina\\word_view_uma_pagina.png')    #Phantom PDF - View -> Uma Página
            break
        except TypeError:
            try:
                pyautogui.click('images\\click_icons\\word_icons\\word_view_uma_pagina\\word_view_uma_pagina_2.png')    #Phantom PDF - View -> Uma Página
                break
            except TypeError:
                try:
                    pyautogui.click('images\\click_icons\\word_icons\\word_view_uma_pagina\\word_view_uma_pagina_3.png')    #Phantom PDF - View -> Uma Página
                    break
                except TypeError:
                    print("Procurando ícone 'Uma Página'...")


def create_empty_word(word_exe_path: str):

    """
    Cria um arquvo .docx em branco.
    O método recebe como parâmetro o caminho do Word.exe.
    """

    pyautogui.hotkey("winleft", "r")    #Atalho 'Executar'
    time.sleep(1)
    pyautogui.write(word_exe_path, interval=0.01)
    time.sleep(0.5)
    pyautogui.press("enter")

    while True:
        try:
            pyautogui.click('images\\click_icons\\word_icons\\word_name\\word_name.png')    #Word - Ponto de referência
            break
        except TypeError:
            try:
                pyautogui.click('images\\click_icons\\word_icons\\word_name\\word_name_2.png')    #Word - Ponto de referência
                break
            except TypeError:
                try:
                    pyautogui.click('images\\click_icons\\word_icons\\word_name\\word_name_3.png')    #Word - Ponto de referência
                    break
                except TypeError:
                    print("Procurando 'Word Name'...")

    pyautogui.hotkey('winleft', 'up')   #Atalho 'tela cheia'
    time.sleep(1)

    while True:
        try:
            pyautogui.click('images\\click_icons\\word_icons\\word_documento_em_branco\\word_documento_em_branco.png')    #Word - Documento em branco
            break
        except TypeError:
            try:
                pyautogui.click('images\\click_icons\\word_icons\\word_documento_em_branco\\word_documento_em_branco_1.png')    #Word - Documento em branco
                break
            except TypeError:
                try:
                    pyautogui.click('images\\click_icons\\word_icons\\word_documento_em_branco\\word_documento_em_branco_2.png')    #Word - Documento em branco
                    break
                except TypeError:
                    try:
                        pyautogui.click('images\\click_icons\\word_icons\\word_documento_em_branco\\word_documento_em_branco_3.png')    #Word - Documento em branco
                        break
                    except TypeError:
                        try:
                            pyautogui.click('images\\click_icons\\word_icons\\word_documento_em_branco\\word_documento_em_branco_4.png')    #Word - Documento em branco
                            break
                        except TypeError:
                            try:
                                pyautogui.click('images\\click_icons\\word_icons\\word_documento_em_branco\\word_documento_em_branco_5.png')    #Word - Documento em branco
                                break
                            except TypeError:
                                print("Procurando 'Documento em branco'...")
 
    while True:
        try:
            pyautogui.click('images\\click_icons\\word_icons\\word_documento1\\word_documento1.png')    #Word - Documento em branco
            break
        except TypeError:
            try:
                pyautogui.click('images\\click_icons\\word_icons\\word_documento1\\word_documento1_2.png')    #Word - Documento em branco
                break
            except TypeError:
                try:
                    pyautogui.click('images\\click_icons\\word_icons\\word_documento1\\word_documento1_3.png')    #Word - Documento em branco
                    break
                except TypeError:
                    print("Procurando 'Documento1 - Word'...")

    n_pages = count_rows(1)
    time.sleep(0.5)
    pyautogui.hotkey("ctrl", "home")    #Atalho 'primeira página'

    while True:
        try:
            pyautogui.click('images\\click_icons\\word_icons\\word_exibir\\word_exibir.png')    #Word - Exibir
            break
        except TypeError:
            try:
                pyautogui.click('images\\click_icons\\word_icons\\word_exibir\\word_exibir_2.png')    #Word - Exibir
                break
            except TypeError:
                try:
                    pyautogui.click('images\\click_icons\\word_icons\\word_exibir\\word_exibir_3.png')    #Word - Exibir
                    break
                except TypeError:
                    try:
                        pyautogui.click('images\\click_icons\\word_icons\\word_exibir\\word_exibir_4.png')    #Word - Exibir
                        break
                    except TypeError:
                        print("Procurando ícone 'Exibir'...")

    while True:
        try:
            pyautogui.click('images\\click_icons\\word_icons\\word_view_uma_pagina\\word_view_uma_pagina.png')    #View -> Uma Página
            break
        except TypeError:
            try:
                pyautogui.click('images\\click_icons\\word_icons\\word_view_uma_pagina\\word_view_uma_pagina_2.png')    #View -> Uma Página
                break
            except TypeError:
                try:
                    pyautogui.click('images\\click_icons\\word_icons\\word_view_uma_pagina\\word_view_uma_pagina_3.png')    #View -> Uma Página
                    break
                except TypeError:
                    print("Procurando ícone 'Uma Página'...")


def open_file(path: str):

    """
    Abre ou executa o arquivo ou diretório fornecido como parâmetro.
    """

    file_path = path.split("\\")
    print(f'Acessando {file_path[file_path.__len__() - 1]}')
    pyautogui.hotkey('winleft', 'r')    #Atalho 'Executar'

    while True:
        try:
            pyautogui.click('images\\click_icons\\executar\\reference\\reference.png')    #View -> Uma Página
            break
        except TypeError:
            print("Procurando ícone 'Executar'")

    pyautogui.write(path, interval=0.01)
    time.sleep(0.3)
    pyautogui.press('enter')

def zoom_adjust(value: str):

    """
    Ajusta o zoom para o valor do parâmetro fornecido.
    Software -> Phantom PDF
    """

    print('Ajustando Zoom...')

    while True:
        try:
            pyautogui.click('images\\click_icons\\pdf_icons\\click_home\\click_home.png')    #Phantom PDF - Home
            break
        except TypeError:
            try:
                pyautogui.click('images\\click_icons\\pdf_icons\\click_home\\click_home_2.png')    #Phantom PDF - Home
                break
            except TypeError:
                try:
                    pyautogui.click('images\\click_icons\\pdf_icons\\click_home\\click_home_3.png')    #Phantom PDF - Home
                    break
                except TypeError:
                    try:
                        pyautogui.click('images\\click_icons\\pdf_icons\\click_home\\click_home_4.png')     #Phantom PDF - Home
                        break
                    except TypeError:
                        print("Procurando ícone 'Home'...")

    time.sleep(1)
    pyautogui.hotkey('ctrl', 'm')   #Atalho 'alterar zoom'
    time.sleep(0.3)
    pyautogui.write(value, interval=0.01)
    time.sleep(0.3)
    pyautogui.press('enter')


def count_pages() -> int:
    """
    Retorna o número total de páginas.
    Software -> Phantom PDF
    """

    print('Extraindo o número de páginas...')

    while True:
        try:
            pyautogui.click('images\\click_icons\\pdf_icons\\click_edit_page\\click_edit_page.png')   #Phantom PDF - Seleciona o número de páginas
            break
        except TypeError:
            print("Procurando ícone 'Número de páginas'...")
    
    pyautogui.hotkey('ctrl', 'c')
    str_contagem = pyperclip.paste()
    int_contagem = int(str_contagem.split(' / ')[1])
    time.sleep(0.3)
    pyautogui.press('esc')

    print('Extração concluída!')

    return int_contagem


def delete_all_files(path: str):
    """
    Deleta todos os arquivos do diretório fornecido.
    """

    for file in os.scandir(path):
        print(f'Removendo: {file.name}')
        os.remove(f"{path}\\{file.name}")
