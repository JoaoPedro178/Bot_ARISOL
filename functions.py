import pyautogui
import time
from openpyxl import load_workbook
import tasks
import locale
import paths
import datetime
import matplotlib.pyplot as plt
import numpy as np
import os


locale.setlocale(locale.LC_ALL, "")
planilha = load_workbook(paths.excel_path)
capa = planilha.worksheets[0]
dados_coletados = planilha.worksheets[1]

additional_info_dict = {}
values_dict = {}
additional_info_collumns = "TUV"
collumns = "ABCDEFGHIJKLMNOPQRS"
cells_list = []
field_list = [
    "(equipamento)","(modelo)", "(potencia)", "(fabricante)", "(tensao)", "(corrente)", "(n_serie)",
    "(fase uv 1min)", "(fase vw 1min)", "(fase wu 1min)", "(fase umassa 1min)", "(fase vmassa 1min)",
    "(fase wmassa 1min)", "(fase uv 10min)", "(fase vw 10min)", "(fase wu 10min)", "(fase umassa 10min)",
    "(fase vmassa 10min)", "(fase wmassa 10min)"
]

values_list_10_min = [
    "(fase uv 10min)", "(fase vw 10min)", "(fase wu 10min)", "(fase umassa 10min)",
    "(fase vmassa 10min)", "(fase wmassa 10min)"
]

values_list_1_min = [
    "(fase uv 1min)", "(fase vw 1min)", "(fase wu 1min)", "(fase umassa 1min)", "(fase vmassa 1min)",
    "(fase wmassa 1min)"
]

additional_information_fields = [
    '(instrumento)', '(data_servico)', '(certificado)', '(temperatura)', '(umidade)', '(tensao_aplicada)'
]

field_index = [
    "(indice fase uv)", "(indice fase vw)", "(indice fase wu)", "(indice fase umassa)", "(indice fase vmassa)",
    "(indice fase wmassa)"
]

result_index = [
    "(resultado fase uv)", "(resultado fase vw)", "(resultado fase wu)", "(resultado fase umassa)",
    "(resultado fase vmassa)", "(resultado fase wmassa)"
]

zoom_sumario = "45.52"
zoom_montar_relatorio = "45.52"
zoom_enumerar_paginas = "75"
zoom_informacoes = "75"

def montar_relatorio():

    row_count = 2

    print("Bot Iniciado!")
    print("Não utilize o mouse ou teclado durante o processo!")

    pyautogui.alert("""Bot Iniciado!\n
Não utilize o mouse ou teclado durante o processo!\n
Desenvolvedor: João Pedro Rodrigues""", "Bot ARISOL", "Iniciar Bot")

    time.sleep(2)
    tasks.open_model(paths.word_model)
    pyautogui.hotkey("ctrl", "t")
    time.sleep(0.5)
    pyautogui.hotkey("ctrl", "c")
    time.sleep(1)
    tasks.create_empty_word(paths.word_exe)
    time.sleep(1)

    #Insere o conteúdo do arquivo Modelo.docx no relatório
    for n in range(tasks.count_rows(1) + 1):
        time.sleep(1)
        pyautogui.hotkey("ctrl", "v")
        time.sleep(1)
        pyautogui.press("pagedown")
    
    time.sleep(0.5)
    pyautogui.press("backspace")
    time.sleep(1)
    pyautogui.hotkey("ctrl", "home")
    time.sleep(1)

    #Preenche o relatório utilizando os dados do Relatorio.xlsx
    for n in range(tasks.count_rows(1)):
        row_count+=1
        cells_list = []
        values_dict = {}
        
        #Cria uma lista contendo as células que seram utilizadas
        for i in range(tasks.count_rows(1)):
            for j in collumns:
                cells_list.append(j+str(n+3))
        
        #Cria um dicionário vinculando as células ao seu respectivo campo.
        cont = 0
        for field in field_list:
            values_dict[field] = cells_list[cont]
            cont+=1
        time.sleep(1)

        #Insere os dados no relatório.
        for k, v in values_dict.items():
            time.sleep(1)
            pyautogui.hotkey("ctrl", "u")
            time.sleep(1)
            pyautogui.write(k, interval=0.01)
            time.sleep(0.1)
            pyautogui.press("tab")
            time.sleep(0.1)
            pyautogui.write(f"{dados_coletados[v].value}", interval=0.01)
            time.sleep(0.1)
            pyautogui.press("enter")
            for n in range(2):
                time.sleep(0.1)
                pyautogui.hotkey("shiftleft", "tab")
            time.sleep(0.1)
            pyautogui.press("enter")
            time.sleep(0.1)
            pyautogui.hotkey("altleft", "f4")
            time.sleep(0.5)
            pyautogui.hotkey("ctrl", "home")
        value_1 = 13
        value_2 = 7
        result_dict = {}

        #Calcula o índice de polarização e o insere no relatório
        for i in field_index:
            result =  int(dados_coletados[(values_dict[field_list[value_1]])].value) / int(dados_coletados[(values_dict[field_list[value_2]])].value)
            time.sleep(1)
            pyautogui.hotkey("ctrl", "u")
            time.sleep(1)
            pyautogui.write(i, interval=0.01)
            time.sleep(0.1)
            pyautogui.press("tab")
            time.sleep(0.1)
            pyautogui.write(f"{str(result).split('.')[0]}.{(str(result).split('.')[1])[0:2]}", interval=0.01)
            time.sleep(0.1)
            pyautogui.press("enter")
            for n in range(2):
                time.sleep(0.1)
                pyautogui.hotkey("shiftleft", "tab")
            time.sleep(0.1)
            pyautogui.press("enter")
            time.sleep(0.1)
            pyautogui.hotkey("altleft", "f4")
            time.sleep(0.5)
            pyautogui.hotkey("ctrl", "home")
            value_1+=1
            value_2+=1
            result_dict[i] = result
        
        #Insere os resultados em um dicionário (result_values)
        result_values = {}
        cont = 0
        for k, v in result_dict.items():
            result_values[result_index[cont]] = v
            cont+=1
        
        #Itera sobre o dicionário contendo os resultados (result_values) parar avaliar o resultado.
        for k, v in result_values.items():
            time.sleep(1)
            pyautogui.hotkey("ctrl", "u")
            time.sleep(1)
            pyautogui.write(k, interval=0.01)
            time.sleep(0.1)
            pyautogui.press("tab")
            time.sleep(0.1)
            if v <= 1:
                pyautogui.write("Inaceitavel", interval=0.01)
            elif v < 1.5:
                pyautogui.write("Perigoso", interval=0.01)
            elif v >= 1.5 and v < 2:
                pyautogui.write("Regular", interval=0.01)
            elif v >= 2 and v < 3:
                pyautogui.write("Bom", interval=0.01)
            elif v >= 3 and v < 4:
                pyautogui.write("Muito Bom", interval=0.01)
            elif v >= 4:
                pyautogui.write("Otimo", interval=0.01)
            time.sleep(0.1)
            pyautogui.press("enter")
            for n in range(2):
                time.sleep(0.1)
                pyautogui.hotkey("shiftleft", "tab")
            time.sleep(0.1)
            pyautogui.press("enter")
            time.sleep(0.1)
            pyautogui.hotkey("altleft", "f4")
            time.sleep(0.5)
            pyautogui.hotkey("ctrl", "home")
            time.sleep(1)
        for n in range(len(additional_information_fields)):
            search_var = None
            write_var = None
            time.sleep(1)
            pyautogui.hotkey("ctrl", "u")
            time.sleep(1)
            if n == 0:
                search_var = additional_information_fields[n]
                write_var = capa['B2'].value
            elif n == 1:
                search_var = additional_information_fields[n]
                write_var = str(capa['C2'].value)
                write_var = write_var.split(' ', 1)
                write_var = write_var[0]
            elif n == 2:
                search_var = additional_information_fields[n]
                write_var = capa['D2'].value
            elif n == 3:
                search_var = additional_information_fields[n]
                write_var = (f'{dados_coletados[(additional_info_collumns[0] + str(row_count))].value}C°')
            elif n == 4:
                search_var = additional_information_fields[n]
                write_var = (f'{dados_coletados[(additional_info_collumns[1] + str(row_count))].value}%')
            elif n == 5:
                search_var = additional_information_fields[n]
                write_var = dados_coletados[(additional_info_collumns[2] + str(row_count))].value
            pyautogui.write(search_var, interval=0.01)
            time.sleep(0.1)
            pyautogui.press("tab")
            time.sleep(0.1)
            pyautogui.write(write_var, interval=0.01)
            time.sleep(0.1)
            pyautogui.press("enter")
            for n in range(2):
                time.sleep(0.1)
                pyautogui.hotkey("shiftleft", "tab")
            time.sleep(0.1)
            pyautogui.press("enter")
            time.sleep(0.1)
            pyautogui.hotkey("altleft", "f4")
            time.sleep(0.5)
            pyautogui.hotkey("ctrl", "home")

def save_pre_report():
    print('Salvando relatório...')
    while True:
        try:
            pyautogui.click('images\\click_icons\\word_icons\\word_pagina_inicial\\word_pagina_inicial.png')    #Word - Página Inicial
            break
        except TypeError:
            try:
                pyautogui.click('images\\click_icons\\word_icons\\word_pagina_inicial\\word_pagina_inicial_2.png')    #Word - Página Inicial
                break
            except TypeError:
                try:
                    pyautogui.click('images\\click_icons\\word_icons\\word_pagina_inicial\\word_pagina_inicial_3.png')    #Word - Página Inicial
                    break
                except TypeError:
                    try:
                        pyautogui.click('images\\click_icons\\word_icons\\word_pagina_inicial\\word_pagina_inicial_4.png')    #Word - Página Inicial
                        break
                    except TypeError:
                        try:
                            pyautogui.click('images\\click_icons\\word_icons\\word_pagina_inicial\\word_pagina_inicial_5.png')    #Word - Página Inicial
                            break
                        except TypeError:
                            try:
                                pyautogui.click('images\\click_icons\\word_icons\\word_pagina_inicial\\word_pagina_inicial_6.png')    #Word - Página Inicial
                                break
                            except TypeError:
                                print("Procurando 'Página Inicial'")

    pyautogui.hotkey('ctrl', 'b')

    cont = 0
    while cont <= 4:
        time.sleep(0.3)
        try:
            pyautogui.click("images\\click_icons\\word_icons\\word_save_mais_opcoes\\word_save_mais_opcoes.png")
            break
        except TypeError:
            cont+=1
    
    while True:
        try:
            pyautogui.click('images\\click_icons\\word_icons\\word_este_pc\word_este_pc.png', clicks=2, interval=0.3)    #Referência -> Editar caminho
            break
        except TypeError:
            try:
                pyautogui.click('images\\click_icons\\word_icons\\word_este_pc\word_este_pc_2.png', clicks=2, interval=0.3)    #Referência -> Editar caminho
                break
            except TypeError:
                try:
                    pyautogui.click('images\\click_icons\\word_icons\\word_este_pc\word_este_pc_3.png', clicks=2, interval=0.3)    #Referência -> Editar caminho
                    break
                except TypeError:
                    try:
                        pyautogui.click('images\\click_icons\\word_icons\\word_este_pc\word_este_pc_4.png', clicks=2, interval=0.3)    #Referência -> Editar caminho
                        break
                    except TypeError:
                        try:
                            pyautogui.click('images\\click_icons\\word_icons\\word_este_pc\word_este_pc_5.png', clicks=2, interval=0.3)    #Referência -> Editar caminho
                            break
                        except TypeError:
                            print("Procurando ícone 'Este PC'")

    while True:
        try:
            pyautogui.click('images\\click_icons\\word_icons\\word_salvar_como\\word_salvar_como.png')    #Referência -> Salvar como
            break
        except TypeError:
            try:
                pyautogui.click('images\\click_icons\\word_icons\\word_salvar_como\\word_salvar_como_2.png')    #Referência -> Salvar como
                break
            except TypeError:
                try:
                    pyautogui.click('images\\click_icons\\word_icons\\word_salvar_como\\word_salvar_como_3.png')    #Referência -> Salvar como
                    break
                except TypeError:
                    try:
                        pyautogui.click('images\\click_icons\\word_icons\\word_salvar_como\\word_salvar_como_4.png')    #Referência -> Salvar como
                        break
                    except TypeError:
                        print("Procurando 'Salvar como'")

    pyautogui.write('Relatorio_preenchido', interval=0.01)

    while True:
        try:
            pyautogui.click('images\\click_icons\\word_icons\\word_edit_path\\edit_path.png')    #Referência -> Editar Caminho
            break
        except TypeError:
            try:
                pyautogui.click('images\\click_icons\\word_icons\\word_edit_path\\edit_path_2.png')    #Referência -> Editar Caminho
                break
            except TypeError:
                print("Procurando 'Edit Path'")

    pyautogui.write(paths.word_save_path, interval=0.01)
    time.sleep(0.3)
    pyautogui.press('enter')
    time.sleep(0.3)

    while True:
        try:
            pyautogui.click('images\\click_icons\\word_icons\\word_edit_type\\word_edit_type.png')    #Word - Editar tipo do arquivo
            break
        except TypeError:
            try:
                pyautogui.click('images\\click_icons\\word_icons\\word_edit_type\\word_edit_type_2.png')    #Word -> Editar tipo do arquivo
                break
            except TypeError:
                print("Procurando 'Edit Type'")
    time.sleep(0.5)
    while True:
        try:
            pyautogui.click('images\\click_icons\\word_icons\\word_select_type_pdf\\word_select_type_pdf.png')    #Referência -> Editar Caminho
            break
        except TypeError:
            try:
                pyautogui.click('images\\click_icons\\word_icons\\word_select_type_pdf\\word_select_type_pdf_2.png')    #Referência -> Editar Caminho
                break
            except TypeError:
                print("Procurando ícone 'PDF'")

    while True:
        try:
            pyautogui.click('images\\click_icons\\word_icons\\word_save_pdf_pre_report\\word_save_pdf_pre_report.png')    #Referência -> Editar Caminho
            break
        except TypeError:
            try:
                pyautogui.click('images\\click_icons\\word_icons\\word_save_pdf_pre_report\\word_save_pdf_pre_report_2.png')    #Referência -> Editar Caminho
                break
            except TypeError:
                try:
                    pyautogui.click('images\\click_icons\\word_icons\\word_save_pdf_pre_report\\word_save_pdf_pre_report_3.png')    #Referência -> Editar Caminho
                    break
                except TypeError:
                        print("Procurando ícone 'Save'")

    cont = 0
    while cont <= 3:
        try:
            pyautogui.click('images\\click_icons\\word_icons\\word_click_sim\\word_click_sim.png')    #Referência -> Editar Caminho
            break
        except TypeError:
            try:
                pyautogui.click('images\\click_icons\\word_icons\\word_click_sim\\word_click_sim_2.png')    #Referência -> Editar Caminho
                break
            except TypeError:
                try:
                    pyautogui.click('images\\click_icons\\word_icons\\word_click_sim\\word_click_sim_3.png')    #Referência -> Editar Caminho
                    break
                except TypeError:
                    print("Procurando ícone 'Sim'")
                    print('Tentando novamente...')
                time.sleep(0.5)
                cont+=1

    while True:
        try:
            pyautogui.click('images\\click_icons\\pdf_icons\\click_home\\click_home.png')    #PDF - Home
            break
        except TypeError:
            try:
                pyautogui.click('images\\click_icons\\pdf_icons\\click_home\\click_home_2.png')    #PDF - Home
                break
            except TypeError:
                try:
                    pyautogui.click('images\\click_icons\\pdf_icons\\click_home\\click_home_3.png')    #PDF - Home
                    break
                except TypeError:
                    print("Procurando ícone 'Home'")
                    print('Tentando novamente...')
    n_pages = tasks.count_pages()
    while True:
        try:
            pyautogui.click('images\\click_icons\\pdf_icons\\click_home\\click_home.png')    #PDF - Home
            break
        except TypeError:
            try:
                pyautogui.click('images\\click_icons\\pdf_icons\\click_home\\click_home_2.png')    #PDF - Home
                break
            except TypeError:
                try:
                    pyautogui.click('images\\click_icons\\pdf_icons\\click_home\\click_home_3.png')    #PDF - Home
                    break
                except TypeError:
                    print("Procurando ícone 'Home'")
                    print('Tentando novamente...')
    pyautogui.click(131, 385)
    time.sleep(0.5)
    pyautogui.hotkey('ctrl', 'shift', 'd')
    while True:
        try:
            pyautogui.click('images\\click_icons\\pdf_icons\\click_delete_page_ok\\click_delete_page_ok.png')    #Referência -> Page Range
            break
        except TypeError:
            print("Procurando ícone 'Page Range'")
            print('Tentando novamente...')
    time.sleep(0.2)
    pyautogui.write(str(n_pages), interval=0.01)
    time.sleep(0.2)
    pyautogui.press('enter', presses=2, interval=0.5)
    

def add_border():
    print('Adicionando borda ao relatório...')
    tasks.open_file(paths.border_path)
    while True:
        try:
            pyautogui.click('images\\click_icons\\pdf_icons\\click_blinder\\click_blinder.png')
            break
        except TypeError:
            try:
                pyautogui.click('images\\click_icons\\pdf_icons\\click_blinder\\click_blinder_2.png')
                break
            except TypeError:
                print("Procurando ícone 'Blinder1.pdf'")
                print('Tentando novamente...')
    n_pages = tasks.count_pages()
    while True:
        try:
            pyautogui.click('images\\click_icons\\pdf_icons\\click_reference_border\\click_reference_border.png')    #View -> Uma Página
            break
        except TypeError:
            try:
                pyautogui.click('images\\click_icons\\pdf_icons\\click_reference_border\\click_reference_border_2.png')    #View -> Uma Página
                break
            except TypeError:
                print("Procurando ícone 'border.pdf'")
                print('Tentando novamente...')
    tasks.zoom_adjust(zoom_montar_relatorio)
    time.sleep(0.5)
    while True:
        try:
            pyautogui.click('images\click_icons\pdf_icons\click_edit\\click_edit.png', clicks=2, interval= 1)    #PDF - Home
            break
        except TypeError:
            try:
                pyautogui.click('images\\click_icons\\pdf_icons\\click_edit\\click_edit_2.png', clicks=2, interval= 1)    #PDF - Home
                break
            except TypeError:
                try:
                    pyautogui.click('images\\click_icons\\pdf_icons\\click_edit\\click_edit_3.png', clicks=2, interval= 1)    #PDF - Home
                    break
                except TypeError:
                    print("Procurando ícone 'Edit'")
                    print('Tentando novamente...')
    time.sleep(1)
    while True:
        try:
            pyautogui.click('images\\click_icons\\pdf_icons\\click_edit_object\\click_edit_object.png', clicks=2, interval=1)    #View -> Uma Página
            break
        except TypeError:
            print("Procurando ícone 'Edit Object'")
            print('Tentando novamente...')
    while True:
        try:
            pyautogui.click('images\\click_icons\\pdf_icons\\click_mouse_position\\click_mouse_position.png')    #View -> Uma Página
            break
        except TypeError:
            print("Procurando ícone 'Drag'")
            print('Tentando novamente...')
    time.sleep(0.5)
    pyautogui.drag(400, 530, duration=1)     #Selecionar conteúdo
    pyautogui.hotkey('ctrl', 'c')   #Copiar a seleção
    while True:
        try:
            pyautogui.click('images\\click_icons\\pdf_icons\\click_blinder\\click_blinder.png')
            break
        except TypeError:
            try:
                pyautogui.click('images\\click_icons\\pdf_icons\\click_blinder\\click_blinder_2.png')
                break
            except TypeError:
                print("Procurando ícone 'Blinder1.pdf'")
                print('Tentando novamente...')
    while True:
        try:
            pyautogui.click('images\\click_icons\\pdf_icons\\click_edit_object\\click_edit_object.png')    #View -> Uma Página
            break
        except TypeError:
            print("Procurando ícone 'Edit Object'")
            print('Tentando novamente...')
    time.sleep(1)
    while True:
        try:
            pyautogui.click('images\\click_icons\\pdf_icons\\click_view_single_page\\click_view_single_page.png')    #View -> Uma Página
            break
        except TypeError:
            try:
                pyautogui.click('images\\click_icons\\pdf_icons\\click_view_single_page\\click_view_single_page_2.png')    #View -> Uma Página
                break
            except TypeError:
                print("Procurando ícone 'Single Page'")
                print('Tentando novamente...')
    tasks.zoom_adjust(zoom_montar_relatorio)
    while True:
        try:
            pyautogui.click('images\click_icons\pdf_icons\click_edit\\click_edit.png', clicks=2, interval=0.5)    #PDF - Home
            break
        except TypeError:
            try:
                pyautogui.click('images\\click_icons\\pdf_icons\\click_edit\\click_edit_2.png', clicks=2, interval=0.5)    #PDF - Home
                break
            except TypeError:
                try:
                    pyautogui.click('images\\click_icons\\pdf_icons\\click_edit\\click_edit_3.png', clicks=2, interval=0.5)    #PDF - Home
                    break
                except TypeError:
                    print("Procurando ícone 'Edit'")
                    print('Tentando novamente...')
    while True:
        try:
            pyautogui.click('images\\click_icons\\pdf_icons\\click_edit_object\\click_edit_object.png')    #View -> Uma Página
            break
        except TypeError:
            print("Procurando ícone 'Edit Object'")
            print('Tentando novamente...')
    while True:
        try:
            pyautogui.click('images\\click_icons\\pdf_icons\\click_home\\click_home.png')    #PDF - Home
            break
        except TypeError:
            try:
                pyautogui.click('images\\click_icons\\pdf_icons\\click_home\\click_home_2.png')    #PDF - Home
                break
            except TypeError:
                try:
                    pyautogui.click('images\\click_icons\\pdf_icons\\click_home\\click_home_3.png')    #PDF - Home
                    break
                except TypeError:
                    try:
                        pyautogui.click('images\\click_icons\\pdf_icons\\click_home\\click_home_4.png')
                    except TypeError:
                        print("Procurando ícone 'Home'")
                        print('Tentando novamente...')
    pyautogui.press('home')
    time.sleep(1)
    for n in range(n_pages - 1):
        time.sleep(0.3)
        pyautogui.click(1076, 399)
        time.sleep(0.3)
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(0.3)
        pyautogui.press('pagedown')
        time.sleep(0.5)


def enumerar_paginas():
    """
    Utiliza o campo com a letra 'X' para inserir a numeração das páginas
    """

    print("Enumeração de páginas iniciado!")
    int_contagem = tasks.count_pages()
    time.sleep(1)
    pyautogui.click(1148, 437)  #Área em branco
    time.sleep(1)
    print('Primeira página...')
    pyautogui.press('home')
    while True:
        try:
            pyautogui.click('images\\click_icons\\pdf_icons\\click_home\\click_home.png')    #PDF - Home
            break
        except TypeError:
            try:
                pyautogui.click('images\\click_icons\\pdf_icons\\click_home\\click_home_2.png')    #PDF - Home
                break
            except TypeError:
                try:
                    pyautogui.click('images\\click_icons\\pdf_icons\\click_home\\click_home_3.png')    #PDF - Home
                    break
                except TypeError:
                    print("Procurando ícone 'Home'")
                    print('Tentando novamente...')
    tasks.zoom_adjust(zoom_enumerar_paginas)
    while True:
        try:
            pyautogui.click('images\click_icons\pdf_icons\click_edit\\click_edit.png', clicks=2, interval=0.5)    #PDF - Home
            break
        except TypeError:
            try:
                pyautogui.click('images\\click_icons\\pdf_icons\\click_edit\\click_edit_2.png', clicks=2, interval=0.5)    #PDF - Home
                break
            except TypeError:
                try:
                    pyautogui.click('images\\click_icons\\pdf_icons\\click_edit\\click_edit_3.png', clicks=2, interval=0.5)    #PDF - Home
                    break
                except TypeError:
                    print("Procurando ícone 'Edit'")
                    print('Tentando novamente...')
    while True:
        try:
            pyautogui.click('images\\click_icons\\pdf_icons\\click_edit_text\\click_edit_text.png')    #PDF - Home
            break
        except TypeError:
            try:
                pyautogui.click('images\\click_icons\\pdf_icons\\click_edit_text\\click_edit_text_2.png')    #PDF - Home
                break
            except TypeError:
                print("Procurando ícone 'Edit Text'")
                print('Tentando novamente...')
    while True:
        try:
            pyautogui.click('images\\click_icons\\pdf_icons\\click_view_single_page\\click_view_single_page.png')    #View -> Uma Página
            break
        except TypeError:
            try:
                pyautogui.click('images\\click_icons\\pdf_icons\\click_view_single_page\\click_view_single_page_2.png')    #View -> Uma Página
                break
            except TypeError:
                print("Procurando ícone 'Single Page'")
                print('Tentando novamente...')
    pyautogui.click(1106, 478)  #Área em branco
    time.sleep(0.5)
    print('Enumeração inicializada...')
    n_page = 1
    cont = 3
    for n in range(int_contagem - 1):
        print(f'Página: {n_page}')
        if n == 0:
            time.sleep(0.5)
            pyautogui.press('pagedown')
            while True:
                try:
                    pyautogui.click('images\\click_icons\\pdf_icons\\click_x\\click_x_2.png')    #Referência -> Editar Caminho
                    break
                except TypeError:
                    try:
                        pyautogui.click('images\\click_icons\\pdf_icons\\click_x\\click_x.png')    #Referência -> Editar Caminho
                        break
                    except TypeError:
                        print("Procurando ícone 'X'")
                        print('Tentando novamente...')
            pyautogui.press('backspace')
            time.sleep(0.5)
            pyautogui.write('1')
            time.sleep(0.5)
            pyautogui.click(1106, 478)  #Área em branco
            time.sleep(0.5)
            pyautogui.press('pagedown')
            n_page+=1
        else:
            time.sleep(0.5)
            pyautogui.press('pagedown')
            while True:
                try:
                    pyautogui.click('images\\click_icons\\pdf_icons\\click_x\\click_x_2.png')    #Referência -> Editar Caminho
                    break
                except TypeError:
                    try:
                        pyautogui.click('images\\click_icons\\pdf_icons\\click_x\\click_x.png')    #Referência -> Editar Caminho
                        break
                    except TypeError:
                        print("Procurando ícone 'X'")
                        print('Tentando novamente...')
            pyautogui.press('backspace')
            time.sleep(0.5)
            pyautogui.write(f'{cont}')
            time.sleep(0.5)
            pyautogui.click(1106, 478)  #Área em branco
            time.sleep(0.5)
            pyautogui.press('pagedown')
            n_page+=1
            cont += 1
    pyautogui.alert("Enumeração concluída!", "Notificação", "Continuar")


def informacoes():
    print("Inserção de informações inicializada...")
    locale.setlocale(locale.LC_ALL, 'eng_SG')
    while True:
        try:
            ano = int(pyautogui.prompt(text='Informe o ano do relatório\nObs: Utilize números inteiros', title='Informações' , default=''))
            mes = str(pyautogui.prompt(text='Informe o mês do relatório\nObs: Utilize números inteiros', title='Informações' , default=''))
            dia = int(pyautogui.prompt(text='Informe o dia do relatório\nObs: Utilize números inteiros', title='Informações' , default=''))
            embarcacao = str(pyautogui.prompt(text='Informe o nome da embarcação\nObs: Não utilize caracteres especiais ou acentuações', title='Informações' , default=''))
            if (int(mes) < 0 or int(mes) > 12) or (dia < 0 or dia > 31):
                raise TypeError
            else:
                break
        except TypeError:
            ano = None
            mes = None
            dia = None
            embarcacao = None
            pyautogui.alert('O valor informado está incorreto. Tente novamente.', 'Dado inválido', 'Inserir novamente')
        except ValueError:
            ano = None
            mes = None
            dia = None
            embarcacao = None
            pyautogui.alert('O valor informado está incorreto. Tente novamente.', 'Dado inválido', 'Inserir novamente')
    
    locale.setlocale(locale.LC_ALL, '')
    objeto_mes = datetime.datetime.strptime(mes, "%m")
    nome_mes = objeto_mes.strftime("%B")
    tasks.zoom_adjust(zoom_informacoes)

    while True:
        try:
            pyautogui.click('images\click_icons\pdf_icons\click_edit\\click_edit.png')    #PDF - Edit
            break
        except TypeError:
            try:
                pyautogui.click('images\\click_icons\\pdf_icons\\click_edit\\click_edit_2.png')    #PDF - Edit
                break
            except TypeError:
                try:
                    pyautogui.click('images\\click_icons\\pdf_icons\\click_edit\\click_edit_3.png')    #PDF - Edit
                    break
                except TypeError:
                    print("Procurando 'Edit'...")

    while True:
        try:
            pyautogui.click('images\\click_icons\\pdf_icons\\click_edit_object\\click_edit_object.png')    #PDF - Edit Object
            break
        except TypeError:
            print("Procurando 'Edit Object'...")

    while True:
        try:
            pyautogui.click('images\\click_icons\\pdf_icons\\click_view_single_page\\click_view_single_page.png')    #PDF View -> Single page
            break
        except TypeError:
            try:
                pyautogui.click('images\\click_icons\\pdf_icons\\click_view_single_page\\click_view_single_page_2.png')    #PDF View -> Single page
                break
            except TypeError:
                try:
                    pyautogui.click('images\\click_icons\\pdf_icons\\click_view_single_page\\click_view_single_page_3.png')    #PDF View -> Single page
                    break
                except TypeError:
                    print("Procurando 'Single Page'...")

    pyautogui.click(1106, 478)  #Área em branco
    time.sleep(0.5)
    pyautogui.press('home')
    time.sleep(0.5)
    pyautogui.press("pagedown")
    time.sleep(0.5)
    print(f'Inserindo a imagem da embarcação {embarcacao}...')

    while True:
        try:
            pyautogui.click('images\\click_icons\\pdf_icons\\click_imagem_da_embarcacao\\click_imagem_da_embarcacao.png')    #View -> Uma Página
            break
        except TypeError:
            print("Procurando 'Imagem da embarcação'...")


    pyautogui.press("delete")

    while True:
        try:
            pyautogui.click('images\\click_icons\\pdf_icons\\click_add_image\\click_add_image.png')    #PDF Phantom - Add image
            break
        except TypeError:
            print("Procurando ícone 'Add Image'...")

    while True:
        try:
            pyautogui.click('images\\click_icons\\pdf_icons\\click_import_image\\click_import_image.png')    #PDF Phantom - Add image
            break
        except TypeError:
            print("Procurando ícone 'Import Image'")

    while True:
        try:
            pyautogui.click('images\\click_icons\\pdf_icons\\click_edit_path_2\\click_edit_path.png')    #View -> Uma Página
            break
        except TypeError:
            print("Procurando ícone 'Edit Path'...")

    time.sleep(0.4)
    pyautogui.write(paths.vessels_path, interval=0.01)
    time.sleep(0.4)
    pyautogui.press('enter')

    while True:
        try:
            pyautogui.click('images\\click_icons\\pdf_icons\\click_nome_imagem\\click_nome_imagem.png')    #PDF Phantom - Editar nome da imagem
            break
        except TypeError:
            print("Procurando ícone 'Nome Imagem'")

    time.sleep(0.4)
    pyautogui.write(f'{embarcacao}.jpg', interval=0.02)

    while True:
        try:
            pyautogui.click("images\\click_icons\\pdf_icons\\click_abir_imagem\\click_abrir_imagem.png")    #PDF Phantom - Abrir
            break
        except TypeError:
            print("Procurando ícone 'Abir Imagem'...")

    time.sleep(1)
    pyautogui.rightClick(456, 209)      #Clique direito na imagem
    time.sleep(0.5)
    pyautogui.press('tab')
    time.sleep(0.2)
    pyautogui.press('enter')

    while True:
        try:
            pyautogui.click("images\\click_icons\\pdf_icons\\click_appearence\\click_appearance.png")   #PDF Phantom - Appearance
            break
        except TypeError:
            print("Procurando ícone 'Appearence'...")

    while True:
        try:
            pyautogui.click("images\\click_icons\\pdf_icons\\click_general\\click_general.png")     #PDF Phantom - General
            break
        except TypeError:
            print("Procurando 'General'...")

    #Ajusta o tamanho e posiciona a imagem.
    time.sleep(0.4)
    pyautogui.hotkey('ctrl', 'a')
    time.sleep(0.5)
    pyautogui.write('2.54', interval=0.01)
    time.sleep(0.5)
    pyautogui.press('tab')
    time.sleep(0.5)
    pyautogui.hotkey('ctrl', 'a')
    time.sleep(0.5)
    pyautogui.write('2.79', interval=0.01)
    time.sleep(0.5)
    pyautogui.press('tab')
    time.sleep(0.5)
    pyautogui.hotkey('ctrl', 'a')
    time.sleep(0.5)
    pyautogui.write('5', interval=0.01)
    time.sleep(0.5)
    pyautogui.press('tab')
    time.sleep(0.5)
    pyautogui.hotkey('ctrl', 'a')
    time.sleep(0.5)
    pyautogui.write('2.9', interval=0.01)
    time.sleep(0.5)
    pyautogui.hotkey('alt', 'f4')
    time.sleep(0.5)

    #Insere o nome da embarcação no relatório
    print('Inserindo o nome da embarcação...')
    pyautogui.click(703, 556)   #Posiciona o mouse
    time.sleep(0.5)
    pyautogui.drag(-135, -64, duration=0.5)
    time.sleep(0.5)
    pyautogui.press('delete')
    time.sleep(0.5)
    pyautogui.click(325, 62)    #Adicionar texto
    time.sleep(0.5)
    pyautogui.click(580, 525)   #Seleciona o local para adicionar o texto
    time.sleep(0.5)
    pyautogui.write(f"{embarcacao}", interval=0.01)
    time.sleep(0.5)
    pyautogui.hotkey('ctrl', 'a')

    while True:
        try:
            pyautogui.click("images\\click_icons\\pdf_icons\\click_edit_font_size\\click_edit_font_size.png")   #PDF Phantom - Font Size
            break
        except TypeError:
            print("Procurando 'Font Size'...")

    time.sleep(0.5)
    pyautogui.write("18", interval=0.01)
    time.sleep(0.5)
    pyautogui.press('enter')
    time.sleep(0.5)

    #Insere a data no relatório
    print('Inserindo a data do relatório...')

    while True:
        try:
            pyautogui.click("images\\click_icons\\pdf_icons\\click_add_text\\click_add_text.png")   #PDF Phantom - Add text
            break
        except TypeError:
            print("Procurando 'Add Text'...")

    time.sleep(0.5)
    pyautogui.click(580, 545)   #Seleciona o local para adicionar a data
    time.sleep(0.5)
    pyautogui.write(f'{dia} de {nome_mes.title()} de {ano}', interval=0.01)
    time.sleep(0.5)
    pyautogui.hotkey('ctrl', 'a')
    time.sleep(0.5)

    while True:
        try:
            pyautogui.click("images\\click_icons\\pdf_icons\\click_edit_font_size\\click_edit_font_size.png")   #PDF Phantom - Edit font size
            break
        except TypeError:
            print("Procurando ícone 'Font Size'...")

    time.sleep(0.5)
    pyautogui.write("18", interval=0.01)
    time.sleep(0.5)
    pyautogui.press('enter')
    time.sleep(0.5)
    pyautogui.alert('Informações inseridas com sucesso!','Notificação', button= 'Continuar')


def unir_pdf():
    """
    Realiza a junção de arquivos: (
        Apresentaçao_do_relatorio.pdf,
        Assinaturas.pdf,
        Certificado.pdf,
        Relatorio_preenchido.pdf
    )
    """

    print("Estruturando o relatório...")
    time.sleep(1)
    pyautogui.click(131, 385)
    time.sleep(1)
    pyautogui.hotkey('altleft', 'f4')

    while True:
        try:
            pyautogui.click("images\\click_icons\\pdf_icons\\click_yes\\click_yes.png")     #PDF Phantom - Yes
            break
        except TypeError:
            try:
                pyautogui.click('images\\click_icons\\pdf_icons\\click_yes\\click_yes_2.png')   #PDF Phantom - Yes
                break
            except TypeError:
                print("Procurando ícone 'Yes'...")

    tasks.open_file(paths.word_save_path)

    while True:
        try:
            pyautogui.click("images\\click_icons\\pdf_icons\\click_pdf_folder\\click_pdf_folder.png")   #PDF Phantom - PDF Folder
            break
        except TypeError:
            try:
                pyautogui.click("images\\click_icons\\pdf_icons\\click_pdf_folder\\click_pdf_folder_2.png")     #PDF Phantom - PDF Folder
                break
            except TypeError:
                print("Procurando ícone 'PDF'")

    pyautogui.hotkey('winleft', 'up')

    while True:
        try:
            pyautogui.click("images\\\click_icons\\\pdf_icons\\\click_file_select\\click_file_select.png")     #PDF Phantom - File Select
            break
        except TypeError:
            try:
                pyautogui.click("images\\\click_icons\\\pdf_icons\\\click_file_select\\click_file_select_2.png")    #PDF Phantom - File Select
                break
            except TypeError:
                try:
                    pyautogui.click("images\\\click_icons\\\pdf_icons\\\click_file_select\\click_file_select_3.png")    #PDF Phantom - File Select
                    break
                except TypeError:
                    try:
                        pyautogui.click("images\\\click_icons\\\pdf_icons\\\click_file_select\\click_file_select_4.png")    #PDF Phantom - File Select
                        break
                    except TypeError:
                        try:
                            pyautogui.click("images\\\click_icons\\\pdf_icons\\\click_file_select\\click_file_select_5.png")    #PDF Phantom - File Select
                            break
                        except TypeError:
                            print("Procurando ícone 'Select File'...")

    time.sleep(5)
    pyautogui.hotkey('ctrl', 'a')
    time.sleep(1)
    while True:
        try:
            pyautogui.rightClick("images\\click_icons\\pdf_icons\\right_click_combine_reference\\right_click_combine_reference.png")    #PDF Phantom - Combine reference
            break
        except TypeError:
            try:
                pyautogui.rightClick("images\\click_icons\\pdf_icons\\right_click_combine_reference\\right_click_combine_reference_2.png")  #PDF Phantom - Combine reference
                break
            except TypeError:
                try:
                    pyautogui.rightClick("images\\click_icons\\pdf_icons\\right_click_combine_reference\\right_click_combine_reference_3.png")  #PDF Phantom - Combine reference
                    break
                except TypeError:
                    try:
                        pyautogui.rightClick("images\\click_icons\\pdf_icons\\right_click_combine_reference\\right_click_combine_reference_4.png")  #PDF Phantom - Combine reference
                        break
                    except TypeError:
                        print("Procurando ícone 'Combine reference'...")

    while True:
        try:
            pyautogui.click("images\\click_icons\\pdf_icons\\click_combine\\click_combine.png")     #PDF Phantom - Combine
            break
        except TypeError:
            print("Procurando ícone 'Combine'...")
    
    while True:
        try:
            pyautogui.click("images\\click_icons\\pdf_icons\\click_combine_reference\\click_combine_reference.png")     #PDF Phantom - Relatorio_preenchido.pdf
            break
        except TypeError:
            print("Procurando ícone 'Relatorio_preenchido'...")

    while True:
        try:
            pyautogui.click("images\\click_icons\\pdf_icons\\click_move_up\\click_move_up.png", clicks=2, interval=0.5)     #PDF Phantom - Move up
            break
        except TypeError:
            print("Procurando ícone 'Move Up'...")

    while True:
        try:
            pyautogui.click("images\\click_icons\\pdf_icons\\click_combine_files\\click_combine_files.png")     #PDF Phantom - Combine
            break
        except TypeError:
            print("Procurando ícone 'Combine'...")

    pyautogui.alert("Combinação de arquivos concluída!", "Notificação", "Continuar")


def sumario():
    """
    Adiciona o sumário ao relatório.
    """

    print("Inserindo sumário...")
    paginas_array = []
    nome_indices_array = []
    i = 1
    sair = False
    n_pages = tasks.count_pages()
    locale.setlocale(locale.LC_ALL, 'eng_SG')

    #Coleta os índices do usuário
    while sair == False:
        
        try:
            nome_indice = str(pyautogui.prompt(text=f'Digite o título do {i}° índice.', title='Nome do indice' , default=''))
            pagina = int(pyautogui.prompt(text=f'Digite a página do {i}° índice.', title='Pagina do pagina' , default=''))
            nome_indices_array.append(nome_indice)
            paginas_array.append(pagina)
            break_loop = pyautogui.confirm(text='Existem mais índices?', buttons=['Sim', 'Não'])
            if break_loop == 'Não':
                sair = True
            else:
                i+=1
                continue
        except TypeError as e:
            nome_indice = None
            pagina = None
            pyautogui.alert(f'O valor informado está incorreto. Tente novamente:\n{e}', 'Dado inválido', 'Inserir novamente')
            continue
        except ValueError as v:
            nome_indice = None
            pagina = None
            pyautogui.alert(f'O valor informado está incorreto. Tente novamente:\n{v}', 'Dado inválido', 'Inserir novamente')
            continue

    locale.setlocale(locale.LC_ALL, '')

    while True:
        try:
            pyautogui.click('images\\click_icons\\pdf_icons\\click_edit_page\\click_edit_page.png')   #PDF Phantom - Seleciona o número de páginas
            break
        except TypeError:
            print("Procurando 'Páginas'...")

    pyautogui.write("2")
    time.sleep(0.3)
    pyautogui.press('esc')
    time.sleep(0.3)
    pyautogui.press("enter")
    time.sleep(0.3)

    while True:
        try:
            pyautogui.click("images\\click_icons\\pdf_icons\\click_index\\click_index.png")   #PDF Phantom - Bookmarks
            break
        except TypeError:
            print("Procurando ícone 'Bookmarks'")

    time.sleep(0.5)
    pyautogui.click(155, 478)   #Área em branco
    time.sleep(0.5)
    pyautogui.hotkey("ctrl", "a")
    time.sleep(0.5)
    pyautogui.press("delete")
    time.sleep(1)

    #Cria o sumário utilizando os dados inseridos pelo usuário
    for n in range(2):
        while True:
            try:
                pyautogui.click("images\\click_icons\\pdf_icons\\click_add_bookmark\\click_add_bookmark.png")   #PDF Phantom - Add Bookmark
                break
            except TypeError:
                print("Procurando 'Add bookmark'...")

        time.sleep(0.3)

        if n == 0:
            pyautogui.write("1 - Objetivo", interval=0.01)
            time.sleep(0.3)
            pyautogui.press("enter")
        elif n == 1:
            pyautogui.write("2 - Embasamento Tecnico", interval=0.01)
            time.sleep(0.3)
            pyautogui.press("enter")

    cont = 3
    for n in range(paginas_array.__len__()):
        while True:
            try:
                pyautogui.click('images\\click_icons\\pdf_icons\\click_edit_page\\click_edit_page.png')   #PDF Phantom - Seleciona o número de páginas
                break
            except TypeError:
                print("Procurando 'Páginas'...")

        time.sleep(0.5)
        pyautogui.write(str(paginas_array[n]), interval=0.01)
        time.sleep(0.5)
        pyautogui.press('esc')
        time.sleep(0.5)
        pyautogui.press("enter")
        time.sleep(0.5)

        while True:
            try:
                pyautogui.click("images\\click_icons\\pdf_icons\\click_add_bookmark\\click_add_bookmark.png")   #PDF Phantom - Add Bookmark
                break
            except TypeError:
                print("Procurando ícone 'Add bookmark'...")

        time.sleep(0.5)
        pyautogui.write(f"{cont} - {str(nome_indices_array[n])}", interval=0.01)
        time.sleep(0.5)
        pyautogui.press("enter")
        cont+=1

    time.sleep(0.3)
    while True:
        try:
            pyautogui.click('images\\click_icons\\pdf_icons\\click_edit_page\\click_edit_page.png')   #PDF Phantom - Seleciona o número de páginas
            break
        except TypeError:
            print("Procurando ícone 'Página'")

    time.sleep(0.3)
    pyautogui.write(str(n_pages - 2))
    time.sleep(0.3)
    pyautogui.press('esc')
    time.sleep(0.3)
    pyautogui.press("enter")

    for n in range(2):
        while True:
            try:
                pyautogui.click("images\\click_icons\\pdf_icons\\click_add_bookmark\\click_add_bookmark.png")   #PDF Phantom - Add Bookmark
                break
            except TypeError:
                print("Procurando ícone 'Add bookmark'...")

        time.sleep(0.3)

        if n == 0:
            pyautogui.write(f"{i+3} - Comentarios e Recomendacoes", interval=0.01)
            time.sleep(0.5)
            pyautogui.press("enter")
        elif n == 1:
            pyautogui.write(f"{i+4} - Conclusao", interval=0.01)
            time.sleep(0.5)
            pyautogui.press("enter")

    time.sleep(0.5)

    while True:
        try:
            pyautogui.click('images\\click_icons\\pdf_icons\\click_edit_page\\click_edit_page.png')   #PDF Phantom - Seleciona o número de páginas
            break
        except TypeError:
            print("Procurando 'Páginas'")

    time.sleep(0.3)
    pyautogui.write(str(n_pages - 1))    
    time.sleep(0.3)
    pyautogui.press("esc")
    time.sleep(0.3)
    pyautogui.press("enter")
    time.sleep(0.5)

    for n in range(1):
        while True:
            try:
                pyautogui.click("images\\click_icons\\pdf_icons\\click_add_bookmark\\click_add_bookmark.png")   #PDF Phantom - Add Bookmark
                break
            except TypeError:
                print("Procurando 'Add bookmark'")

        time.sleep(0.3)

        if n == 0:
            pyautogui.write(f"{i + 5} - Assinaturas", interval=0.01)
            time.sleep(0.5)
            pyautogui.press("enter")

    while True:
            try:
                pyautogui.click("images\\click_icons\\pdf_icons\\click_add_reference_page\\click_add_reference_page.png")   #PDF Phantom - Add References page
                break
            except TypeError:
                print("Procurando ícone 'Add references page'")

    time.sleep(0.5)

    while True:
            try:
                pyautogui.click("images\\click_icons\\pdf_icons\\click_new_toc_from_bookmarks\\click_new_toc_from_bookmarks.png")   #PDF Phantom - Add References page
                break
            except TypeError:
                print("Procurando 'New TOC from bookmarks'...")

    while True:
            try:
                pyautogui.click("images\\click_icons\\pdf_icons\\click_bookmarks_ok\\click_bookmarks_ok.png")   #PDF Phantom - Ok
                break
            except TypeError:
                try:
                    pyautogui.click("images\\click_icons\\pdf_icons\\click_bookmarks_ok\\click_bookmarks_ok_2.png")   #PDF Phantom - Ok
                    break
                except TypeError:
                    try:
                        pyautogui.click("images\\click_icons\\pdf_icons\\click_bookmarks_ok\\click_bookmarks_ok_3.png")   #PDF Phantom - Ok
                        break
                    except TypeError:
                        print("Procurando 'Ok'...")

    time.sleep(0.5)
    pyautogui.press("home")

    while True:
            try:
                pyautogui.click("images\\click_icons\\pdf_icons\\click_pages\\click_pages.png")   #PDF Phantom - Pages
                break
            except TypeError:
                print("Procurando 'Pages'...")

    pyautogui.click(147, 284)   #Selecionar página
    time.sleep(1)
    pyautogui.drag(0, 260, duration=0.5)
    time.sleep(1)

    while True:
            try:
                pyautogui.click("images\\click_icons\\pdf_icons\\click_hide\\click_hide.png")   #PDF Phantom - Hide
                break
            except TypeError:
                print("Procurando ícone 'Hide'")

    time.sleep(0.3)

    while True:
        try:
            pyautogui.click('images\\click_icons\\pdf_icons\\click_edit_page\\click_edit_page.png')   #PDF Phantom - Edit page
            break
        except TypeError:
            print("Procurando 'Páginas'...")

    time.sleep(0.3)
    pyautogui.write("2")
    time.sleep(0.3)
    pyautogui.press('esc')
    time.sleep(0.3)
    pyautogui.press("enter")
    time.sleep(1)

    while True:
        try:
            pyautogui.click('images\click_icons\pdf_icons\click_edit\\click_edit.png')    #PDF - Edit
            break
        except TypeError:
            try:
                pyautogui.click('images\\click_icons\\pdf_icons\\click_edit\\click_edit_2.png')    #PDF - Edit
                break
            except TypeError:
                try:
                    pyautogui.click('images\\click_icons\\pdf_icons\\click_edit\\click_edit_3.png')    #PDF - Edit
                    break
                except TypeError:
                    print("Procurando 'Edit'...")

    while True:
        try:
            pyautogui.click('images\\click_icons\\pdf_icons\\click_view_single_page\\click_view_single_page.png')    #PDF - Single Page View
            break
        except TypeError:
            try:
                pyautogui.click('images\\click_icons\\pdf_icons\\click_view_single_page\\click_view_single_page_2.png')    #PDF - Single Page View
                break
            except TypeError:
                try:
                    pyautogui.click('images\\click_icons\\pdf_icons\\click_view_single_page\\click_view_single_page_3.png')    #PDF - Single Page View
                    break
                except TypeError:
                    print("Procurando 'Single Page'...")

    tasks.zoom_adjust(zoom_sumario)

    while True:
        try:
            pyautogui.click('images\click_icons\pdf_icons\click_edit\\click_edit.png')    #PDF - Edit
            break
        except TypeError:
            try:
                pyautogui.click('images\\click_icons\\pdf_icons\\click_edit\\click_edit_2.png')    #PDF - Edit
                break
            except TypeError:
                try:
                    pyautogui.click('images\\click_icons\\pdf_icons\\click_edit\\click_edit_3.png')    #PDF - Edit
                    break
                except TypeError:
                    print("Procurando ícone 'Edit'")

    while True:
        try:
            pyautogui.click('images\\click_icons\\pdf_icons\\click_edit_text\\click_edit_text.png')    #PDF - Edit text
            break
        except TypeError:
            try:
                pyautogui.click('images\\click_icons\\pdf_icons\\click_edit_text\\click_edit_text_2.png')    #PDF - Edit text
                break
            except TypeError:
                print("Procurando 'Edit Text'...")

    time.sleep(0.5)
    pyautogui.click(699, 220)   #Editar o titulo do sumário
    time.sleep(1)
    pyautogui.hotkey("ctrl", "a")
    time.sleep(1)
    pyautogui.write("Sumario", interval=0.02)
    time.sleep(1)
    pyautogui.hotkey("ctrl", "a")
    time.sleep(1)

    while True:
        try:
            pyautogui.click('images\\click_icons\\pdf_icons\\click_edit_font_size_2\\click_edit_font_size.png')   #PDF Phantom - Edit Font Size
            break
        except TypeError:
            print("Procurando 'Font size'...")

    time.sleep(0.3)
    pyautogui.write("22", interval=0.02)
    time.sleep(0.3)
    pyautogui.press("enter")
    time.sleep(0.3)
    pyautogui.click(1073, 430)  #Área em branco
    time.sleep(0.3)

    while True:
        try:
            pyautogui.click('images\\click_icons\\pdf_icons\\click_edit_page\\click_edit_page.png')   #PDF Phantom - Seleciona o número de páginas
            break
        except TypeError:
            print("Procurando 'Página'...")

    time.sleep(0.3)
    pyautogui.write(str(n_pages), interval=0.05)
    time.sleep(0.3)
    pyautogui.press('esc')
    time.sleep(0.3)
    pyautogui.press("enter")
    time.sleep(0.3)

    while True:
        try:
            pyautogui.click('images\click_icons\pdf_icons\click_edit\\click_edit.png')    #PDF - Edit
            break
        except TypeError:
            try:
                pyautogui.click('images\\click_icons\\pdf_icons\\click_edit\\click_edit_2.png')    #PDF - Edit
                break
            except TypeError:
                try:
                    pyautogui.click('images\\click_icons\\pdf_icons\\click_edit\\click_edit_3.png')    #PDF - Edit
                    break
                except TypeError:
                    print("Procurando 'Edit'...")

    while True:
        try:
            pyautogui.click('images\\click_icons\\pdf_icons\\click_edit_text\\click_edit_text.png')    #PDF - Edit Text
            break
        except TypeError:
            try:
                pyautogui.click('images\\click_icons\\pdf_icons\\click_edit_text\\click_edit_text_2.png')    #PDF - Edit Text
                break
            except TypeError:
                print("Procurando 'Edit Text'...")

    while True:
        try:
            pyautogui.click('images\\click_icons\\pdf_icons\\click_assinaturas\\click_assinaturas.png')    #PDF Phantom - Assinaturas
            break
        except TypeError:
            print("Procurando ícone 'Assinaturas'...")

    time.sleep(0.3)
    pyautogui.hotkey("home")
    time.sleep(0.3)
    pyautogui.press('right')
    time.sleep(0.3)
    pyautogui.press('backspace')
    time.sleep(0.3)
    pyautogui.write(str(i+5), interval=0.02)
    time.sleep(0.3)
    pyautogui.click(1073, 430)  #Área em branco
    time.sleep(0.5)

    while True:
        try:
            pyautogui.click('images\\click_icons\\pdf_icons\\click_edit_page\\click_edit_page.png')   #PDF Phantom - Seleciona o número de páginas
            break
        except TypeError:
            print("Procurando 'Página'...")

    pyautogui.write(str(n_pages-1))
    time.sleep(0.5)
    pyautogui.press('esc')
    time.sleep(0.5)
    pyautogui.press('enter')
    time.sleep(0.5)

    while True:
        try:
            pyautogui.click('images\\click_icons\\pdf_icons\\click_comentarios\\click_comentarios.png')   #Seleciona o número de páginas
            break
        except TypeError:
            print("Procurando ícone 'Comentarios'")

    time.sleep(0.5)
    pyautogui.hotkey("home")
    time.sleep(0.3)
    pyautogui.press('right')
    time.sleep(0.3)
    pyautogui.press('backspace')
    time.sleep(0.3)
    pyautogui.write(str(i+3), interval=0.02)
    time.sleep(0.3)
    pyautogui.click(1073, 430)  #Área em branco
    time.sleep(0.5)
    while True:
        try:
            pyautogui.click('images\\click_icons\\pdf_icons\\click_conclusao\\click_conclusao.png')   #Seleciona o número de páginas
            break
        except TypeError:
            try:
                pyautogui.click('images\\click_icons\\pdf_icons\\click_conclusao\\click_conclusao_2.png')   #Seleciona o número de páginas
                break
            except TypeError:
                print("Procurando ícone 'Conclusao'")
                print('Tentando novamente...')
    time.sleep(0.5)
    pyautogui.hotkey("home")
    time.sleep(0.3)
    pyautogui.press('right')
    time.sleep(0.3)
    pyautogui.press('backspace')
    time.sleep(0.3)
    pyautogui.write(str(i+4), interval=0.02)
    time.sleep(0.3)
    pyautogui.click(1073, 430)  #Área em branco
    time.sleep(0.5)
    pyautogui.alert("Sumário criado com sucesso!", "Notificação", "Continuar")

def graphics():
    tasks.delete_all_files(paths.graphs_dir)
    for n in range(tasks.count_rows(1)):
        cells_list = []
        values_dict = {}
        min_1 = []
        min_10 = []

        for i in range(tasks.count_rows(1)):
            for j in collumns:
                cells_list.append(j+str(n+3))
        
        cont = 0
        for field in field_list:
            values_dict[field] = cells_list[cont]
            cont+=1
        
        graph_name = dados_coletados[values_dict["(equipamento)"]].value

        time.sleep(1)
        for i in values_list_1_min:
            min_1.append(int(dados_coletados[values_dict[i]].value))
        
        for i in values_list_10_min:
            min_10.append(int(dados_coletados[values_dict[i]].value))
        
        labels = ['Fase U/V', 'Fase V/W', 'Fase W/U', 'Fase U/M', 'FaseV/M', 'Fase W/M']

        x = np.arange(len(labels))  # the label locations
        width = 0.35  # the width of the bars

        fig, ax = plt.subplots()
        rects1 = ax.bar(x - width/2, min_1, width, label='1 min')
        rects2 = ax.bar(x + width/2, min_10, width, label='10 min')

        # Add some text for labels, title and custom x-axis tick labels, etc.
        ax.set_ylabel('Resistência (MegaOhm)')
        ax.set_xticks(x, labels)
        ax.legend()

        ax.bar_label(rects1, padding=2)
        ax.bar_label(rects2, padding=2)

        fig.tight_layout()

        print(f"Gráfico '{graph_name}' criado com sucesso.")

        plt.savefig(f'graphs\\{graph_name}')



def insert_graphic():
    locale.setlocale(locale.LC_ALL, 'eng_SG')
    paths_list = []
    n_page = []
    graphics_dir = {}
    graphics_name = []
    model_path = f"{os.getcwd()}\\graphs\\"
    cont = 0
    for file in os.scandir(f"{os.getcwd()}\\graphs"):
        while True:
            try:
                page = int(pyautogui.prompt(text=(f"Informe a página onde o componente '{file.name.split('.')[0]}' está localizado."), title='Gráficos' , default=''))
                print(f"Informe a página onde o componente '{file.name.split('.')[0]}' está localizado: ")
                n_page.append(page)
                page = None
                break
            except TypeError:
                page = None
                pyautogui.alert('O valor informado está incorreto. Tente novamente.', 'Dado inválido', 'Inserir novamente')
                continue
            except ValueError:
                page = None
                pyautogui.alert('O valor informado está incorreto. Tente novamente.', 'Dado inválido', 'Inserir novamente')
                continue
    locale.setlocale(locale.LC_ALL, '')
    tasks.zoom_adjust(zoom_montar_relatorio)
    time.sleep(0.5)
    pyautogui.click(239, 402)   #Área em branco
    time.sleep(0.5)
    pyautogui.press('home')
    for file in os.scandir(f"{os.getcwd()}\\graphs"):
        graphics_name.append(file.name.split('.')[0])
    for file in os.scandir(f"{os.getcwd()}\\graphs"):
        paths_list.append(f"{model_path}{file.name}")
    for path in paths_list:
        graphics_dir[graphics_name[cont]] = path
        cont+=1
    while True:
        try:
            pyautogui.click('images\click_icons\pdf_icons\click_edit\\click_edit.png')    #PDF - Home
            break
        except TypeError:
            try:
                pyautogui.click('images\\click_icons\\pdf_icons\\click_edit\\click_edit_2.png')    #PDF - Home
                break
            except TypeError:
                try:
                    pyautogui.click('images\\click_icons\\pdf_icons\\click_edit\\click_edit_3.png')    #PDF - Home
                    break
                except TypeError:
                    print("Procurando ícone 'Edit'")
                    print('Tentando novamente...')
    while True:
        try:
            pyautogui.click('images\\click_icons\\pdf_icons\\click_edit_object\\click_edit_object.png')    #View -> Uma Página
            break
        except TypeError:
            print("Procurando ícone 'Edit Object'")
            print('Tentando novamente...')
    cont=0
    for k, v in graphics_dir.items():
        while True:
            try:
                pyautogui.click('images\\click_icons\\pdf_icons\\click_edit_page\\click_edit_page.png')   #Seleciona o número de páginas
                break
            except TypeError:
                print("Procurando ícone 'Actual page'")
                print('Tentando novamente...')
        time.sleep(0.5)
        pyautogui.write(str(n_page[cont]), interval=0.05)
        time.sleep(0.5)
        pyautogui.press('esc')
        time.sleep(0.5)
        pyautogui.press('enter')
        while True:
            try:
                pyautogui.click('images\\click_icons\\pdf_icons\\click_add_image\\click_add_image.png')    #View -> Uma Página
                break
            except TypeError:
                print("Procurando ícone 'Add Image'")
                print('Tentando novamente...')
        while True:
            try:
                pyautogui.click('images\\click_icons\\pdf_icons\\click_import_image\\click_import_image.png')    #View -> Uma Página
                break
            except TypeError:
                print("Procurando ícone 'Import Image'")
                print('Tentando novamente...')
        while True:
            try:
                pyautogui.click('images\\click_icons\\pdf_icons\\click_edit_path_2\\click_edit_path.png')    #View -> Uma Página
                break
            except TypeError:
                print("Procurando ícone 'Edit Path'")
                print('Tentando novamente...')
        time.sleep(0.4)
        pyautogui.write(str(model_path), interval=0.01)
        time.sleep(0.4)
        pyautogui.press('enter')
        while True:
            try:
                pyautogui.click('images\\click_icons\\pdf_icons\\click_nome_imagem\\click_nome_imagem.png')    #View -> Uma Página
                break
            except TypeError:
                print("Procurando ícone 'Nome Imagem'")
                print('Tentando novamente...')
        time.sleep(0.4)
        pyautogui.write(f'{k}.png', interval=0.02)
        while True:
            try:
                pyautogui.click("images\\click_icons\\pdf_icons\\click_abir_imagem\\click_abrir_imagem.png")    #View -> Uma Página
                break
            except TypeError:
                print("Procurando ícone 'Abir Imagem'")
                print('Tentando novamente...')
        time.sleep(1)
        pyautogui.rightClick(648, 280)
        time.sleep(0.2)
        pyautogui.press('tab')
        time.sleep(0.2)
        pyautogui.press('enter')
        while True:
            try:
                pyautogui.click("images\\click_icons\\pdf_icons\\click_appearence\\click_appearance.png")
                break
            except TypeError:
                print("Procurando ícone 'Appearence'")
                print('Tentando novamente...')
        while True:
            try:
                pyautogui.click("images\\click_icons\\pdf_icons\\click_general\\click_general.png")
                break
            except TypeError:
                print("Procurando ícone 'General'")
                print('Tentando novamente...')
        time.sleep(0.4)
        pyautogui.hotkey('ctrl', 'a')
        time.sleep(0.5)
        pyautogui.write('1.25', interval=0.01)
        time.sleep(0.5)
        pyautogui.press('tab')
        time.sleep(0.5)
        pyautogui.hotkey('ctrl', 'a')
        time.sleep(0.5)
        pyautogui.write('1.31', interval=0.01)
        time.sleep(0.5)
        pyautogui.press('tab')
        time.sleep(0.5)
        pyautogui.hotkey('ctrl', 'a')
        time.sleep(0.5)
        pyautogui.write('5.77', interval=0.01)
        time.sleep(0.5)
        pyautogui.press('tab')
        time.sleep(0.5)
        pyautogui.hotkey('ctrl', 'a')
        time.sleep(0.5)
        pyautogui.write('2.40', interval=0.01)
        time.sleep(0.5)
        pyautogui.hotkey('alt', 'f4')
        cont+=1
        

def save_report():
    print("Salvando relatório...")
    locale.setlocale(locale.LC_ALL, 'eng_SG')
    while True:
        try:
            ano = int(pyautogui.prompt(text='Informe o ano do relatório\nObs: Utilize números inteiros', title='Informações' , default=''))
            mes = str(pyautogui.prompt(text='Informe o mês do relatório\nObs: Utilize números inteiros', title='Informações' , default=''))
            dia = int(pyautogui.prompt(text='Informe o dia do relatório\nObs: Utilize números inteiros', title='Informações' , default=''))
            embarcacao = str(pyautogui.prompt(text='Informe o nome da embarcação\nObs: Não utilize caracteres especiais ou acentuações', title='Informações' , default=''))
            if (int(mes) < 0 or int(mes) > 12) or (dia < 0 or dia > 31):
                raise TypeError
            else:
                break
        except TypeError:
            ano = None
            mes = None
            dia = None
            embarcacao = None
            pyautogui.alert('O valor informado está incorreto. Tente novamente.', 'Dado inválido', 'Inserir novamente')
        except ValueError:
            ano = None
            mes = None
            dia = None
            embarcacao = None
            pyautogui.alert('O valor informado está incorreto. Tente novamente.', 'Dado inválido', 'Inserir novamente')
    locale.setlocale(locale.LC_ALL, '')
    while True:
        try:
            pyautogui.click("images\\click_icons\\pdf_icons\\click_blank_space\\click_blank_space.png")
            break
        except TypeError:
            print("Procurando ícone 'Blank Space'")
            print('Tentando novamente...')
    pyautogui.hotkey("ctrl", "s")
    while True:
        try:
            pyautogui.click("images\\click_icons\\pdf_icons\\click_editar_caminho_imagem\\click_editar_caminho_imagem.png")
            break
        except TypeError:
            print("Procurando ícone 'Click Editar Imagem'")
            print('Tentando novamente...')
    time.sleep(0.5)
    pyautogui.write(paths.save_path, interval=0.02)
    time.sleep(0.5)
    pyautogui.press("enter")
    time.sleep(0.5)
    for n in range(6):
        pyautogui.press("tab")
        time.sleep(0.5)
    pyautogui.write(f"{ano}{mes}{dia}_ARISOL_{embarcacao}", interval=0.02)
    time.sleep(0.5)
    pyautogui.press("enter")
    time.sleep(0.5)
    cont = 0
    while cont <= 4:
        time.sleep(0.3)
        try:
            pyautogui.click("images\\click_icons\\pdf_icons\\click_substituir_yes\\click_substituir_yes.png")
            break
        except TypeError:
            try:
                pyautogui.click("images\\click_icons\\pdf_icons\\click_substituir_yes\\click_substituir_yes_2.png")
                break
            except TypeError:
                print("Procurando ícone 'Yes'")
                print('Tentando novamente...')
                cont+=1
    pyautogui.hotkey("winleft", "r")
    while True:
        try:
            pyautogui.click("images\\click_icons\\executar\\reference\\reference.png")
            break
        except TypeError:
            print("Procurando ícone 'Executar'")
            print('Tentando novamente...')
    pyautogui.write(paths.save_path, interval=0.02)
    time.sleep(0.5)
    pyautogui.press("enter")
    time.sleep(0.5)
    hoje = datetime.datetime.today()
    hoje_formatado = hoje.strftime("%d/%m/%Y %H:%M:%S")
    pyautogui.alert(f"""RELATÓRIO FINALIZADO!
                    \nEmbarcação: {embarcacao}
                    \nData do relatório: {dia}/{mes}/{ano}
                    \nCriado em: {hoje_formatado}
                    \nDesenvolvedor: João Pedro Rodrigues""", "Notificação")