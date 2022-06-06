from functions import (
    sumario,
    unir_pdf,
    enumerar_paginas,
    informacoes,
    montar_relatorio,
    save_report,
    add_border,
    save_pre_report,
    graphics,
    insert_graphic
)

import pyautogui

if __name__ == "__main__":
    montar_relatorio()
    pyautogui.alert('Montar relatório concluído')
    save_pre_report()
    pyautogui.alert('Save pre report concluído')
    unir_pdf()
    pyautogui.alert('Unir pdf concluído')
    add_border()
    pyautogui.alert('Add border concluído')
    enumerar_paginas()
    pyautogui.alert('Enumerar paginas concluído')
    informacoes()
    pyautogui.alert('Informações concluído')
    sumario()
    pyautogui.alert('Sumário concluído')
    graphics()
    pyautogui.alert('Gráficos criados com sucesso')
    insert_graphic()
    pyautogui.alert('Gráficos inseridos com sucesso')
    save_report()