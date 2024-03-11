import os
import win32com.client
import pyautogui
import time

def abrir_atualizar_salvar_fechar_excel(file_path):
    # Verifica se o arquivo existe
    if os.path.exists(file_path):
        file_name = os.path.basename(file_path)
        print(f"O arquivo {file_name} foi encontrado.")
    else:
        print(f"O arquivo {file_path} não foi encontrado.")
        return
    
    # Abre o arquivo Excel em primeiro plano
    xlapp = win32com.client.Dispatch("Excel.Application")
    xlapp.Visible = True
    xlapp.DisplayAlerts = False
    print("Abrindo o arquivo Excel...")
    wb = xlapp.Workbooks.Open(file_path)
    time.sleep(4)

    # Atualiza o Excel (pressiona a tecla F9)
    print("Atualizando o Excel...")
    pyautogui.press('f9')
    time.sleep(2)

    # Salva o arquivo Excel
    print("Salvando as alterações...")
    wb.Save()
    time.sleep(2)

    # Fecha o arquivo Excel
    # print("Fechando o arquivo Excel...")
    # wb.Close()