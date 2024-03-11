import os
import openpyxl
import pyautogui
import time
import win32com.client
from disparo_email import enviar_email

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

if __name__ == "__main__":
    # Caminho para o arquivo Excel
    arquivo_excel = "C:\\Users\\pedro.lopes\\teste.xlsx"
    nome_arquivo = os.path.basename(arquivo_excel)

    # Abrir, atualizar, salvar e fechar o arquivo Excel
    abrir_atualizar_salvar_fechar_excel(arquivo_excel)
    print("Planilha atualizada, salva e fechada com sucesso.")

    #Disparo via Email
    print("Enviando planilha por email...")
    enviar_email(arquivo_excel, nome_arquivo)
    