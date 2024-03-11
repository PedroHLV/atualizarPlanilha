import os
import openpyxl
import pyautogui
import time
import win32com.client
from functions.disparo_email import enviar_email
from functions.atualizar_planilha import abrir_atualizar_salvar_fechar_excel

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
    