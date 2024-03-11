# Atualizador e Enviador de Planilha por E-mail



Este script Python foi desenvolvido para automatizar a atualização de uma planilha do Excel e enviá-la por e-mail. Ele é projetado para ser usado no Agendador de Tarefas do Windows, permitindo que o script seja executado periodicamente.

- Funcionalidades
Abre uma planilha do Excel.
Atualiza a planilha.
Salva as alterações.
Fecha a planilha.
Envia a planilha atualizada por e-mail como anexo.


- Pré-requisitos
Python 3.x instalado no sistema.
Bibliotecas Python: openpyxl, pyautogui, pywin32.
Conta de e-mail com suporte a SMTP para enviar e-mails.


- Configuração
Instale as bibliotecas necessárias executando o seguinte comando:
pip install openpyxl pyautogui pywin32

Configure as informações do servidor SMTP no arquivo enviar_email.py:
smtp_server: endereço do servidor SMTP.
smtp_port: porta do servidor SMTP.
smtp_username: nome de usuário para autenticação no servidor SMTP.
smtp_password: senha para autenticação no servidor SMTP.
No arquivo app.py, ajuste o caminho para a planilha do Excel que deseja atualizar e enviar por e-mail:
arquivo_excel = "Caminho\\para\\a\\planilha.xlsx"
Configure o Agendador de Tarefas do Windows para executar o script app.py periodicamente.

- Uso
Execute o script app.py para abrir, atualizar e salvar a planilha do Excel. O script também enviará a planilha atualizada por e-mail.

- Observações
Certifique-se de que o ambiente em que o script é executado está desbloqueado para permitir a interação com a interface do usuário, especialmente se estiver usando a biblioteca pyautogui.
Verifique as configurações de segurança da conta de e-mail para permitir o acesso SMTP.
Contribuição
Contribuições são bem-vindas! Sinta-se à vontade para enviar pull requests com melhorias, correções de bugs ou novos recursos.

- Licença
# Este projeto está licenciado sob a Licença MIT.
