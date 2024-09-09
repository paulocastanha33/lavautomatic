import sys
import pandas as pd
import pywhatkit as kit
from datetime import datetime
from PyQt5 import QtWidgets
from ui_janela import Ui_MainWIndow  
from PyQt5.QtCore import Qt
import pyautogui
import random

class MainWindow(QtWidgets.QMainWindow):
    def __init__(self):
        super(MainWindow, self).__init__()
        self.ui = Ui_MainWIndow()
        self.ui.setupUi(self)
        # Defina a transparência da janela
        self.setAttribute(Qt.WA_TranslucentBackground)
        # Defina a opacidade da janela (ajuste conforme necessário)
        self.setWindowOpacity(0.9)

        # Conectar os botões às funções
        self.ui.loadButton.clicked.connect(self.load_excel)
        self.ui.sendButton.clicked.connect(self.send_messages)

        self.df = None

        # Lista de palavras relacionadas com lavanderia
        self.laundry_words = ["SABÃO", "AMACIANTE", "DETERGENTE", "ROUPA", "LAVAGEM", "CESTO", "SECAGEM", "MÁQUINA"]

    def load_excel(self):
        file_path, _ = QtWidgets.QFileDialog.getOpenFileName(self, "Selecionar arquivo Excel", "", "Excel Files (*.xlsx)")
        if file_path:
            self.df = pd.read_excel(file_path)
            self.ui.statusText.setText(f"Arquivo carregado: {file_path}")

    def format_phone_number(self, number, country_code="+55"):
        if not number.startswith('+'):
            number = country_code + number.lstrip('0')
        return number

    def send_whatsapp_message(self, phone_number, message):
        now = datetime.now()
        kit.sendwhatmsg(phone_number, message, now.hour, now.minute + 1)
        # Aguarde alguns segundos para o pywhatkit abrir o WhatsApp Web e digitar a mensagem
        pyautogui.sleep(10)
        pyautogui.press('enter')
        pyautogui.hotkey('ctrl', 'w')


    def send_confirmation_message(self, emitent_number, keyword):
        now = datetime.now()
        confirmation_message = f"Sua mensagem foi enviada com sucesso usando a palavra-chave: {keyword}."
        self.send_whatsapp_message(emitent_number, confirmation_message)
        pyautogui.press('enter')

    def send_messages(self):
        if self.df is None:
            self.ui.statusText.setText("Nenhum arquivo carregado.")
            return

        if 'numero_de_telefone' not in self.df.columns or 'quantidade_de_compras' not in self.df.columns:
            self.ui.statusText.setText("A planilha deve conter as colunas 'numero_de_telefone' e 'quantidade_de_compras'.")
            return

        self.df['numero_de_telefone'] = self.df['numero_de_telefone'].astype(str)
        self.df['quantidade_de_compras'] = self.df['quantidade_de_compras'].astype(int)
        self.df['numero_de_telefone'] = self.df['numero_de_telefone'].apply(self.format_phone_number)
        filtered_df = self.df[self.df['quantidade_de_compras'] >= 10]

        for _, row in filtered_df.iterrows():
            phone_number = row['numero_de_telefone']
            # Escolha uma palavra aleatória da lista
            random_word = random.choice(self.laundry_words)
            message = f"Parabéns você acabou de ganhar um desconto especial! Use a palavra-chave: {random_word} para ativar seu desconto."
            self.send_whatsapp_message(phone_number, message)

            # Enviar uma mensagem de confirmação para o emitente (número do WhatsApp do emitente deve ser adicionado na planilha ou definido)
            emitent_number = '+5548984320047'  # Substitua pelo número do emitente
            self.send_confirmation_message(emitent_number, random_word)

            self.ui.statusText.append(f"Mensagem enviada para {phone_number}")

        self.ui.statusText.append("Todas as mensagens foram enviadas.")

if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    mainWin = MainWindow()
    mainWin.show()
    sys.exit(app.exec_())
