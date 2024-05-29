import pandas as pd
import pyautogui as gui
import pyperclip
import time

banco_de_dados = pd.read_excel("bancoDados.xlsx")
print(banco_de_dados)

gui.press("win")
time.sleep(1.0)
gui.write("opera")
time.sleep(1.0)
gui.press("enter")
time.sleep(5.0)
gui.write("https://mail.google.com/mail/u/0/#inbox")
time.sleep(0.7)
gui.press("enter")
time.sleep(5.0)

for index, row in banco_de_dados.iterrows():
    email = row['e-mail']
    nome = row['Nome']
    empresa = row['Empresa']

    mensagem = f"""Bom dia {nome}, tudo bem?
Sou André, do comercial da AdsVantage. Peguei seu e-mail através do seu site.
Vimos que a {empresa} possui inúmeros procedimentos em seu portfólio, além de ser bem consolidada no mercado e nos perguntamos se vocês estão com dificuldades com a aquisição de clientes através do marketing digital.
Nós, da AdsVantage somos especialistas em clínicas de estética, ajudando clínicas como a sua a se destacar da concorrência por meio de tráfego pago.
Podemos marcar uma ligação amanhã para detalhar um pouco mais?
Abraços!"""

    gui.press("c")
    time.sleep(3.0)

    pyperclip.copy(email)
    gui.hotkey('ctrl', 'v')
    time.sleep(1.0)
    gui.press("tab")
    time.sleep(2.0)

    assunto = f"Transforme Cliques em Clientes: Descubra as vantagens do Tráfego Pago para Clínicas de Estética, faça a {empresa} voar"
    pyperclip.copy(assunto)
    gui.hotkey('ctrl', 'v')
    time.sleep(1.0)
    gui.press("tab")

    pyperclip.copy(mensagem)
    gui.hotkey('ctrl', 'v')
    time.sleep(1.0)

    gui.hotkey('ctrl', 'enter')
    time.sleep(3.0)

    time.sleep(2.0)
