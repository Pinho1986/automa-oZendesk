#VERSÃO ATUALIZADA A PRIMEIRA VERSÃO ESTÁ BARRANDO NO VERIFIADOR DE BOT DO CLOUDFAIRE - 
#ESSA VERSÃO USA O undetected_chromedriveR QUE CONSEGUE PASSSAR PELA VERIFICAÇÃO 
#-------------------------------V E R S Ã O 03 --------------------------------------
#------------------------COM ESCOLHA DE INTERVALO DE DATA ----------------------------
#------------------------CODIGO PASSANDO PELO DECTECTOR DE BOT------------------------ 

import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from time import sleep
import pyautogui
import pandas as pd
from datetime import datetime
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.wait import WebDriverWait
from bs4 import BeautifulSoup
import re
import tkinter as tk
from tkcalendar import Calendar, DateEntry  # Importar DateEntry para seleção de intervalo

def process_tickets(start_date, end_date):
    navegador = uc.Chrome()

    # Preparando o site
    navegador.get("SITE")
    pyautogui.press('enter')
    pyautogui.sleep(1)

    pyautogui.write('LOG')
    sleep(0.8)
    pyautogui.write('LOG')
    pyautogui.press('enter')
    sleep(1)
    pyautogui.write('PASSWORD')
    sleep(0.5)
    pyautogui.write('PASSWORD')
    sleep(1)
    pyautogui.press('enter')
    pyautogui.press('enter')

    navegador.get("https://vsdata.zendesk.com/agent/filters/16152570698765")

    sleep(1)

    tabela = pd.read_excel('AtualizacaoPlanilhas.xlsx')
    tabela.loc[tabela['TICKETS'].notna() & tabela['PROCESSAMENTO'].isnull(), 'PROCESSAMENTO'] = 'Pendente'

    for linha in tabela.itertuples(index=True, name='TICKETS'):
        if linha.PROCESSAMENTO == 'Pendente':
            num_tickets = int(linha.TICKETS)

            navegador.get(f'https://vsdata.zendesk.com/tickets/{num_tickets}/print')
            page_source = navegador.page_source
            soup = BeautifulSoup(page_source, 'html.parser')

            status = WebDriverWait(navegador, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="fields"]/div[2]/p')))
            status_texto = status.text

            mensagens = []

            interacoes = soup.find_all('div', class_='comment')

            for interacao in interacoes:
                # Extrair a data e hora da interação
                time_tag = interacao.find('time', class_='date')
                if time_tag:
                    datetime_str = re.search(r'\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}', time_tag['datetime']).group()
                    # Converter a data e hora do texto para o formato datetime
                    data_hora_texto = datetime.strptime(datetime_str, '%Y-%m-%d %H:%M:%S')

                # Verificar se a data da interação está dentro do intervalo selecionado pelo usuário
                if start_date <= data_hora_texto.date() <= end_date:
                    # Extrair o texto da interação
                    texto_msg = interacao.find('div', class_='zd-comment').text

                    # Extrair todos os autores da interação
                    autores_tags = interacao.find_all('div', class_='mast')
                    autores_interacao = [autor.text.strip() for autor in autores_tags if autor.text.strip()]

                    # Verificar se há autores disponíveis
                    if autores_interacao:
                        # Adicionar a data, hora e autores ao texto da mensagem
                        autores_str = ', '.join(autores_interacao)
                        texto_com_info = f"{data_hora_texto.strftime('%Y-%m-%d %H:%M:%S')} - Autor : {autores_str} -\n {texto_msg}\n"
                    else:
                        texto_com_info = f"{data_hora_texto.strftime('%Y-%m-%d %H:%M:%S')} - Autor não disponível - {texto_msg}"

                    print(f"Ticket {num_tickets} - {texto_com_info}")
                    mensagens.append(texto_com_info)

            tabela.loc[linha.Index, 'INTERAÇÕES'] = '\n'.join(mensagens)
            tabela.loc[linha.Index, 'STATUS'] = status_texto
            tabela.loc[linha.Index, 'PROCESSAMENTO'] = 'FEITO'

            tabela.to_excel('AtualizacaoPlanilhas.xlsx', index=False)
    navegador.quit()
def centralizar_janela(janela, largura, altura):
    # Obter as dimensões da tela
    largura_tela = janela.winfo_screenwidth()
    altura_tela = janela.winfo_screenheight()

    # Calcular as coordenadas para centralizar a janela
    x = (largura_tela // 2) - (largura // 2)
    y = (altura_tela // 2) - (altura // 2)

    # Definir a geometria da janela para centralizá-la
    janela.geometry(f"{largura}x{altura}+{x}+{y}")    

def get_date_input():
    Janela_Calendario = tk.Tk()
    Janela_Calendario.withdraw()  # Esconder a janela principal
    Janela_Calendario.title("Vs Data - Automação Tickets Zendesk")
             
    top = tk.Toplevel(Janela_Calendario)
    top.geometry("470x550")
    centralizar_janela(top, 470, 550)
    
    #CONFIGURAÇÕES DO CALENDARIO 
    calendario_Inicio = Calendar(top, selectmode="day", year=datetime.now().year, month=datetime.now().month, day=datetime.now().day,
                     background='white', foreground='black', bordercolor='blue', headersbackground='blue', headersforeground='white',
                     selectbackground='lightblue', selectforeground='black')
    calendario_Inicio.pack()  # Coloca o primeiro widget

    calendario_Final = Calendar(top, selectmode="day", year=datetime.now().year, month=datetime.now().month, day=datetime.now().day,
                   background='white', foreground='black', bordercolor='blue', headersbackground='blue', headersforeground='white',
                   selectbackground='lightblue', selectforeground='black')
    calendario_Final.pack()  # Coloca o segundo widget

    def get_dates():
        data_inicio = calendario_Inicio.selection_get()
        data_final = calendario_Final.selection_get()
        Janela_Calendario.quit()
        return data_inicio, data_final

    # Adicionando espaço entre os calendários e o botão
    tk.Label(top, text="").pack()  # Adiciona um espaço em branco entre os widgets

    # Adicionando o botão abaixo dos calendários
    botao = tk.Button(top, text="Selecionar Intervalo de Datas", command=get_dates, bg='blue', fg='white', font=('Arial', 10, 'bold'))
    botao.pack()  # Coloca o botão abaixo dos calendários

    Janela_Calendario.mainloop()

    return get_dates()


def main():
    start_date, end_date = get_date_input()
    if start_date and end_date:
        process_tickets(start_date, end_date)

if __name__ == "__main__":
    main()
