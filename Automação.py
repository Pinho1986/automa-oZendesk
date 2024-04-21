#-------------------------CALENDARIO CONFIGURADO ----------------------------------
#-------------------------COM INTERVALO DE DATAS ----------------------------------

from selenium import webdriver
from selenium.webdriver.common.by import By
from time import sleep
import pyautogui
import pandas as pd
from datetime import datetime
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.chrome.options import Options
from bs4 import BeautifulSoup
import re
import tkinter as tk
from tkcalendar import Calendar 

def process_tickets(data_Inicio, data_Final):
    navegador = webdriver.Chrome()

    # Preparando o site
    navegador.get("https://vsdata.zendesk.com")

    pyautogui.sleep(1)
    
    # Fazendo validação login e senha 
    pyautogui.write('SEU LOGIN')
    pyautogui.press('enter')
    pyautogui.write('SUA SENHA')
    pyautogui.press('enter')
    pyautogui.press('enter')

    navegador.get("https://vsdata.zendesk.com")

    sleep(1)
    
    # Planilha excel 
    tabela = pd.read_excel('AtualizacaoPlanilhas.xlsx')
    tabela.loc[tabela['TICKETS'].notna() & tabela['ETAPA1'].isnull(), 'ETAPA1'] = 'Pendente'

    for linha in tabela.itertuples(index=True, name='TICKETS'):
        if linha.ETAPA1 == 'Pendente':
            num_tickets = int(linha.TICKETS)

            #Passando nos tickets 
            navegador.get(f'https://vsdata.zendesk.com')
            page_source = navegador.page_source
            soup = BeautifulSoup(page_source, 'html.parser')

        
            status = WebDriverWait(navegador, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="fields"]/div[2]/p')))
            status_texto = status.text

            mensagens = []

            interacoes = soup.find_all('div', class_='comment')

            for interacao in interacoes:
                # Extrair a data e hora da interação
                hora_Tag = interacao.find('time', class_='date')
                if hora_Tag:
                    datetime_str = re.search(r'\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}', hora_Tag['datetime']).group()
                    # Converter a data e hora do texto para o formato datetime
                    data_hora_texto = datetime.strptime(datetime_str, '%Y-%m-%d %H:%M:%S')

                # Verificar se a data da interação está dentro do intervalo selecionado pelo usuário
                if data_Inicio <= data_hora_texto.date() <= data_Final:
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
            tabela.loc[linha.Index, 'ETAPA1'] = 'FEITO'

            tabela.to_excel('AtualizacaoPlanilhas.xlsx', index=False)
    navegador.quit()

def get_date_input():
    Janela_Calendario = tk.Tk()
    Janela_Calendario.withdraw()  # Esconder a janela principal
    Janela_Calendario.title("Vs Data")
             
    top = tk.Toplevel(Janela_Calendario)
    
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
    data_inicio, data_inicio = get_date_input()
    if data_inicio and data_inicio:
        process_tickets(data_inicio, data_inicio)

if __name__ == "__main__":
    main()


