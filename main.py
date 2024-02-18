import time
import tkinter as tk
from tkinter import ttk
from tkcalendar import DateEntry
from tkinter.filedialog import askopenfilename
import pandas as pd
import numpy as np
from datetime import datetime
import requests

requisicao = requests.get('https://economia.awesomeapi.com.br/json/all')
dict_moedas = requisicao.json()

lista_moedas = list(dict_moedas.keys())


def pegar_cotacao():
    moeda = ComboBox1.get()
    data = calendario_moeda.get()
    ano = data[-4:]
    mes = data[3:5]
    dia = data[:2]
    link = f"https://economia.awesomeapi.com.br/json/daily/{moeda}-BRL/?start_date={ano}{mes}{dia}&end_date={ano}{mes}{dia}"
    requisicao_moeda = requests.get(link)
    cotacao = requisicao_moeda.json()
    valor_moeda = cotacao[0]['bid']
    l3['text'] = f'A cotação da {moeda} no dia {data} é de R$ {valor_moeda}'


def selecionar_arquivo():
    caminho_arquivo = askopenfilename(title="Selecione o Arquivo Moeda")
    var_caminho_arquivo.set(caminho_arquivo)

    if caminho_arquivo:
        l13['text'] = f'Arquivo Selecionado: {caminho_arquivo}'


def atualizar_cotacoes():
    try:
        df = pd.read_excel(var_caminho_arquivo.get())
        moedas = df.iloc[:, 0] # selecionando a lista de moedas
        data_inicial = d14.get()
        data_final = d15.get()

        ano_inicial = data_inicial[-4:]
        mes_inicial = data_inicial[3:5]
        dia_inicial = data_inicial[:2]

        ano_final = data_final[-4:]
        mes_final = data_final[3:5]
        dia_final = data_final[:2]


        for moeda in moedas:
            link = f"https://economia.awesomeapi.com.br/json/daily/{moeda}-BRL/?" \
                f"start_date={ano_inicial}{mes_inicial}{dia_inicial}&" \
                f"end_date={ano_final}{mes_final}{dia_final}"

            requisicao_moeda = requests.get(link)
            cotacao = requisicao_moeda.json()

            for cotacao in cotacao:
                timestamp = int(cotacao['timestamp'])
                bid = float(cotacao['bid'])
                data = datetime.timestamp(pd.to_datetime(timestamp, utc=True))
                data = time.strftime("%d/%m/%Y")
                if data not in df:
                    df[data] = np.nan

                df.loc[df.iloc[:, 0] == moeda, data] = bid
        df.to_excel('Moedas.xlsx')
        l16['text'] = "Arquivo atualizado com sucesso!"

    except:
        l16['text'] = "Selecione um arquivo xlsx no formato correto"


janela = tk.Tk()

janela.title('Ferramenta de cotação de Moedas')

labelCotacaoMoeda = tk.Label(text='Cotação de uma Moeda Específica', borderwidth=2, relief='solid')
labelCotacaoMoeda.grid(row=0,column=0,padx=10, pady=10, sticky='nswe', columnspan=3)

l1 = tk.Label(text='Selecione a moeda', anchor='w')
l1.grid(row=1, column=0, padx=10, pady=10, sticky='nswe', columnspan=2)

ComboBox1 = ttk.Combobox(values=lista_moedas)
ComboBox1.grid(row=1, column=2, padx=10, pady=10, sticky='nswe')

l2 = tk.Label(text='Selecione o dia da cotação', anchor='w')
l2.grid(row=2, column=0, padx=10, pady=10, sticky='nswe', columnspan=2)

calendario_moeda = DateEntry(year=2020, locale='pt_br')
calendario_moeda.grid(row=2, column=2, padx=10, pady=10, sticky='nswe')

l3 = tk.Label(text='')
l3.grid(row=3,column=0,columnspan=2, padx=10,pady=10,sticky='nswe')



button1 = tk.Button(text='Pegar Cotação', command=pegar_cotacao)
button1.grid(row=3, column=2, padx=10,pady=10, sticky='nswe')



# Cotação de várias moedas

labelCotacaoVarias = tk.Label(text='Cotação de Multiplas Moedas Específica', borderwidth=2, relief='solid')
labelCotacaoVarias.grid(row=4,column=0, padx=10, pady=10, sticky='nswe', columnspan=3)

l12 = tk.Label(text="Selecione um arquivo em excel com as moedas na coluna A")
l12.grid(row=5, column=0, padx=10, pady=10, columnspan=2, sticky='nswe')

var_caminho_arquivo = tk.StringVar()

button2 = tk.Button(text="Clique para selecionar", command=selecionar_arquivo)
button2.grid(row=5, column=2, padx=10, pady=10,sticky='nswe')

l13 = tk.Label(text="Nenhum arquivo selecionado", anchor='e')
l13.grid(row=6, column=2,columnspan=3, padx=10,pady=10,sticky='nswe')

l14 = tk.Label(text='Data inicial')
l14.grid(row=7, column=0, padx=10, pady=10, sticky='nswe')

l15 = tk.Label(text='Data final')
l15.grid(row=8, column=0, padx=10, pady=10, sticky='nswe')

d14 = DateEntry(year=2024, locale='pt_br')
d14.grid(row=7, column=1, padx=10, pady=10, sticky='nswe')

d15 = DateEntry(year=2024, locale='pt_br')
d15.grid(row=8, column=1, padx=10, pady=10, sticky='nswe')


button3 = tk.Button(text='Atualizar Cotações', command=atualizar_cotacoes)
button3.grid(row=9,column=0,padx=10,pady=10,sticky='nswe')


l16 = tk.Label(text='')
l16.grid(row=9, column=1, padx=10, pady=10, columnspan=2, sticky='nswe')


button4 = tk.Button(text='Fechar', command=janela.quit)
button4.grid(row=10, column=2, padx=10, pady=10, sticky='nswe')



janela.mainloop()