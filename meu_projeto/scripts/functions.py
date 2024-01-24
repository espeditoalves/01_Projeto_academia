from datetime import datetime

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import seaborn as sns


def cria_grafico(base_dados, nome, data):
    sns.set(style='whitegrid')
    plt.figure(figsize=(10, 6))

    cores_personalizadas = {'-': 'black', 'Pago': 'green', 'Falha': 'red'}
    # Criando o gráfico de barras
    sns.countplot(
        x='Mes',
        hue='Status',
        data=base_dados[base_dados['Nome'] == f'{nome}'],
        palette=cores_personalizadas,
    )
    plt.xticks(rotation=45)
    # Adicionando rótulos e título
    plt.xlabel('Mês')
    plt.ylabel('Contagem')
    plt.title(f'Contagem de dias Pagos e dias de Falha por Mês - {nome}')
    # Salvando o gráfico como imagem (por exemplo, formato PNG)
    path = 'C:/Users/esped/OneDrive/Documentos/_repositorios_/Projetos/01_Projeto_academia/meu_projeto/graficos/'
    plt.savefig(f'{path}{data}_grafico_seaborn_{nome}.png')


def analisa_excel():
    caminho = 'C:/Users/esped/OneDrive/2.Contas_casa/2024_Treinos.xlsx'
    colunas = ['Data', 'Janaina', 'Espedito']
    df_tabela = pd.read_excel(io=caminho, sheet_name='Base', usecols=colunas)
    data_atual = datetime.now()
    data_formatada_str = data_atual.strftime('%Y-%m-%d')
    data_formatada = np.datetime64(data_formatada_str)

    df_tabela_filtrada = df_tabela[df_tabela['Data'] <= data_formatada]

    # Crie as colunas: Dia_da_semana, Mes, e Ano
    df_tabela_filtrada = df_tabela_filtrada.copy()
    df_tabela_filtrada.loc[:, 'Dia_Semana'] = df_tabela_filtrada[
        'Data'
    ].dt.day_name()
    df_tabela_filtrada.loc[:, 'Mes'] = df_tabela_filtrada[
        'Data'
    ].dt.month_name()
    df_tabela_filtrada.loc[:, 'Ano'] = df_tabela_filtrada['Data'].dt.year

    # Cria 2 tabelas temporarias
    df_janaina = df_tabela_filtrada.drop(columns=['Espedito']).rename(
        columns={'Janaina': 'Status'}
    )
    df_janaina['Nome'] = 'Janaina'
    df_espedito = df_tabela_filtrada.drop(columns=['Janaina']).rename(
        columns={'Espedito': 'Status'}
    )
    df_espedito['Nome'] = 'Espedito'

    # Concatenando os DataFrames ao longo das linhas (eixo 0)
    base_dados = pd.concat([df_janaina, df_espedito]).sort_values(by=['Data'])

    cria_grafico(base_dados=base_dados, nome='Espedito', data=data_formatada)
    cria_grafico(base_dados=base_dados, nome='Janaina', data=data_formatada)

    # Contando os registros com base no Mês
    contagem_mes = (
        base_dados.groupby(['Mes', 'Nome', 'Status'])
        .size()
        .reset_index(name='Contagem_mes')
    )

    return contagem_mes, data_formatada
