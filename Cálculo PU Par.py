##########################################################################
# Importando as bibliotecas necessárias
##########################################################################

import matplotlib, json, openpyxl
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
from datetime import datetime
from tkinter import filedialog
from tkinter.filedialog import askopenfilename
from tkinter import Tk

##########################################################################
# Inputs
##########################################################################

# Deverá ser calculado PU ou Taxa?
taxa_pu = str(input('Deverá ser calculado PU ou Taxa (preencha com "PU" ou "Taxa"): '))

if taxa_pu == 'PU':
    data_inicio_analise = str(input('Digite a data inicial para plotar o gráfico (formato yyyy-mm-dd): '))
    data_fim_analise = str(input('Digite a data final para plotar o gráfico (formato yyyy-mm-dd): '))

if taxa_pu == 'Taxa':
    pu_operacao = int(input('Digite PU Operação utilizado para encontrar a taxa: '))
    dias_calc_taxa = int(input('Digite número de dias úteis utilizado para encontrar a taxa: '))

##########################################################################
# Importando o JSON
##########################################################################

# Abrir uma caixa de diálogo para importar o arquivo do contrato
def selecionar_arquivo_json():
    Tk().withdraw()
    arquivo_selecionado = askopenfilename(
        title="Selecione o JSON",
        filetypes=[("Arquivos JSON", "*.json"), ("Todos os Arquivos", "*.*")]
    )
    return arquivo_selecionado

# Seleção do arquivo
arquivo_json = selecionar_arquivo_json()

# Abrir o JSON
with open(arquivo_json, 'r', encoding='utf-8') as file:
    dados = json.load(file)

##########################################################################
# Criando variáveis para os dados do contrato
##########################################################################

# Trazendo os dados para calcular o PU Par
vne = dados['emission_price']
dt_emissao = dados['start_date']
taxa_emissao = dados['spread']

# Extraindo o JSON dos fluxos e armazenando em um dataframe
schedules = dados['schedules']
df_fluxos = pd.DataFrame(schedules)   
df_fluxos.rename(columns={'due_date': 'dias_uteis', 'amount': 'perc_amortizacao'}, inplace= True)
df_fluxos['dias_uteis'] =pd.to_datetime(df_fluxos['dias_uteis'])

##########################################################################
# Definindo a data inicial e final do contrato
##########################################################################

# Definindo as datas de início e fim do contrato
data_inicial = datetime.strptime(dt_emissao, "%Y-%m-%d")

# Puxando a data da última amortização do dataframe e transformando de timestamp para datetime
data_final = df_fluxos['dias_uteis'].iloc[-1]
data_final = data_final.to_pydatetime()

##########################################################################
# Importando o arquivo e excluindo finais de semana e feriados
##########################################################################

# Importando o arquivo.
feriados = pd.read_excel("feriados_nacionais.xls")

# O arquivo tinha, ao final, alguns textos. Decidi tratar no código e não na planilha.
feriados['Data'] = pd.to_datetime(feriados['Data'], errors='coerce')
feriados = feriados.dropna(subset=['Data'])

# Convertendo em data.
feriados['Data'] = pd.to_datetime(feriados['Data'], dayfirst=True, errors='coerce')
feriados = feriados.dropna(subset=['Data'])

# Filtrando por data inicial e final.
todas_as_datas = pd.date_range(start=data_inicial, end=data_final)

# Removendo finais de semana e feriados.
du = [
    data for data in todas_as_datas 
    if data.weekday() < 5 and data not in feriados['Data'].values
]

# Consolidando as datas.
df_calculos = pd.DataFrame({'dias_uteis': du})

# Calculando a quantidade de dias úteis entre o início e a data observada
df_calculos['du'] = range(1, len(df_calculos) + 1)

##########################################################################
# Inserindo os percentuais de amortização, somando e calculando o VNA
##########################################################################

# Fazendo o merge do datraframe que contém os fluxos
df_calculos = df_calculos.merge(df_fluxos, how='left', on='dias_uteis')
df_calculos['perc_amortizacao'] = df_calculos['perc_amortizacao'].fillna(0)

# Somando os percentuais de amortização para calcular o VNA
df_calculos['perc_amortizacao_somado'] = df_calculos['perc_amortizacao'].cumsum()

# Calculando o VNA
df_calculos['vna'] = vne - vne * df_calculos['perc_amortizacao_somado']

##########################################################################
# Calculando o PU Par e plotando o gráfico
##########################################################################

# Calcular o preço para cada data usando a fórmula
if taxa_pu == 'PU':
    df_calculos['pu_par'] = df_calculos['vna'] * (1 + taxa_emissao) ** (df_calculos['du'] / 252)

    # Filtrando o dataframe pelas datas inputadas pelo usuário
    df_calculos_filtrado = df_calculos[(df_calculos['dias_uteis'] >= data_inicio_analise) & (df_calculos['dias_uteis'] <= data_fim_analise)]

    # Parametrizando o gráfico
    plt.figure(figsize=(10, 6))
    plt.plot(df_calculos_filtrado['dias_uteis'], df_calculos_filtrado['pu_par'], marker='o', linestyle='-')
    plt.title('Evolução do PU Par ao Longo do Tempo', fontsize=14)
    plt.xlabel('Ano', fontsize=12)
    plt.ylabel('PU Par', fontsize=12)
    plt.grid(True, linestyle='--', alpha=0.4)
    plt.legend(fontsize=10)
    plt.xticks(rotation=45)
    plt.tight_layout()
    plt.show()

##########################################################################
# Calculando a taxa
##########################################################################

if taxa_pu == 'Taxa':
    # Procurando o VNA do ativo, baseado no número de dias úteis informado inicialmente
    vna = df_calculos.loc[df_calculos['du'] == dias_calc_taxa, 'vna'].values[0]

    # Cálculo da taxa
    taxa_emissao = (pu_operacao / vna) ** (252 / dias_calc_taxa) - 1

    print(f"A taxa calculada é: {taxa_emissao:.2f}%")