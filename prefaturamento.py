import re
import warnings
import pandas as pd
from tkinter.filedialog import askopenfilename
warnings.filterwarnings('ignore')


def descricao(row):
    cat2 = row['CATEGORIA OPERACIONAL 2']
    cat3 = row['CATEGORIA OPERACIONAL 3']
    if len(cat3) > 20:
        return cat3
    else:
        return cat2 + ' - ' + cat3

def importa_excel(arquivo):
    colunas_iniciais = ['DR CLIENTE', 'ID DA REQ', 'ID DA WO', 'PRIORIDADE', 'PIB', 'DATA ATRIBUIÇÃO',
                        'DATA CONCLUSÃO', 'NOME DO CLIENTE', 'MCU', 'ORGANIZAÇÃO', 'CATEGORIA OPERACIONAL 2',
                        'CATEGORIA OPERACIONAL 3', 'DESCRIÇÃO', 'DESIGNADO', 'SOLUÇÃO', 'REABERTURA']
    df = pd.read_excel(arquivo, skiprows=7,
                    usecols=colunas_iniciais,
                    sheet_name='QUANTIDADE DE WO RESOLVIDAS POR',
                    dtype={'PIB': object, 'MCU': object})
    return df

def preprocessamento(df):
    # Pre-processamento
    # Remove linhas sem informação.
    df = df.loc[df['ID DA REQ'].notna(), : ]

    # Cria colunas
    # Descrição
    df['Descrição'] = df.apply(descricao, axis=1)
    # Prioridade
    df['Prioridade'] = df.loc[:, 'PRIORIDADE'].replace({'Crítico': '1', 'Alto': '2', 'Médio': '3', 'Baixo': '4'})
    # Informações
    df['Informações'] = None
    df['Informações'] = df['DESCRIÇÃO'].str.extract(r'Descr.*?[:].*? (.*)')
    df.loc[df['Informações'].isnull(), 'Informações'] = df['DESCRIÇÃO'][:250]
    df['Informações'] = df['Informações'].str.strip()
    # Local de Atendimento
    df['Endereço'] = df['DESCRIÇÃO'].str.extract(r'Local de [a|A]tendimento: (.*)')
    df.loc[df['Endereço'].isnull(), 'Endereço'] = 'Endereço não informado'
    df['Endereço'] = df['Endereço'].str.strip()
    df['Telefone'] = df['DESCRIÇÃO'].str.extract(r'Telefone de [c|C]ontato: (.*)')
    df.loc[df['Telefone'].isnull(), 'Telefone'] = 'Telefone não informado'
    df['Telefone'] = df['Telefone'].str.strip()
    df['Localidade'] = df['DESCRIÇÃO'].str.extract(r'Informe [\w]{1,3} locali[\w].*: (.*)')
    df.loc[df['Localidade'].isnull(), 'Localidade'] = df['DR CLIENTE']
    df['Localidade'] = df['Localidade'].str.strip()
    df.loc[:, 'Local de Atendimento'] = (
        df[
            ['Endereço', 'Telefone', 'ORGANIZAÇÃO', 'Localidade']
        ].agg(' / '.join, axis=1)
    )
    # Solução
    df['Solução'] = df['SOLUÇÃO'].str.extract(r'^(.*)\n')
    df['Número da OS do contratado'] = ''
    df['Evento (CONCLUIDO/CANCELADO)'] = ''
    df['Qualidade RAT'] = ''
    df['Observações'] = ''

    renomeia_colunas = {'DR CLIENTE': 'SE',
                        'ID DA REQ': 'ID',
                        'DATA ATRIBUIÇÃO': 'Data da atribuição',
                        'NOME DO CLIENTE': 'Cliente',
                        'ORGANIZAÇÃO': 'Unidade',
                        'DESIGNADO': 'Técnico',
                        'DATA CONCLUSÃO': 'Data de Fechamento',
                        'REABERTURA': 'Reabertura'}
    df.rename(columns=renomeia_colunas, inplace=True)

    colunas_finais = ['SE', 'ID', 'Data da atribuição', 'Prioridade', 'PIB', 'Cliente',
                      'MCU', 'Unidade', 'Descrição', 'Informações', 'Local de Atendimento',
                      'Número da OS do contratado', 'Técnico', 'Data de Fechamento', 
                      'Evento (CONCLUIDO/CANCELADO)', 'Solução', 'Reabertura',
                      'Qualidade RAT', 'Observações']

    return df[colunas_finais]

def exporta_excel(df):
    mes = df['Data da atribuição'].dt.month.mode().values[0]
    ano = df['Data da atribuição'].dt.year.mode().values[0]
    se = df['SE'].mode().values[0].replace('/', '_')
    arquivo = 'Relatorio_Prefaturamento-{}_{}-{}.xls'.format(mes, ano, se)
    writer = pd.ExcelWriter(arquivo, datetime_format='dd/mm/yyyy hh:mm:ss')
    df.to_excel(writer, sheet_name='Pré-Faturamento', index=False)
    writer.close()

def main():
    arquivo = askopenfilename(title = "Seleciona o arquivo Excel.")
    df = importa_excel(arquivo)
    df_processado = preprocessamento(df)
    exporta_excel(df_processado)

if __name__ == '__main__':
    main()