import os
import pandas as pd
import numpy as np
import datetime as dt
from ftplib import FTP
from zipfile import ZipFile
import time
import os
import d6tstack
import pandas as pd
import numpy as np
import datetime as dt
from dateutil.relativedelta import relativedelta
import win32com.client as win32

import xlwings as xl
from git import Repo

import pandas as pd
import numpy as np
import datetime as dt
from ftplib import FTP
from dateutil.relativedelta import relativedelta
import os
from zipfile import ZipFile
import time
from sqlalchemy import create_engine
from itertools import *
import requests

engine = create_engine('postgresql://postgres:v3rd4d31r0@10.202.202.1:5432/db_ghp')

def four_digital_care(data_analise=dt.date.today(),
                      mes='Dezembro',
                      path = r'C:\Users\Utilizador\OneDrive\Grupo HealthPorto\Suporte - General\SO\2020'):

    data_fim=(data_analise + relativedelta(day=31, months=-1)).strftime('%Y%m%d')

    # Aqui rola a montagem do caminho para os .csv da 4DigitalCare... também vale ressaltar que são os argumentos
    # da função em questão.

    dir_path = list(map(lambda x: os.path.join(os.path.join(path, mes), x),
                        os.listdir(os.path.join(path, mes))))

    final_df = pd.DataFrame([])
    df = pd.DataFrame([])

    for i in dir_path:
        try:
                print(i)

                df = pd.read_csv(i, sep=';', encoding='latin-1',
                                 decimal=',', error_bad_lines=False,
                                 warn_bad_lines=False, engine='python')
        except:

            df = pd.read_excel(i)[0].tolist()
            cut_index= df.loc[(df[df.columns[1]] == 'Cd Prd')].index.values
            col_index= [_columns.index(i) for i in _columns if type(i) == str]

            df = df[df.columns[col_index]]

            df.columns = ['cnp', 'nome', 'lab', 'classe', 'quantidade',
                         'preco_venda', 'pvp', 'pvp2', 'pva', 'preco_custo', 'stock']

            df = df.iloc[cut_index[0] + 1:][['cnp', 'quantidade', 'preco_venda',
                                           'preco_custo', 'stock']].assign(data=data_fim,
                                                                           codigo=6335)

        finally:
            if any(['asmpd' in i.lower(), 'matias' in i.lower(),
                    'neves' in i.lower(), 'vaz' in i.lower()]):
                df=df[['CodANF', 'CPR', 'Existencias',
                       'PVP', 'PC', 'V1' if 'neves' in i.lower() else 'V0']].rename(
                    columns={'CodANF':'codigo',
                             'CPR':'cnp',
                             'Existencias':'stock',
                             'PVP':'preco_venda',
                             'PC':'preco_custo',
                             'V1' if 'neves' in i.lower() else 'V0' : 'quantidade'}
                         ).assign(data=data_fim)



        df = df.reindex(sorted(df.columns), axis=1)
        final_df = pd.concat([final_df, df], ignore_index=True, sort=False)

    final_df = final_df[['codigo', 'cnp', 'stock',
                         'preco_custo', 'preco_venda',
                         'data', 'quantidade']]

    return final_df


def farma_lobo(data_analise=dt.date.today(),
               mes='Dezembro',
               path = r'C:\Users\Utilizador\OneDrive\Grupo HealthPorto\Suporte - General\SO\2020'):

    data_fim=(data_analise + relativedelta(day=31, months=-1)).strftime('%Y%m%d')

    # Aqui rola a montagem do caminho para os .xls da Farmácia Lobo... também vale ressaltar que são os argumentos
    # da função em questão.

    dir_path = list(map(lambda x: os.path.join(os.path.join(path, mes), x),
                        os.listdir(os.path.join(path, mes))))
    # Lobo
    for i in [i for i in dir_path if 'lobo' in i.lower()]:
        if 'pvp' in i.lower():
            df_lobo=pd.read_excel(i)
            pc_index = [0, 3, 18, 25]
            cut_index = df_lobo.loc[df_lobo[df_lobo.columns[0]] == 'Codigo'].index.values[0]

            df_lobo = df_lobo.iloc[cut_index + 1:][df_lobo.columns[pc_index]]
            df_lobo.columns = ['cnp', 'preco_venda', 'quantidade', 'stock']
        else:
            df_lobo_2 = pd.read_excel(i)
            pc_index = [0, 3, 18, 25]
            cut_index = df_lobo_2.loc[df_lobo_2[df_lobo_2.columns[0]] == 'Codigo'].index.values[0]

            df_lobo_2 = df_lobo_2.iloc[cut_index + 1:][df_lobo_2.columns[pc_index]]
            df_lobo_2.columns = ['cnp', 'preco_custo', 'quantidade', 'stock']

    df_lobo=df_lobo.merge(df_lobo_2,
                          how='left', on=['cnp', 'quantidade', 'stock']).assign(data=data_fim, codigo=7170)

    return df_lobo

def soft_reis(data=None):

    path = r'C:\Users\Utilizador\Desktop\GHP\Histórico - Sell Out\Softreis'

    if data == None:
        data = dt.datetime.today() + relativedelta(days=-1)
        data = data.strftime("%Y%m%d")

    files = [i for i in os.listdir(path) if data in i and '.zip' in i]
    [ZipFile(os.path.join(path, eachFile)).extractall(path) for eachFile in files]
    cols = ['Existencia Actual', 'TotalPVP', 'TOtalPC', 'TotalVen']

    dates = {'1': 'jan', '2': 'fev', '3': 'mar', '4': 'abr',
             '5': 'mai', '6': 'jun', '7': 'jul', '8': 'ago',
             '9': 'set', '10': 'out', '11': 'nov', '12': 'dez'}

    the_df = pd.DataFrame()

    for eachFile in sorted([i for i in os.listdir(path) if '.csv' in i]):

        data = eachFile.split('_')[-1].replace('.csv', '')
        ano = data[:4]
        mes = str(int(data[4:6]))

        farmacia = eachFile.split('_')[1]

        try:

            primeiro_item = pd.read_csv(os.path.join(path, eachFile), sep='\t', engine='python')
            primeiro_item = primeiro_item.columns[0].split(';')[0]

            new_entrada = pd.read_csv(os.path.join(path, eachFile),
                                      usecols=[primeiro_item] + cols + [dates[mes] + ' ' + ano],
                                      sep=';', encoding='latin-1', engine='python')

            new_entrada.rename(columns={dates[mes] + ' ' + ano: 'quantidade'}, inplace=True)

        except:

            primeiro_item = pd.read_csv(os.path.join(path, eachFile), sep='\t', engine='python')
            primeiro_item = primeiro_item.columns[0].split(';')[0]

            new_entrada = pd.read_csv(os.path.join(path, eachFile),
                                      usecols=[primeiro_item] + cols + [dates[mes].title() + ' ' + ano],
                                      sep=';', encoding='latin-1', engine='python')

            new_entrada.rename(columns={dates[mes].title() + ' ' + ano: 'quantidade'}, inplace=True)

        new_entrada['codigo'] = farmacia
        new_entrada['data'] = dt.datetime.strftime(dt.datetime.strptime(data, '%Y%m%d'), '%Y%m%d')
        new_entrada.rename(columns={primeiro_item: 'cnp', 'Existencia Actual': 'stock', 'TotalPVP': 'preco_venda',
                                    'TOtalPC': 'preco_custo', 'TotalVen': 'total'}, inplace=True)

        new_entrada.preco_venda = new_entrada.preco_venda.apply(lambda x: x.replace(',', '.')).convert_objects(
            convert_numeric=True)
        new_entrada.preco_custo = new_entrada.preco_custo.apply(lambda x: x.replace(',', '.')).convert_objects(
            convert_numeric=True)
        new_entrada.stock = new_entrada.stock.convert_objects(convert_numeric=True)
        new_entrada.total = new_entrada.total.convert_objects(convert_numeric=True)

        new_entrada.preco_venda /= new_entrada.total
        new_entrada.preco_custo /= new_entrada.total

        new_entrada.drop(columns='total', inplace=True)

        the_df = pd.concat([the_df, new_entrada], ignore_index=True, sort=True)

    the_df = the_df.loc[(the_df.preco_venda.notnull()) &
                        (the_df.preco_custo.notnull()) &
                        (the_df.stock.notnull())]

    the_df.stock = the_df.stock.apply(lambda x: int(x))

    try:
        [os.remove(os.path.join(path, i)) for i in os.listdir(path) if '.csv' in i]
    except:
        pass

    col_gambiarra = {'codigo': u'Número ANF', 'cnp': u'Código', 'data': 'Data',
                     'stock': 'Stock', 'quantidade': 'v1', 'preco_custo': 'Preço Custo',
                     'preco_venda': 'Preço Venda'}

    the_df[u'Ano/Mês'] = [i[:6] for i in the_df.data]
    fill_df = pd.DataFrame([[0 for i in range(2, 15)]], columns=['v' + str(i) for i in range(2, 15)])
    the_df = pd.concat([the_df, fill_df], axis=1)
    the_df.rename(columns=col_gambiarra, inplace=True)
    the_df.fillna(0, inplace=True)

    the_df = the_df.assign(aux=the_df['Código'].apply(lambda x: len(str(x))))

    the_df = (the_df
              .loc[the_df.aux == 7]
              .drop(columns='aux'))

    return the_df

