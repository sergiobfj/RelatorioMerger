# merger/processor.py

import pandas as pd
import os

def mesclar_relatorios(path_pos, path_vendas):
    # Lê os arquivos
    posVenda_df = pd.read_excel(path_pos)
    vendasI_df = pd.read_excel(path_vendas)

    # Merge
    vendasI_df = vendasI_df.merge(posVenda_df, on='Título')

    # Colunas a remover e converter
    col_drop = ['Valor', 'Estágio', 'Dias no estágio', 'Endereço do Cliente']
    col_int = [
        'Código da Proposta',
        'Quantidade de inversores da Proposta',
        'Potência dos inversores - (Número) da Proposta',
        'Quantidade de Módulos da Proposta',
        'Potência dos Módulos (W) da Proposta'
    ]

    # Remove colunas terminadas em _y
    vendasI_df = vendasI_df.loc[:, ~vendasI_df.columns.str.endswith('_y')]

    # Remove sufixo _x das colunas
    vendasI_df.columns = vendasI_df.columns.str.replace('_x$', '', regex=True)

    # Remove colunas totalmente nulas
    vendasI_df = vendasI_df.dropna(axis=1)

    # Remove colunas específicas
    vendasI_df = vendasI_df.drop(columns=col_drop, errors='ignore')

    # Converte colunas para inteiro (somente se existirem)
    cols_to_convert = vendasI_df.columns.intersection(col_int)
    vendasI_df[cols_to_convert] = vendasI_df[cols_to_convert].astype(int)


    # Caminho de saída
    pasta_base = os.path.dirname(path_vendas)
    output_path = os.path.join(pasta_base, "RelatórioMesclado.xlsx")

    # Salva parcialmente (sem formatação)
    vendasI_df.to_excel(output_path, index=False)

    return output_path
