import pandas as pd
import numpy as np
from openpyxl import Workbook

def processar_relatorio_col(caminho_entrada, caminho_saida='COL_1410_Atualizado.xlsx'):
    """
    Carrega os dados gerais de coletas, recalcula a tabela dinâmica com precisão matemática,
    separa as métricas de peso médio e valor unitário, e exporta uma nova planilha formatada.
    """
    print(f"[{datetime_str()}] Carregando o arquivo: {caminho_entrada}...")
    
    try:
        xls = pd.ExcelFile(caminho_entrada)
    except Exception as e:
        raise FileNotFoundError(f"Erro ao abrir o arquivo Excel: {e}")
    
    if 'DADOS GERAIS' not in xls.sheet_names:
        raise ValueError("A aba 'DADOS GERAIS' não foi encontrada na planilha.")
    
    # 1. Leitura da aba de Dados Gerais
    df_geral = pd.read_excel(caminho_entrada, sheet_name='DADOS GERAIS')
    
    # Tratamento e limpeza de colunas numéricas críticas
    df_geral['Peso_Considerado'] = pd.to_numeric(df_geral['Peso_Considerado'], errors='coerce').fillna(0)
    df_geral['Total_Mercadoria'] = pd.to_numeric(df_geral['Total_Mercadoria'], errors='coerce').fillna(0)
    
    print("Processando agrupamento por Filial de Entrega...")
    
    # 2. Agrupamento rigoroso (Substitui a dinâmica manual imprecisa)
    df_dinamica = df_geral.groupby('Filial_Entrega').agg(
        QNTDE=('id', 'count'),
        PESO=('Peso_Considerado', 'sum'),
        VALOR=('Total_Mercadoria', 'sum')
    ).reset_index()
    
    # 3. Correção e Separação de Indicadores Unitários
    # - PESO_MEDIO_UNIT: Peso médio por volume/etiqueta (Peso Total / Quantidade)
    # - VALOR_MEDIO_UNIT: Valor financeiro médio por remessa
    df_dinamica['PESO_MEDIO_UNIT'] = np.where(
        df_dinamica['QNTDE'] > 0, 
        df_dinamica['PESO'] / df_dinamica['QNTDE'], 
        0
    )
    df_dinamica['VALOR_MEDIO_UNIT'] = np.where(
        df_dinamica['QNTDE'] > 0, 
        df_dinamica['VALOR'] / df_dinamica['QNTDE'], 
        0
    )
    
    # 4. Mapeamento da UF do Remetente correspondente
    if 'UF_Remetente' in df_geral.columns:
        uf_map = df_geral.groupby('Filial_Entrega')['UF_Remetente'].first().reset_index()
        df_dinamica = pd.merge(uf_map, df_dinamica, on='Filial_Entrega', how='right')
    else:
        df_dinamica.insert(0, 'UF_Remetente', 'SP') # Padrão caso não exista
        
    # Reorganização das colunas conforme o padrão da aba Dinâmica
    colunas_ordenadas = [
        'UF_Remetente', 
        'Filial_Entrega', 
        'QNTDE', 
        'PESO', 
        'VALOR', 
        'PESO_MEDIO_UNIT', 
        'VALOR_MEDIO_UNIT'
    ]
    
    # Garantir que apenas colunas existentes sejam selecionadas
    colunas_existentes = [col for col in colunas_ordenadas if col in df_dinamica.columns]
    df_dinamica = df_dinamica[colunas_existentes]
    
    # Renomear para exibição final idêntica à planilha original
    df_dinamica.rename(columns={'Filial_Entrega': 'DESTINO'}, inplace=True)
    
    # Arredondamento de precisão decimal para relatórios limpos
    df_dinamica['PESO'] = df_dinamica['PESO'].round(2)
    df_dinamica['VALOR'] = df_dinamica['VALOR'].round(2)
    df_dinamica['PESO_MEDIO_UNIT'] = df_dinamica['PESO_MEDIO_UNIT'].round(4)
    df_dinamica['VALOR_MEDIO_UNIT'] = df_dinamica['VALOR_MEDIO_UNIT'].round(2)
    
    print("\n--- PRÉVIA DO RESULTADO CONSOLIDADO E CORRIGIDO ---")
    print(df_dinamica.to_string(index=False))
    
    # 5. Exportação para nova planilha Excel mantendo as duas abas
    print(f"\nSalvando arquivo consolidado em: {caminho_saida}...")
    with pd.ExcelWriter(caminho_saida, engine='openpyxl') as writer:
        df_dinamica.to_excel(writer, sheet_name='DINAMICA', index=False)
        df_geral.to_excel(writer, sheet_name='DADOS GERAIS', index=False)
        
    print("Processo finalizado com sucesso!")
    return df_dinamica

def datetime_str():
    import datetime
    return datetime.datetime.now().strftime("%H:%M:%S")

if __name__ == '__main__':
    # Execução direta do script
    arquivo_entrada = 'COL - 1410.xlsx'
    processar_relatorio_col(arquivo_entrada)
