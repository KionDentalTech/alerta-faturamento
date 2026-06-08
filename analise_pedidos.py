import pandas as pd
import warnings
import os
warnings.filterwarnings('ignore')

# ============================================================
# 1. Analise do arquivo relatorio (Pedidos Janeiro ate Abril)
# ============================================================
caminho_rel = u'C:\\KionDental\\dados\\Pedidos Janeiro até Abril.xlsx'
df_rel = pd.read_excel(caminho_rel)
datas_rel = pd.to_datetime(df_rel['Data de entrada'], errors='coerce')
min_rel = datas_rel.min()
max_rel = datas_rel.max()
pedidos_rel = df_rel['Nº pedido'].nunique()

print('=== ARQUIVO RELATORIO (pasta dados/) ===')
print('Arquivo: Pedidos Janeiro ate Abril.xlsx')
print('Linhas: ' + str(len(df_rel)))
print('Pedidos unicos: ' + str(pedidos_rel))
print('Data min: ' + str(min_rel.date()))
print('Data max: ' + str(max_rel.date()))
print()

# ============================================================
# 2. Analise individual dos arquivos da pasta Pedidos/
# ============================================================
pasta = u'C:\\KionDental\\dados\\Pedidos'
arquivos = sorted(os.listdir(pasta))
arquivos_xlsx = [f for f in arquivos if f.endswith('.xlsx')]

print('=== TABELA RESUMO - TODOS OS ARQUIVOS ===')
print()
linhas_tab = []
for arq in arquivos_xlsx:
    caminho = os.path.join(pasta, arq)
    df = pd.read_excel(caminho)
    linhas = len(df)
    if 'Data de entrada' in df.columns:
        datas = pd.to_datetime(df['Data de entrada'], errors='coerce')
        dmin = str(datas.min().date()) if not pd.isnull(datas.min()) else 'N/A'
        dmax = str(datas.max().date()) if not pd.isnull(datas.max()) else 'N/A'
    else:
        dmin = 'N/A'
        dmax = 'N/A'
    if 'Nº pedido' in df.columns:
        nunicos = df['Nº pedido'].nunique()
    else:
        nunicos = 'N/A'
    linhas_tab.append({
        'Arquivo': arq,
        'Linhas': linhas,
        'Pedidos unicos': nunicos,
        'Data min': dmin,
        'Data max': dmax,
    })

df_tab = pd.DataFrame(linhas_tab)
print(df_tab.to_string(index=False))
print()

# ============================================================
# 3. Comparacao 03-26.xlsx vs 04-26.xlsx
# ============================================================
print('=== COMPARACAO 03-26.xlsx vs 04-26.xlsx ===')
df_03 = pd.read_excel(os.path.join(pasta, '03-26.xlsx'))
df_04 = pd.read_excel(os.path.join(pasta, '04-26.xlsx'))
print('03-26 shape: ' + str(df_03.shape))
print('04-26 shape: ' + str(df_04.shape))
print('Estrutura: ' + str(list(df_03.columns)))
print('Arquivos identicos? ' + str(df_03.equals(df_04)))
print('NOTA: esses dois arquivos NAO tem coluna "Nº pedido" nem "Data de entrada".')
print('      Sao relatorios de faturamento por cliente (jan-dez 2026), nao listas de pedidos.')
print('      Como o tamanho e identico e df.equals=True, sao copias exatas.')
print()

# ============================================================
# 4. Sobreposicao do relatorio com os arquivos da Pedidos/
# ============================================================
print('=== SOBREPOSICAO DO RELATORIO vs ARQUIVOS DA PASTA Pedidos/ ===')
print('Relatorio cobre: ' + str(min_rel.date()) + ' a ' + str(max_rel.date()))
print()
print('Arquivo              | Data min      | Data max      | Sobreposicao')
print('-' * 78)
for arq in arquivos_xlsx:
    caminho = os.path.join(pasta, arq)
    df = pd.read_excel(caminho)
    if 'Data de entrada' in df.columns:
        datas = pd.to_datetime(df['Data de entrada'], errors='coerce')
        mn = datas.min()
        mx = datas.max()
        overlap = (mn <= max_rel) and (mx >= min_rel)
        ov = 'SIM' if overlap else 'NAO'
        print(arq.ljust(20) + ' | ' + str(mn.date()).ljust(13) + ' | ' + str(mx.date()).ljust(13) + ' | ' + ov)
    else:
        print(arq.ljust(20) + ' | N/A           | N/A           | N/A (estrutura diferente)')
