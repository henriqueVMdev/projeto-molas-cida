# Análises complementares das molas. Ficam fora do relatório da Etapa 1 e do programa principal (molas.py):
# elasticidades da rigidez por regressão múltipla em log, modelo com variável de origem, comparação entre origens
# e resíduo do modelo por mola. Lê a mesma planilha do molas.py e grava os resultados em saidas/complementar/.
import os
import pandas as pd
import numpy as np
import seaborn as sb
import matplotlib.pyplot as plt
import statsmodels.api as sm
from scipy import stats
from adjustText import adjust_text
from matplotlib.ticker import NullLocator

NIVEL_CONFIANCA = 0.95
PASTA_SAIDAS = os.path.join("saidas", "complementar")
IDS_PROFESSOR = [107, 101, 111, 82, 94, 113, 105, 112, 122]
# Expoentes que a teoria de viga prevê para a rigidez: K proporcional a t·h³ / (n·Ltot³).
TEORIA_VIGA = {'t': 1, 'h': 3, 'n': -1, 'Ltot': -3}
LIMITE_RESIDUO = 2.0 #resíduo do modelo acima de 2 desvios-padrão = mola fora do padrão
COR_ORIGEM = {'grupo': '#2a78d6', 'professor': '#eb6834'}
DESCRICOES = {'h': 'Espessura do fio h (mm)', 't': 'Altura t (mm)', 'n': 'Espiras n', 'Ltot': 'Largura total Ltot (mm)'}
pd.set_option('display.width', 200)
pd.set_option('display.max_columns', None)
plt.rcParams.update({'axes.spines.top': False, 'axes.spines.right': False})

# Entrada de dados e rigidez de cada mola, do mesmo jeito que no molas.py (média de carga ÷ alongamento).
arquivo = "CIDA_Dados_Complementares.xlsx"
dados_df = pd.read_excel(arquivo, sheet_name="tabela1", usecols="B:I", header=1)
dados_df = dados_df.dropna(subset=['h']).astype({'ID': int, 'n': int})
ensaios_df = pd.read_excel(arquivo, sheet_name="tabela2", usecols="B:D", header=1)
ensaios_df['rigidez_ponto'] = ensaios_df['carga'] / ensaios_df['alongamento']
dados_df['rigidez'] = dados_df['ID'].map(ensaios_df.groupby('ID')['rigidez_ponto'].mean())
dados_df['dutilidade'] = dados_df['ID'].map(ensaios_df.groupby('ID')['alongamento'].last()) / dados_df['C']
dados_df['origem'] = np.where(dados_df['ID'].isin(IDS_PROFESSOR), 'professor', 'grupo')
dados_df['indice_rigidez'] = dados_df.eval('t * h**3 / (n * Ltot**3)')


# Termina todo gráfico novo do mesmo jeito: ajusta as margens (margem_inferior reserva espaço para a legenda
# embaixo), salva o PNG em saidas/ e mostra na tela.
def salvar_grafico(figura, nome_arquivo, margem_inferior=0):
    figura.tight_layout(rect=(0, margem_inferior, 1, 1))
    figura.savefig(os.path.join(PASTA_SAIDAS, nome_arquivo), dpi=150)
    plt.show()
    plt.close(figura) #libera a figura depois de salva e mostrada


# As duas funções abaixo só cuidam da aparência dos eixos em escala log das dispersões: escolhem até 4
# números legíveis (sem notação científica e com vírgula decimal). Grandeza com poucos valores, como n, mostra
# só os valores que existem.
def marcas_do_eixo(valores):
    if len(set(valores)) <= 4:
        return sorted(set(valores))
    return sorted({float(f"{valor:.2g}") for valor in np.geomspace(min(valores), max(valores), 4)})


def formatar_eixo_log(eixo_xy, marcas):
    eixo_xy.set_ticks(marcas, labels=[np.format_float_positional(marca, trim='-').replace('.', ',') for marca in marcas])
    eixo_xy.set_minor_locator(NullLocator())


# Gráfico das elasticidades: para cada parâmetro, o expoente medido (ponto escuro) com o IC 95% (traço) e o
# expoente da teoria de viga (losango vazado). Se o losango cai dentro do traço, os dados não contradizem a
# teoria. Expoente positivo = aumentar o parâmetro aumenta K; negativo = diminui.
def grafico_elasticidades(elasticidades, nome_arquivo):
    figura, eixo = plt.subplots(figsize=(8, 1.2 + 0.8 * len(elasticidades)))
    posicoes = np.arange(len(elasticidades))[::-1]
    erros = [elasticidades['expoente'] - elasticidades['IC_inf'], elasticidades['IC_sup'] - elasticidades['expoente']]
    eixo.errorbar(elasticidades['expoente'], posicoes, xerr=erros, fmt='o', color='#333333', capsize=4, label='Medido (IC 95%)')
    eixo.scatter(elasticidades['teoria_viga'], posicoes, marker='D', s=60, facecolor='white', edgecolor='#777777', linewidth=1.5, zorder=3, label='Teoria de viga')
    for posicao, expoente in zip(posicoes, elasticidades['expoente']):
        eixo.annotate(f"{expoente:+.2f}".replace('.', ','), (expoente, posicao), xytext=(0, 9), textcoords='offset points', ha='center', fontsize=9)
    eixo.axvline(0, color='#bbbbbb', linewidth=1)
    eixo.set_ylim(-0.6, len(elasticidades) - 0.2) #espaço para o valor de cima não encostar no título
    eixo.set(yticks=posicoes, yticklabels=elasticidades['descricao'], xlabel="Expoente (% de mudança em K para 1% de mudança no parâmetro)",
             title="Elasticidades da rigidez K, com os outros parâmetros fixos")
    eixo.legend(frameon=False, loc='lower right')
    salvar_grafico(figura, nome_arquivo)


# Rigidez prevista pelo modelo de elasticidades contra a medida, em escala log. Mola em cima da linha
# diagonal = o modelo acerta; longe da linha = mola fora do padrão (é a base do método 2 de outliers).
def grafico_previsto_medido(tabela, rigidez_prevista, r2, nome_arquivo):
    figura, eixo = plt.subplots(figsize=(7, 6))
    sb.scatterplot(x=rigidez_prevista, y=tabela['rigidez'], hue=tabela['origem'], palette=COR_ORIGEM, s=40, ax=eixo, zorder=3)
    limites = [min(rigidez_prevista.min(), tabela['rigidez'].min()) * 0.85, max(rigidez_prevista.max(), tabela['rigidez'].max()) * 1.15]
    eixo.plot(limites, limites, color='#777777', linewidth=1.2, label='Previsto = medido')
    eixo.set(xscale='log', yscale='log', xlim=limites, ylim=limites, xlabel="Rigidez prevista pelo modelo (kgf/mm)", ylabel="Rigidez medida K (kgf/mm)",
             title=f"Modelo de elasticidades: previsto × medido (R² = {r2:.3f})".replace('.', ','))
    formatar_eixo_log(eixo.xaxis, marcas_do_eixo(tabela['rigidez']))
    formatar_eixo_log(eixo.yaxis, marcas_do_eixo(tabela['rigidez']))
    eixo.legend(frameon=False, title=None)
    figura.tight_layout()
    rotulos = [eixo.text(x, y, str(id_mola), fontsize=7, color='#444444') for id_mola, x, y in zip(tabela['ID'], rigidez_prevista, tabela['rigidez'])]
    adjust_text(rotulos, x=rigidez_prevista.to_numpy(), y=tabela['rigidez'].to_numpy(), ax=eixo)
    salvar_grafico(figura, nome_arquivo)


# Resíduo padronizado de cada mola no modelo de rigidez: quantos desvios-padrão o K medido fica acima (barra
# para cima) ou abaixo (para baixo) do previsto pela geometria. Fora das linhas tracejadas (±2) = outlier pelo
# método 2. As cores por origem mostram se professor e grupo ficam de lados diferentes.
def grafico_residuos(tabela, residuo_padronizado, nome_arquivo):
    residuos = tabela.assign(residuo=residuo_padronizado.to_numpy(), mola=tabela['ID'].astype(str)).sort_values('residuo')
    figura, eixo = plt.subplots(figsize=(10, 4.5))
    sb.barplot(residuos, x='mola', y='residuo', hue='origem', palette=COR_ORIGEM, saturation=1, dodge=False, ax=eixo)
    for limite in (-LIMITE_RESIDUO, LIMITE_RESIDUO):
        eixo.axhline(limite, color='#999999', linestyle='--', linewidth=1)
    eixo.axhline(0, color='#555555', linewidth=1)
    eixo.set(xlabel='ID da mola', ylabel='Resíduo padronizado (desvios-padrão)', title='Resíduo do modelo de rigidez por mola (método 2 de outliers)')
    eixo.legend(title=None, frameon=False)
    salvar_grafico(figura, nome_arquivo)


# Elasticidades da rigidez: regressão em log (statsmodels OLS) ln K = a0 + a1·ln t + a2·ln h + a3·ln n +
# a4·ln Ltot. Cada expoente diz quantos % K muda para 1% de mudança no parâmetro, com os outros fixos. A tabela
# compara cada expoente com a teoria de viga (1, 3, −1, −3): o p_vs_teoria testa se a diferença para a teoria é
# significativa, e teoria_no_IC diz se o valor teórico cai dentro do IC 95%.
parametros_modelo = list(TEORIA_VIGA)
modelo_rigidez = sm.OLS(np.log(dados_df['rigidez']), sm.add_constant(np.log(dados_df[parametros_modelo]))).fit()
intervalo_modelo = modelo_rigidez.conf_int(1 - NIVEL_CONFIANCA).loc[parametros_modelo]
elasticidades_df = pd.DataFrame({
    'parametro': parametros_modelo,
    'descricao': [DESCRICOES[parametro] for parametro in parametros_modelo],
    'expoente': modelo_rigidez.params[parametros_modelo].to_numpy(),
    'erro_padrao': modelo_rigidez.bse[parametros_modelo].to_numpy(),
    'IC_inf': intervalo_modelo[0].to_numpy(),
    'IC_sup': intervalo_modelo[1].to_numpy(),
    'teoria_viga': list(TEORIA_VIGA.values()),
})
diferenca_padronizada = (elasticidades_df['expoente'] - elasticidades_df['teoria_viga']) / elasticidades_df['erro_padrao']
elasticidades_df['p_vs_teoria'] = 2 * stats.t.sf(diferenca_padronizada.abs(), modelo_rigidez.df_resid)
elasticidades_df['teoria_no_IC'] = elasticidades_df['teoria_viga'].between(elasticidades_df['IC_inf'], elasticidades_df['IC_sup'])
rigidez_prevista = np.exp(modelo_rigidez.fittedvalues)

# Modelo com uma variável de origem (1 = professor, 0 = grupo): mede quanto as molas do professor são mais
# rígidas que as do grupo com a mesma geometria medida. No molas.py essa diferença aparece como desvio do h
# medido em relação ao nominal.
log_com_origem = np.log(dados_df[parametros_modelo]).assign(professor=(dados_df['origem'] == 'professor').astype(float))
modelo_com_origem = sm.OLS(np.log(dados_df['rigidez']), sm.add_constant(log_com_origem)).fit()
intervalo_com_origem = modelo_com_origem.conf_int(1 - NIVEL_CONFIANCA)
razao_rigidez_indice = (dados_df['rigidez'] / dados_df['indice_rigidez']).groupby(dados_df['origem']).agg(['median', 'min', 'max'])

# Resíduo padronizado de cada mola no modelo de rigidez e comparação das origens em valores absolutos
# (Mann-Whitney, que não exige distribuição normal).
residuo_padronizado = modelo_rigidez.resid / modelo_rigidez.resid.std()
residuos_df = pd.DataFrame({'ID': dados_df['ID'], 'origem': dados_df['origem'], 'K_vs_previsto': np.exp(modelo_rigidez.resid) - 1,
                            'residuo_padronizado': residuo_padronizado, 'fora': residuo_padronizado.abs() > LIMITE_RESIDUO})
compara_origens = {coluna: stats.mannwhitneyu(*[dados_df.loc[dados_df['origem'] == origem, coluna] for origem in ('professor', 'grupo')]).pvalue for coluna in ('rigidez', 'dutilidade')}

print("Elasticidades da rigidez (regressão múltipla em log)")
print(elasticidades_df.to_string(index=False, float_format='{:.3f}'.format))
print(f"R² do modelo: {modelo_rigidez.rsquared:.3f}")
print(f"\nCom a variável de origem: molas do professor {np.exp(modelo_com_origem.params['professor']) - 1:.0%} mais rígidas com a mesma geometria medida.")
print("Rigidez ÷ índice de rigidez por origem:")
print(razao_rigidez_indice.round(1))
print(f"\nMann-Whitney professor × grupo: rigidez p = {compara_origens['rigidez']:.2f}, dutilidade p = {compara_origens['dutilidade']:.2f}")
print("\nResíduo do modelo por mola (fora = acima de", LIMITE_RESIDUO, "desvios-padrão)")
print(residuos_df.sort_values('residuo_padronizado').to_string(index=False, float_format='{:.2f}'.format))

os.makedirs(PASTA_SAIDAS, exist_ok=True)
grafico_elasticidades(elasticidades_df, "elasticidades_rigidez.png")
grafico_previsto_medido(dados_df, rigidez_prevista, modelo_rigidez.rsquared, "rigidez_prevista_x_medida.png")
grafico_residuos(dados_df, residuo_padronizado, "residuos_modelo_rigidez.png")

arquivo_resultados = os.path.join(PASTA_SAIDAS, "analise_complementar.xlsx")
with pd.ExcelWriter(arquivo_resultados) as planilha_resultados:
    elasticidades_df.to_excel(planilha_resultados, sheet_name='elasticidades', index=False)
    pd.concat([modelo_com_origem.params.rename('coeficiente'), intervalo_com_origem.rename(columns={0: 'IC_inf', 1: 'IC_sup'})],
              axis=1).to_excel(planilha_resultados, sheet_name='modelo_com_origem')
    razao_rigidez_indice.to_excel(planilha_resultados, sheet_name='razao_K_indice')
    residuos_df.to_excel(planilha_resultados, sheet_name='residuos', index=False)
print(f"\nPlanilha gravada em {arquivo_resultados}")
