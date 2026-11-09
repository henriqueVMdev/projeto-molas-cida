import pandas as pd # leitura de excel
import openpyxl #biblioteca auxiliar para ler excel
import seaborn as sb
# os cria a pasta saidas/ e monta o caminho dos arquivos de gráfico.
import os
import matplotlib.pyplot as plt
import numpy as np
# scipy.stats faz o intervalo de confiança (t de Student), a média geométrica, a regressão linear (R²) e a
# correlação de Spearman. statsmodels faz a regressão múltipla das elasticidades. adjust_text afasta os
# rótulos dos pontos para um não cobrir o outro. Line2D, Patch, NullLocator e LinearSegmentedColormap só
# montam legendas, eixos e a escala de cor dos gráficos.
from scipy import stats
import statsmodels.api as sm
from adjustText import adjust_text
from matplotlib.lines import Line2D
from matplotlib.patches import Patch
from matplotlib.ticker import NullLocator
from matplotlib.colors import LinearSegmentedColormap

# Parâmetros ajustáveis: dá para mudar estes valores sem mexer no resto do código. São o nível de confiança
# dos intervalos (95%), o nível de significância das correlações (5%), o mínimo de pontos de ensaio para
# calcular IC e R², a pasta onde os gráficos são salvos, quais molas são do professor (as demais são do
# grupo) e quais têm o comprimento C suspeito (marcadas, não corrigidas).
NIVEL_CONFIANCA = 0.95
NIVEL_SIGNIFICANCIA = 0.05
MINIMO_PONTOS_POR_MOLA = 3 #com menos pontos, a mola sai sem IC e sem R²
PASTA_SAIDAS = "saidas"
IDS_PROFESSOR = [107, 101, 111, 82, 94, 113, 105, 112, 122]
# Expoentes que a teoria de viga prevê para a rigidez: K proporcional a t·h³ / (n·Ltot³).
TEORIA_VIGA = {'t': 1, 'h': 3, 'n': -1, 'Ltot': -3}
# Limites dos três métodos de outliers: regra do boxplot (1,5 × intervalo entre quartis), resíduo do modelo
# de rigidez acima de 2 desvios-padrão, e linearidade do ensaio (R² abaixo de 0,98 ou rigidez ponto a ponto
# variando mais de 20% dentro da mesma mola, medida como maior ÷ menor − 1 sem o primeiro ponto, porque o
# primeiro ponto tem carga baixa e uma pequena folga nele infla a variação sem mudar o K).
FATOR_IQR = 1.5
LIMITE_RESIDUO = 2.0
LIMITE_R2 = 0.98
LIMITE_VARIACAO_PONTOS = 0.20
# Interpretação de cada mola apontada como outlier, escrita depois de olhar os dados e conferida no loop de
# conferência. Mola apontada que não estiver aqui aparece como "sem interpretação registrada".
INTERPRETACAO_OUTLIERS = {
    107: "Comportamento real: junta tudo o que aumenta K (fio grosso, mola alta, só 3 espiras, largura pequena). Fica 21% acima do previsto pelo modelo, dentro do limite.",
    56: "Comportamento real: é a mola mais esbelta (Ltot/h = 32,3), que alonga muito antes de falhar. Conferir na folha se 150 mm foi falha ou limite do ensaio.",
    71: "Provável erro de dado: C = 260,1 idêntico ao da 36 e fora do padrão C/n das molas de h parecido. Afeta só a dutilidade.",
    36: "Provável erro de dado: C = 260,1 idêntico ao da 71 e fora do padrão C/n das molas de h parecido. Afeta só a dutilidade.",
    94: "Provável erro de dado na massa: densidade aparente muito acima das demais, enquanto a rigidez segue o modelo. A variação de 22% no ensaio vem só do primeiro ponto (0,1 kgf); sem ele cai para 6% e o K muda 0,2%.",
    112: "Efeito de medida em carga baixa: a variação de 22% vem do primeiro ponto (0,1 kgf), onde uma pequena folga pesa muito; sem ele cai para 6% e o K muda 0,2%.",
    65: "Incerto, provavelmente limite do método: só 2 pontos e K 29% abaixo do previsto, mas boa parte vem da diferença professor × grupo (com origem, −1,6) e com 17 molas uma marcação assim pode sair por acaso.",
}
# Grandezas que cada erro de dado provável afeta. No cenário sem prováveis erros, a mola sai só das análises que
# usam essas grandezas: o C errado das molas 71 e 36 só pesa em C e na dutilidade (que divide por C); a massa errada da 94 só pesa em m e densidade aparente. Mola com erro sem entrada aqui sai de tudo.
GRANDEZAS_AFETADAS_POR_ERRO = {
    71: ['C', 'dutilidade'],
    36: ['C', 'dutilidade'],
    94: ['m', 'densidade_aparente'],
}
IDS_C_SUSPEITO = [71, 36] #C = 260,1 idêntico nas duas e C/n fora do padrão das molas de h parecido; o valor não foi corrigido

# Cores fixas dos gráficos: azul = molas do grupo, laranja = molas do professor; roxo = correlação
# positiva, vermelho = negativa, em tom claro quando a correlação não é significativa.
COR_ORIGEM = {'grupo': '#2a78d6', 'professor': '#eb6834'}
CORES_CORRELACAO = {
    'Positivo, significativo': '#6250d6',
    'Negativo, significativo': '#e34948',
    'Positivo, não significativo': '#cfc9f2',
    'Negativo, não significativo': '#f6c6c5',
}
# Escala de cor das matrizes de correlação: vermelho (−1), cinza (0) e roxo (+1), as mesmas cores do sinal.
ESCALA_CORRELACAO = LinearSegmentedColormap.from_list('sinal', ['#e34948', '#d9d9d9', '#6250d6'])

# Nome de cada coluna e o texto que aparece nas tabelas e gráficos. Os dois grupos ficam sempre separados:
# os parâmetros medidos nas molas (principais) e as dimensões calculadas a partir deles (derivadas).
PARAMETROS_PRINCIPAIS = {
    'h': 'Espessura do fio h (mm)',
    't': 'Altura t (mm)',
    'L': 'Passo L (mm)',
    'n': 'Espiras n',
    'Ltot': 'Largura total Ltot (mm)',
    'C': 'Comprimento C (mm)',
    'm': 'Massa m (g)',
}
PARAMETROS_DERIVADOS = {
    'indice_rigidez': 'Índice de rigidez (mm)',
    'esbeltez': 'Esbeltez do braço',
    'fio_desenvolvido': 'Fio desenvolvido (mm)',
    'momento_inercia': 'Momento de inércia (mm⁴)',
    'densidade_aparente': 'Densidade aparente (g/cm³)',
}

# Faz o pandas imprimir as tabelas inteiras no console, sem esconder colunas com '...'.
pd.set_option('display.width', 200) #evita cortar colunas ao imprimir as tabelas
pd.set_option('display.max_columns', None)
# Aparência comum dos gráficos: sem as bordas de cima e da direita.
plt.rcParams.update({'axes.spines.top': False, 'axes.spines.right': False})


#entrada de dados
# Lê a tabela1 (dimensões de cada mola) e a tabela2 (pontos carga × alongamento de cada ensaio). Abaixo da
# última mola da tabela1 há uma nota de texto: o dropna tira as linhas sem h, o que remove essa nota sem
# depender de um número fixo de molas (assim molas novas entram sozinhas).
arquivo = "CIDA_Dados_Complementares.xlsx"
dados_df = pd.read_excel(arquivo, sheet_name= "tabela1", usecols="B:K", header=1) #resultado vem em um dataframe
dados_df = dados_df.dropna(subset=['h']).astype({'ID': int, 'n': int})
ensaios_df = pd.read_excel(arquivo, sheet_name="tabela2", usecols="B:D", header=1) #pontos de ensaio: ID, carga, alongamento

# Confere se toda mola da tabela1 tem pontos de ensaio na tabela2 e vice-versa. Se faltar alguma, avisa no
# console em vez de calcular errado em silêncio.
ids_sem_par = set(dados_df['ID']) ^ set(ensaios_df['ID'])
if ids_sem_par:
    print("AVISO: IDs presentes em só uma das abas:", sorted(ids_sem_par))

# Cálculos de cada ensaio, agrupando os pontos por mola (groupby). Para cada mola:
# - rigidez K: inclinação da reta carga × alongamento que passa pela origem, ajustada por mínimos quadrados
#   com todos os pontos (np.linalg.lstsq; o mesmo que K = Σ(alongamento·carga) / Σ(alongamento²));
# - rigidez ponto a ponto (carga ÷ alongamento em cada ponto): quantidade de pontos, média, desvio e IC 95%
#   pelo t de Student. Mostra se a mola se manteve linear; esse IC descreve a variação entre os pontos, não a
#   incerteza do K;
# - R² de uma reta livre (com intercepto), pelo linregress;
# - alongamento no último ponto, que é o ponto de falha.
# Molas com menos de 3 pontos ficam sem IC e sem R², porque com 2 pontos a reta passa exatamente por eles.
ensaios_df['rigidez_ponto'] = ensaios_df['carga'] / ensaios_df['alongamento']
pontos_por_mola = ensaios_df.groupby('ID')
ensaio_df = pontos_por_mola.agg(
    pontos=('carga', 'size'),
    media_pontos=('rigidez_ponto', 'mean'),
    desvio_pontos=('rigidez_ponto', 'std'),
    erro_pontos=('rigidez_ponto', 'sem'),
    alongamento_falha=('alongamento', 'last'),
    variacao_pontos=('rigidez_ponto', lambda pontos: pontos.max() / pontos.min() - 1), #quanto a rigidez ponto a ponto varia dentro da mola
    variacao_sem_primeiro=('rigidez_ponto', lambda pontos: pontos.iloc[1:].max() / pontos.iloc[1:].min() - 1 if len(pontos) > 2 else np.nan),
)
ensaio_df['rigidez'] = pontos_por_mola.apply(lambda pontos: np.linalg.lstsq(pontos[['alongamento']], pontos['carga'], rcond=None)[0][0])
tem_pontos_suficientes = ensaio_df['pontos'] >= MINIMO_PONTOS_POR_MOLA
ensaio_df['R2'] = pontos_por_mola.apply(lambda pontos: stats.linregress(pontos['alongamento'], pontos['carga']).rvalue**2).where(tem_pontos_suficientes)
limite_inferior, limite_superior = stats.t.interval(NIVEL_CONFIANCA, ensaio_df['pontos'] - 1, loc=ensaio_df['media_pontos'], scale=ensaio_df['erro_pontos'])
ensaio_df['IC_pontos_inf'] = pd.Series(limite_inferior, index=ensaio_df.index).where(tem_pontos_suficientes)
ensaio_df['IC_pontos_sup'] = pd.Series(limite_superior, index=ensaio_df.index).where(tem_pontos_suficientes)

# Leva rigidez e dutilidade para a tabela das molas. Dutilidade = alongamento no ponto de falha ÷ comprimento
# C. Os valores calculados substituem as colunas rigidez e dutilidade que já vinham na planilha, porque todo
# cálculo deve ser feito no Python. Também marca cada mola como 'professor' ou 'grupo'.
dados_df['rigidez'] = dados_df['ID'].map(ensaio_df['rigidez'])
dados_df['dutilidade'] = dados_df['ID'].map(ensaio_df['alongamento_falha']) / dados_df['C']
dados_df['origem'] = np.where(dados_df['ID'].isin(IDS_PROFESSOR), 'professor', 'grupo')
por_mola_df = dados_df[['ID', 'origem', 'rigidez', 'dutilidade']].merge(
    ensaio_df[['pontos', 'R2', 'media_pontos', 'desvio_pontos', 'IC_pontos_inf', 'IC_pontos_sup', 'variacao_pontos', 'variacao_sem_primeiro']], left_on='ID', right_index=True)

# Calcula as 5 dimensões derivadas a partir das medidas da tabela1, com as fórmulas escritas como na tabela
# do CLAUDE.md (pandas eval). Elas ficam numa tabela separada (derivadas_df, junto com origem, rigidez e
# dutilidade de cada mola) para não se misturarem com os
# parâmetros principais: índice de rigidez (rigidez esperada pela teoria de viga), esbeltez do braço (o índice
# de dutilidade saiu por ser praticamente a mesma medida que a esbeltez, ρ 0,99), fio desenvolvido (comprimento aproximado de fio), momento de inércia da seção e
# densidade aparente (massa ÷ volume aproximado; serve para checar se massa e geometria combinam).
derivadas_df = dados_df[['ID', 'origem', 'rigidez', 'dutilidade']].assign(
    indice_rigidez=dados_df.eval('t * h**3 / (n * Ltot**3)'),
    esbeltez=dados_df.eval('Ltot / h'),
    fio_desenvolvido=dados_df.eval('2 * n * Ltot'),
    momento_inercia=dados_df.eval('t * h**3 / 12'),
    densidade_aparente=dados_df.eval('m / (t * h * 2 * n * Ltot) * 1000'),
)


# Termina todo gráfico novo do mesmo jeito: ajusta as margens (margem_inferior reserva espaço para a legenda
# embaixo), salva o PNG em saidas/ e mostra na tela.
def salvar_grafico(figura, nome_arquivo, margem_inferior=0):
    figura.tight_layout(rect=(0, margem_inferior, 1, 1))
    figura.savefig(os.path.join(PASTA_SAIDAS, nome_arquivo), dpi=150)
    plt.show()


# Legenda de cores por origem (azul = grupo, laranja = professor), usada em vários gráficos.
LEGENDA_ORIGEM = [Line2D([0], [0], marker='o', linestyle='', color=cor, label=f"Molas: {origem}") for origem, cor in COR_ORIGEM.items()]


# Monta a tabela de estatísticas de cada grandeza em três conjuntos (todas as molas, só as do grupo, só as do
# professor). O agg calcula quantidade, média, desvio, mínimo, máximo, mediana e erro padrão de uma vez. Em
# seguida vêm o IC 95% pelo t de Student e o mesmo IC feito na escala log, que dá a média geométrica e um
# intervalo que nunca fica negativo: serve para grandezas positivas e assimétricas, como rigidez e
# dutilidade, em que poucas molas muito altas puxam a média comum.
def estatisticas_por_conjunto(tabela, grandezas):
    conjuntos = {'todas': tabela, **dict(list(tabela.groupby('origem')))}
    partes = []
    for nome_conjunto, conjunto in conjuntos.items():
        valores = conjunto[list(grandezas)]
        resumo = valores.agg(['count', 'mean', 'std', 'min', 'max', 'median', 'sem']).T
        resumo.columns = ['molas', 'media', 'desvio', 'minimo', 'maximo', 'mediana', 'erro']
        logaritmos = np.log(valores)
        resumo['IC_inf'], resumo['IC_sup'] = stats.t.interval(NIVEL_CONFIANCA, resumo['molas'] - 1, loc=resumo['media'], scale=resumo['erro'])
        resumo['amplitude'] = resumo['maximo'] - resumo['minimo']
        resumo['media_geometrica'] = stats.gmean(valores)
        resumo['IC_log_inf'], resumo['IC_log_sup'] = np.exp(stats.t.interval(NIVEL_CONFIANCA, resumo['molas'] - 1, loc=logaritmos.mean(), scale=logaritmos.sem()))
        resumo.insert(0, 'grandeza', [grandezas[coluna] for coluna in resumo.index])
        resumo.insert(0, 'conjunto', nome_conjunto)
        partes.append(resumo)
    colunas = ['conjunto', 'grandeza', 'molas', 'media', 'desvio', 'IC_inf', 'IC_sup', 'minimo', 'maximo', 'mediana',
               'amplitude', 'media_geometrica', 'IC_log_inf', 'IC_log_sup']
    return pd.concat(partes)[colunas].astype({'molas': int}).reset_index(drop=True)


# Matriz de Spearman entre todas as colunas pedidas, numa única chamada do spearmanr: devolve uma tabela com
# o ρ de cada par e outra com o p-valor. Serve para a matriz parâmetro × parâmetro e para as correlações com
# rigidez e dutilidade.
def matriz_spearman(tabela, colunas):
    teste = stats.spearmanr(tabela[colunas])
    return (pd.DataFrame(teste.statistic, index=colunas, columns=colunas),
            pd.DataFrame(teste.pvalue, index=colunas, columns=colunas))


# Correlação de Spearman de cada parâmetro com a rigidez e com a dutilidade, aproveitando a matriz
# de Spearman da tabela inteira. Spearman compara a ordem dos valores, então mede se um cresce quando o
# outro cresce mesmo que a relação não seja uma reta. Para cada par guarda o ρ (de −1 a +1), o p-valor, o
# sentido (positivo, negativo ou nulo), se é significativo (p < 0,05) e o IC 95% do ρ.
def correlacoes_spearman(tabela, parametros):
    colunas = [*parametros, 'rigidez', 'dutilidade']
    rho, p_valor = matriz_spearman(tabela, colunas)
    partes = []
    for resultado, coluna_resultado in (('Rigidez K', 'rigidez'), ('Dutilidade', 'dutilidade')):
        partes.append(pd.DataFrame({
            'resultado': resultado,
            'parametro': list(parametros),
            'descricao': list(parametros.values()),
            'rho': rho.loc[list(parametros), coluna_resultado].to_numpy(),
            'p_valor': p_valor.loc[list(parametros), coluna_resultado].to_numpy(),
        }))
    correlacoes = pd.concat(partes, ignore_index=True)
    correlacoes['rho'] = correlacoes['rho'].where(correlacoes['rho'].abs() > 1e-12, 0.0) #resto de arredondamento vira zero
    correlacoes['efeito'] = np.sign(correlacoes['rho']).map({1.0: 'positivo', -1.0: 'negativo', 0.0: 'nulo'})
    correlacoes['significativo'] = correlacoes['p_valor'] < NIVEL_SIGNIFICANCIA
    # IC 95% do ρ pela transformação z de Fisher, com o erro padrão de Fieller para Spearman (1,06 / raiz(n − 3)).
    # Mostra a firmeza de cada correlação: IC que encosta no zero = conclusão fraca, mesmo se p < 0,05.
    margem_z = stats.norm.ppf((1 + NIVEL_CONFIANCA) / 2) * 1.06 / np.sqrt(len(tabela) - 3)
    correlacoes['IC_rho_inf'] = np.tanh(np.arctanh(correlacoes['rho']) - margem_z)
    correlacoes['IC_rho_sup'] = np.tanh(np.arctanh(correlacoes['rho']) + margem_z)
    return correlacoes


# Calcula o menor |ρ| que já é significativo com a quantidade de molas atual (com 17 molas, cerca de 0,48).
# Vira a linha tracejada dos gráficos de barras e se ajusta sozinho se entrarem molas novas. Nenhuma
# biblioteca tem essa conta pronta: ela sai do t crítico de Student.
def rho_minimo_significativo(quantidade_molas):
    t_critico = stats.t.ppf(1 - NIVEL_SIGNIFICANCIA / 2, quantidade_molas - 2)
    return t_critico / np.sqrt(quantidade_molas - 2 + t_critico**2)


# Gráfico de barras da influência (seaborn barplot): uma barra por parâmetro, com comprimento igual ao ρ.
# Roxo = positivo, vermelho = negativo, tom claro = não significativo. Um painel para a rigidez e outro para a
# dutilidade; as linhas tracejadas marcam o limite de significância. Salva o PNG em saidas/.
def grafico_barras_correlacao(correlacoes, rho_minimo, titulo, nome_arquivo):
    sentido = np.where(correlacoes['rho'] >= 0, 'Positivo', 'Negativo')
    significancia = np.where(correlacoes['significativo'], ', significativo', ', não significativo')
    correlacoes = correlacoes.assign(categoria=np.char.add(sentido, significancia))
    figura, eixos = plt.subplots(1, 2, figsize=(11, 1.9 + 0.5 * correlacoes['parametro'].nunique()), sharey=True)
    for eixo, resultado in zip(eixos, ['Rigidez K', 'Dutilidade']):
        sb.barplot(correlacoes[correlacoes['resultado'] == resultado], x='rho', y='descricao', hue='categoria',
                   palette=CORES_CORRELACAO, saturation=1, dodge=False, legend=False, width=0.6, ax=eixo)
        for barras in eixo.containers:
            for rotulo in eixo.bar_label(barras, fmt=lambda rho: f"{rho:+.2f}".replace('.', ','), padding=4, fontsize=10):
                rotulo.set_bbox({'facecolor': 'white', 'edgecolor': 'none', 'pad': 1})
        eixo.axvline(0, color='#555555', linewidth=1)
        for limite in (-rho_minimo, rho_minimo):
            eixo.axvline(limite, color='#999999', linewidth=1, linestyle='--')
        eixo.set(xlim=(-1.2, 1.2), xlabel="ρ de Spearman", ylabel="", title=resultado)
        eixo.spines['left'].set_visible(False)
        eixo.tick_params(axis='y', length=0)
    legenda = [Patch(color=cor, label=categoria) for categoria, cor in CORES_CORRELACAO.items()] + [
        Line2D([0], [0], color='#999999', linestyle='--', label=f"Limite de significância (|ρ| = {rho_minimo:.2f})".replace('.', ','))]
    figura.legend(handles=legenda, loc='lower center', ncol=3, frameon=False, fontsize=9)
    figura.suptitle(titulo)
    salvar_grafico(figura, nome_arquivo, 0.1)


# Matriz de correlação em mapa de calor: cada célula traz o ρ entre dois parâmetros, com * quando é
# significativo. Mostra só a metade de baixo, porque a matriz é simétrica. Roxo = os dois crescem juntos,
# vermelho = um cresce quando o outro diminui, cinza = sem relação. Parâmetros que andam juntos no conjunto de
# molas não têm efeitos separáveis por correlação simples.
def grafico_matriz(rho, p_valor, rotulos, titulo, nome_arquivo):
    textos = rho.map(lambda valor: f"{valor:+.2f}".replace('.', ',')) + p_valor.map(lambda p: '*' if p < NIVEL_SIGNIFICANCIA else '')
    rho, textos = rho.iloc[1:, :-1], textos.iloc[1:, :-1] #tira a 1ª linha e a última coluna, que só teriam a diagonal
    acima_da_diagonal = np.triu(np.ones(rho.shape, dtype=bool), k=1)
    figura, eixo = plt.subplots(figsize=(1.2 * len(rotulos) + 3, 0.9 * len(rotulos) + 2))
    sb.heatmap(rho, mask=acima_da_diagonal, annot=textos, fmt='', cmap=ESCALA_CORRELACAO, vmin=-1, vmax=1, center=0,
               linewidths=2, linecolor='white', cbar_kws={'label': 'ρ de Spearman'}, ax=eixo,
               xticklabels=rotulos[:-1], yticklabels=rotulos[1:])
    eixo.tick_params(axis='x', rotation=35)
    plt.setp(eixo.get_xticklabels(), ha='right')
    eixo.set_title(f"{titulo}\n(* = significativo, p < 0,05)")
    salvar_grafico(figura, nome_arquivo)


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


# Método 1 de outliers, regra do boxplot: para cada grandeza, valor abaixo de Q1 − 1,5·IQR ou acima de
# Q3 + 1,5·IQR (Q1 e Q3 = quartis, IQR = Q3 − Q1). Devolve uma linha por valor fora, com o motivo.
def outliers_boxplot(tabela, rotulos, metodo):
    colunas = list(rotulos)
    quartil_1, quartil_3 = tabela[colunas].quantile(0.25), tabela[colunas].quantile(0.75)
    limite_inferior = quartil_1 - FATOR_IQR * (quartil_3 - quartil_1)
    limite_superior = quartil_3 + FATOR_IQR * (quartil_3 - quartil_1)
    valores = tabela.melt(id_vars='ID', value_vars=colunas, var_name='grandeza', value_name='valor')
    acima = valores['valor'] > valores['grandeza'].map(limite_superior)
    abaixo = valores['valor'] < valores['grandeza'].map(limite_inferior)
    fora = valores[acima | abaixo]
    motivos = [f"{rotulos[grandeza]} = {valor:.4g}, " + (f"acima do limite {limite_superior[grandeza]:.4g}" if valor > limite_superior[grandeza] else f"abaixo do limite {limite_inferior[grandeza]:.4g}")
               for grandeza, valor in zip(fora['grandeza'], fora['valor'])]
    return pd.DataFrame({'ID': fora['ID'], 'metodo': metodo, 'grandeza': fora['grandeza'], 'motivo': motivos})


# Boxplots com os pontos de cada mola por cima (azul = grupo, laranja = professor). A caixa vai do 1º ao 3º
# quartil, a linha do meio é a mediana e os bigodes param no limite da regra do boxplot; as molas fora desse
# limite levam o ID escrito ao lado.
def grafico_boxplots(tabela, rotulos, outliers, titulo, nome_arquivo):
    colunas_figura = 3
    linhas_figura = int(np.ceil(len(rotulos) / colunas_figura))
    figura, eixos = plt.subplots(linhas_figura, colunas_figura, figsize=(12, 3.3 * linhas_figura + 1))
    eixos = eixos.flatten()
    for eixo, (coluna, rotulo) in zip(eixos, rotulos.items()):
        sb.boxplot(y=tabela[coluna], color='#eeeeee', width=0.35, whis=FATOR_IQR, showfliers=False, ax=eixo)
        sb.stripplot(data=tabela, y=coluna, hue='origem', palette=COR_ORIGEM, jitter=0.06, size=5, legend=False, ax=eixo)
        molas_fora = tabela[tabela['ID'].isin(outliers.loc[outliers['grandeza'] == coluna, 'ID'])].groupby(coluna)['ID'].apply(list)
        for valor, ids_molas in molas_fora.items(): #molas com o mesmo valor ganham um rótulo só
            texto = f"mola {ids_molas[0]}" if len(ids_molas) == 1 else "molas " + " e ".join(map(str, ids_molas))
            eixo.annotate(texto, (0.08, valor), xytext=(0.2, valor), fontsize=8, va='center',
                          arrowprops={'arrowstyle': '-', 'color': '#888888'})
        eixo.set(ylabel=rotulo, xlabel='', xlim=(-0.5, 0.75))
    for eixo in eixos[len(rotulos):]:
        eixo.set_visible(False)
    figura.legend(handles=LEGENDA_ORIGEM, loc='lower center', ncol=2, frameon=False)
    figura.suptitle(f"{titulo}\n(bigodes no limite da regra do boxplot; IDs = molas fora do limite)")
    salvar_grafico(figura, nome_arquivo, 0.04)


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


# Curva carga × alongamento de cada mola: os pontos do ensaio e a reta da rigidez K pela origem. O título traz
# o R² e a variação da rigidez ponto a ponto (com e sem o 1º ponto); as molas fora do padrão pelo método 3 têm o título marcado.
def grafico_curvas_ensaio(ensaios, por_mola, nome_arquivo):
    colunas_figura = 5
    linhas_figura = int(np.ceil(len(por_mola) / colunas_figura))
    figura, eixos = plt.subplots(linhas_figura, colunas_figura, figsize=(16, 3.2 * linhas_figura + 0.8))
    eixos = eixos.flatten()
    for eixo, (_, mola) in zip(eixos, por_mola.iterrows()):
        pontos = ensaios[ensaios['ID'] == mola['ID']]
        eixo.scatter(pontos['alongamento'], pontos['carga'], color=COR_ORIGEM[mola['origem']], s=22, zorder=3)
        alongamento_maximo = pontos['alongamento'].max()
        eixo.plot([0, alongamento_maximo], [0, mola['rigidez'] * alongamento_maximo], color='#777777', linewidth=1.2)
        fora_do_padrao = mola['R2'] < LIMITE_R2 or mola['variacao_sem_primeiro'] > LIMITE_VARIACAO_PONTOS
        texto_r2 = 'sem R²' if pd.isna(mola['R2']) else f"R² {mola['R2']:.4f}"
        titulo = f"Mola {mola['ID']}" + (' (fora do padrão)' if fora_do_padrao else '') + f"\n{texto_r2}, variação {mola['variacao_pontos']:.0%} ({'sem 1º ponto: ' + format(mola['variacao_sem_primeiro'], '.0%') if pd.notna(mola['variacao_sem_primeiro']) else 'só 2 pontos'})"
        eixo.set_title(titulo.replace('.', ','), fontsize=9, fontweight='bold' if fora_do_padrao else 'normal')
        eixo.tick_params(labelsize=8)
    for eixo in eixos[len(por_mola):]:
        eixo.set_visible(False)
    figura.supxlabel('Alongamento (mm)')
    figura.supylabel('Carga (kgf)')
    figura.suptitle('Curvas carga × alongamento de cada mola, com a reta da rigidez K (método 3 de outliers)')
    salvar_grafico(figura, nome_arquivo)


# Gráfico da comparação com e sem outliers: para cada parâmetro, o ρ com todas as molas e nos cenários sem
# as prováveis de erro de dado e no teste de estresse. Marcadores próximos = conclusão estável; marcador que
# cruza a faixa cinza (não significativo) = conclusão que depende das molas marcadas.
def grafico_com_sem_outliers(correlacoes_cenarios, parametros, rho_minimo, titulo, nome_arquivo):
    cenarios_mostrados = {'todas as molas': ('o', '#222222'), 'sem prováveis erros de dado': ('s', '#777777'), 'teste de estresse: sem todas as marcadas': ('^', '#aaaaaa')}
    figura, eixos = plt.subplots(1, 2, figsize=(11, 1.9 + 0.5 * len(parametros)), sharey=True)
    posicoes = np.arange(len(parametros))[::-1]
    for eixo, resultado in zip(eixos, ['Rigidez K', 'Dutilidade']):
        eixo.axvspan(-rho_minimo, rho_minimo, color='#eeeeee', zorder=0)
        for deslocamento, (cenario, (marcador, cor)) in zip((0.2, 0, -0.2), cenarios_mostrados.items()):
            linhas = correlacoes_cenarios[(correlacoes_cenarios['cenario'] == cenario) & (correlacoes_cenarios['resultado'] == resultado)].set_index('parametro').loc[list(parametros)]
            eixo.scatter(linhas['rho'], posicoes + deslocamento, marker=marcador, color=cor, s=40, label=cenario, zorder=3)
        eixo.axvline(0, color='#555555', linewidth=1)
        eixo.set(xlim=(-1.1, 1.1), yticks=posicoes, yticklabels=list(parametros.values()), xlabel='ρ de Spearman', title=resultado)
    handles, rotulos = eixos[0].get_legend_handles_labels()
    figura.legend(handles + [Patch(color='#eeeeee', label=f"faixa não significativa (|ρ| < {rho_minimo:.2f} com todas as molas)".replace('.', ','))],
                  rotulos + [f"faixa não significativa (|ρ| < {rho_minimo:.2f} com todas as molas)".replace('.', ',')], loc='lower center', ncol=2, frameon=False, fontsize=9)
    figura.suptitle(titulo)
    salvar_grafico(figura, nome_arquivo, 0.12)


# Rigidez e dutilidade por número de espiras (n tem só 3 valores, então um boxplot por grupo mostra melhor que
# um ρ). Rigidez em escala log, porque varia 28 vezes. As outras dimensões também mudam entre os grupos, então
# a diferença entre as caixas não é efeito isolado de n (o efeito isolado é o expoente do modelo).
def grafico_por_espiras(tabela, nome_arquivo):
    figura, eixos = plt.subplots(1, 2, figsize=(11, 4.8))
    for eixo, (coluna, rotulo) in zip(eixos, (('rigidez', 'Rigidez K (kgf/mm)'), ('dutilidade', 'Dutilidade (fração de C)'))):
        sb.boxplot(tabela, x='n', y=coluna, color='#eeeeee', width=0.5, showfliers=False, ax=eixo)
        sb.stripplot(tabela, x='n', y=coluna, hue='origem', palette=COR_ORIGEM, jitter=0.12, size=6, legend=False, ax=eixo)
        eixo.set(xlabel='Número de espiras n', ylabel=rotulo)
    eixos[0].set_yscale('log')
    formatar_eixo_log(eixos[0].yaxis, marcas_do_eixo(tabela['rigidez']))
    figura.legend(handles=LEGENDA_ORIGEM, loc='lower center', ncol=2, frameon=False)
    figura.suptitle("Rigidez e dutilidade por número de espiras\n(n muda junto com outras dimensões: não é efeito isolado)")
    salvar_grafico(figura, nome_arquivo, 0.06)


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


# Uma figura com um painel por parâmetro, mostrando o resultado (rigidez ou dutilidade) contra esse
# parâmetro (seaborn scatterplot). Pontos azuis = grupo, laranja = professor, número = ID da mola, afastado
# dos vizinhos pelo adjust_text. A linha cinza é a tendência, ajustada com os dois eixos em escala log, porque
# as relações da mola são do tipo potência (por exemplo, K cresce com h³). O título de cada painel traz o ρ e
# se ele é significativo (as correlações chegam já filtradas para o resultado do gráfico).
def grafico_dispersoes(tabela, parametros, coluna_resultado, rotulo_resultado, correlacoes, nome_arquivo):
    colunas_figura = min(4, len(parametros))
    linhas_figura = int(np.ceil(len(parametros) / colunas_figura))
    figura, eixos = plt.subplots(linhas_figura, colunas_figura, figsize=(max(4 * colunas_figura, 7), 3.8 * linhas_figura + 1.3), sharey=True, squeeze=False)
    eixos = eixos.flatten()
    for eixo, (coluna, descricao) in zip(eixos, parametros.items()):
        sb.scatterplot(tabela, x=coluna, y=coluna_resultado, hue='origem', palette=COR_ORIGEM, s=36, legend=False, zorder=3, ax=eixo)
        inclinacao, intercepto = np.polyfit(np.log(tabela[coluna]), np.log(tabela[coluna_resultado]), 1)
        valores_x = np.array([tabela[coluna].min(), tabela[coluna].max()])
        eixo.plot(valores_x, np.exp(intercepto) * valores_x**inclinacao, color='#777777', linewidth=1.5, zorder=2)
        correlacao = correlacoes.set_index('parametro').loc[coluna]
        marca = 'significativo' if correlacao['significativo'] else 'não significativo'
        eixo.set(xscale='log', yscale='log', xlabel=descricao, ylabel=rotulo_resultado)
        eixo.set_title(f"ρ = {correlacao['rho']:+.2f} ({marca})".replace('.', ','), fontsize=10)
        formatar_eixo_log(eixo.xaxis, marcas_do_eixo(tabela[coluna]))
        eixo.grid(color='#e6e6e6', linewidth=0.7)
    formatar_eixo_log(eixos[0].yaxis, marcas_do_eixo(tabela[coluna_resultado]))
    for eixo in eixos[len(parametros):]:
        eixo.set_visible(False)
    figura.legend(handles=LEGENDA_ORIGEM + [Line2D([0], [0], color='#777777', label='Tendência')], loc='lower center', ncol=3, frameon=False)
    nome_x = 'cada parâmetro' if len(parametros) > 1 else next(iter(parametros.values()))
    figura.suptitle(f"{rotulo_resultado} em função de {nome_x}\n(eixos em escala log)")
    figura.tight_layout(rect=(0, 0.06, 1, 1)) #antes do adjust_text, que precisa das posições finais
    for eixo, coluna in zip(eixos, parametros):
        rotulos = [eixo.text(x, y, str(id_mola), fontsize=7, color='#444444') for id_mola, x, y in zip(tabela['ID'], tabela[coluna], tabela[coluna_resultado])]
        adjust_text(rotulos, x=tabela[coluna].to_numpy(), y=tabela[coluna_resultado].to_numpy(), ax=eixo)  # desvia também dos pontos
    salvar_grafico(figura, nome_arquivo, 0.06)


# Estatísticas entre molas: uma tabela para rigidez e dutilidade (em %) e outra para as dimensões derivadas.
estatisticas_df = estatisticas_por_conjunto(
    dados_df.assign(dutilidade_pct=dados_df['dutilidade'] * 100),
    {'rigidez': 'Rigidez K (kgf/mm)', 'dutilidade_pct': 'Dutilidade (%)'},
)
estatisticas_derivadas_df = estatisticas_por_conjunto(derivadas_df, PARAMETROS_DERIVADOS)

# Aplica Spearman aos dois grupos de parâmetros, cada um na sua tabela, e calcula o |ρ| mínimo significativo
# para o número atual de molas.
correlacoes_principais_df = correlacoes_spearman(dados_df, PARAMETROS_PRINCIPAIS)
correlacoes_derivadas_df = correlacoes_spearman(derivadas_df, PARAMETROS_DERIVADOS)
rho_minimo = rho_minimo_significativo(len(dados_df))

# Parâmetro × parâmetro: matriz de Spearman dentro de cada grupo (principais e derivadas, sem misturar) e a
# lista dos pares que andam juntos de forma significativa. Também calcula a correlação entre rigidez e
# dutilidade, para ver se o que aumenta uma tende a reduzir a outra.
matriz_principais_rho, matriz_principais_p = matriz_spearman(dados_df, list(PARAMETROS_PRINCIPAIS))
matriz_derivadas_rho, matriz_derivadas_p = matriz_spearman(derivadas_df, list(PARAMETROS_DERIVADOS))


def pares_significativos(rho, p_valor):
    abaixo_da_diagonal = np.tril(np.ones(rho.shape, dtype=bool), k=-1)
    pares = pd.DataFrame({'rho': rho.where(abaixo_da_diagonal).stack(), 'p_valor': p_valor.where(abaixo_da_diagonal).stack()})
    return pares[pares['p_valor'] < NIVEL_SIGNIFICANCIA].sort_values('rho', key=abs, ascending=False)


# Elasticidades da rigidez: regressão em log (statsmodels OLS) ln K = a0 + a1·ln t + a2·ln h + a3·ln n +
# a4·ln Ltot. Cada expoente diz quantos % K muda para 1% de mudança no parâmetro, com os outros fixos; por isso
# separa efeitos que a correlação simples mistura (ex.: t e n parecem fracos no ρ). A tabela compara cada
# expoente com a teoria de viga (1, 3, −1, −3): o p_vs_teoria testa se a diferença para a teoria é
# significativa, e teoria_no_IC diz se o valor teórico cai dentro do IC 95%.
parametros_modelo = list(TEORIA_VIGA)
modelo_rigidez = sm.OLS(np.log(dados_df['rigidez']), sm.add_constant(np.log(dados_df[parametros_modelo]))).fit()
intervalo_modelo = modelo_rigidez.conf_int(1 - NIVEL_CONFIANCA).loc[parametros_modelo]
rho_com_rigidez = correlacoes_principais_df.query("resultado == 'Rigidez K'").set_index('parametro')['rho']
elasticidades_df = pd.DataFrame({
    'parametro': parametros_modelo,
    'descricao': [PARAMETROS_PRINCIPAIS[parametro] for parametro in parametros_modelo],
    'expoente': modelo_rigidez.params[parametros_modelo].to_numpy(),
    'erro_padrao': modelo_rigidez.bse[parametros_modelo].to_numpy(),
    'IC_inf': intervalo_modelo[0].to_numpy(),
    'IC_sup': intervalo_modelo[1].to_numpy(),
    'teoria_viga': list(TEORIA_VIGA.values()),
    'rho_simples': rho_com_rigidez[parametros_modelo].to_numpy(),
})
diferenca_padronizada = (elasticidades_df['expoente'] - elasticidades_df['teoria_viga']) / elasticidades_df['erro_padrao']
elasticidades_df['p_vs_teoria'] = 2 * stats.t.sf(diferenca_padronizada.abs(), modelo_rigidez.df_resid)
elasticidades_df['teoria_no_IC'] = elasticidades_df['teoria_viga'].between(elasticidades_df['IC_inf'], elasticidades_df['IC_sup'])
rigidez_prevista = np.exp(modelo_rigidez.fittedvalues)

# Versão das derivadas: o índice de rigidez já junta os quatro parâmetros na fórmula da teoria de viga, então
# basta uma regressão ln K = b0 + b1·ln(índice). A teoria prevê b1 = 1 (K proporcional ao índice). A razão
# K ÷ índice faz o papel do módulo do material; comparar essa razão entre professor e grupo mostra se as duas
# origens seguem a mesma constante.
modelo_indice = sm.OLS(np.log(derivadas_df['rigidez']), sm.add_constant(np.log(derivadas_df['indice_rigidez']))).fit()
intervalo_indice = modelo_indice.conf_int(1 - NIVEL_CONFIANCA).loc['indice_rigidez']
# Versão complementar do modelo com uma variável de origem (1 = professor, 0 = grupo). Ela mede quanto as molas
# do professor são mais rígidas que as do grupo com a mesma geometria. Se for grande, há diferença sistemática
# entre as origens (na medida de h do grupo ou na impressão), e os expoentes desta versão ficam mais precisos.
log_com_origem = np.log(dados_df[parametros_modelo]).assign(professor=(dados_df['origem'] == 'professor').astype(float))
modelo_com_origem = sm.OLS(np.log(dados_df['rigidez']), sm.add_constant(log_com_origem)).fit()
intervalo_com_origem = modelo_com_origem.conf_int(1 - NIVEL_CONFIANCA)
razao_rigidez_indice = (derivadas_df['rigidez'] / derivadas_df['indice_rigidez']).groupby(derivadas_df['origem']).agg(['median', 'min', 'max'])

# Outliers pelos três métodos. Método 1: regra do boxplot em rigidez, dutilidade e cada parâmetro (e, em
# separado, nas derivadas). Método 2: mola cujo resíduo no modelo de rigidez passa de 2 desvios-padrão, ou
# seja, K muito diferente do previsto pela geometria, mesmo que o K em si pareça normal. Método 3: ensaio pouco
# linear (R² baixo ou rigidez ponto a ponto variando mais de 20%). Cada mola apontada recebe o motivo e a
# interpretação (erro de dado ou comportamento real). Nada é descartado: as molas só são marcadas.
rotulos_boxplot = {'rigidez': 'Rigidez K (kgf/mm)', 'dutilidade': 'Dutilidade (fração de C)', **PARAMETROS_PRINCIPAIS}
residuo_padronizado = modelo_rigidez.resid / modelo_rigidez.resid.std()
fora_residuo = dados_df[residuo_padronizado.abs() > LIMITE_RESIDUO]
fora_linearidade = por_mola_df[(por_mola_df['R2'] < LIMITE_R2) | (por_mola_df['variacao_sem_primeiro'] > LIMITE_VARIACAO_PONTOS)]
outliers_df = pd.concat([
    outliers_boxplot(dados_df, rotulos_boxplot, '1. boxplot'),
    outliers_boxplot(derivadas_df, PARAMETROS_DERIVADOS, '1. boxplot (derivadas)'),
    pd.DataFrame({'ID': fora_residuo['ID'], 'metodo': '2. resíduo do modelo', 'grandeza': 'rigidez',
                  'motivo': [f"K {(np.exp(residuo) - 1):+.0%} em relação ao previsto ({padronizado:+.2f} desvios-padrão)"
                             for residuo, padronizado in zip(modelo_rigidez.resid[fora_residuo.index], residuo_padronizado[fora_residuo.index])]}),
    pd.DataFrame({'ID': fora_linearidade['ID'], 'metodo': '3. linearidade do ensaio', 'grandeza': 'ensaio',
                  'motivo': [f"sem o 1º ponto, a rigidez ponto a ponto varia {variacao:.0%}" + ('' if pd.isna(r2) or r2 >= LIMITE_R2 else f", R² {r2:.4f}")
                             for variacao, r2 in zip(fora_linearidade['variacao_sem_primeiro'], fora_linearidade['R2'])]}),
], ignore_index=True)
outliers_df['origem'] = outliers_df['ID'].map(dados_df.set_index('ID')['origem'])
resumo_outliers_df = outliers_df.groupby('ID').agg(origem=('origem', 'first'), metodos=('metodo', lambda metodos: ', '.join(sorted(set(metodos))))).reset_index()
resumo_outliers_df['interpretacao'] = resumo_outliers_df['ID'].map(INTERPRETACAO_OUTLIERS).fillna('Sem interpretação registrada: analisar.')
residuo_com_origem = modelo_com_origem.resid / modelo_com_origem.resid.std()

# Comparação com e sem outliers. Cada cenário tira um conjunto de molas marcadas e refaz as correlações de
# Spearman (principais e derivadas) e os expoentes do modelo de rigidez. Cenários: sem cada mola marcada, uma
# de cada vez; sem as prováveis de erro de dado (interpretação que começa com "Provável erro de dado"), tiradas
# só das análises que usam a grandeza com erro; e, como teste de estresse, sem todas as marcadas (não serve como
# conclusão, porque tira também valores extremos verdadeiros e encolhe a faixa dos dados). Uma conclusão "muda"
# quando o par troca de significativo para não significativo (ou o contrário), ou quando um par significativo
# troca de sinal. Com menos molas o |ρ| exigido sobe, então parte
# das perdas de significância vem só de ter menos dados. Nenhuma mola sai da análise principal.
ids_marcados = resumo_outliers_df['ID'].tolist()
ids_erro_de_dado = [id_mola for id_mola in ids_marcados if INTERPRETACAO_OUTLIERS.get(id_mola, '').startswith('Provável erro de dado')]
# Cada cenário diz, para cada mola tirada, de quais grandezas ela sai (None = sai de tudo).
cenarios = {'todas as molas': {}, **{f'sem a {id_mola}': {id_mola: None} for id_mola in ids_marcados},
            'sem prováveis erros de dado': {id_mola: GRANDEZAS_AFETADAS_POR_ERRO.get(id_mola) for id_mola in ids_erro_de_dado},
            'teste de estresse: sem todas as marcadas': {id_mola: None for id_mola in ids_marcados}}


def ids_fora(cenario, colunas_usadas):
    """Molas que o cenário tira de uma análise que usa estas colunas."""
    return [id_mola for id_mola, afetadas in cenario.items() if afetadas is None or set(afetadas) & set(colunas_usadas)]


# Todos os pares (grupo de parâmetros, parâmetro, resultado) que entram na comparação.
pares_analisados = [(tabela, parametro, descricao, resultado, coluna_resultado)
                    for tabela, parametros in ((dados_df, PARAMETROS_PRINCIPAIS), (derivadas_df, PARAMETROS_DERIVADOS))
                    for parametro, descricao in parametros.items()
                    for resultado, coluna_resultado in (('Rigidez K', 'rigidez'), ('Dutilidade', 'dutilidade'))]
partes_correlacao, partes_expoentes = [], []
for nome_cenario, cenario in cenarios.items():
    for tabela, parametro, descricao, resultado, coluna_resultado in pares_analisados:
        restantes = tabela[~tabela['ID'].isin(ids_fora(cenario, [parametro, coluna_resultado]))]
        linha = correlacoes_spearman(restantes, {parametro: descricao}).query('resultado == @resultado')
        partes_correlacao.append(linha.assign(cenario=nome_cenario, molas=len(restantes)))
    ficam = ~dados_df['ID'].isin(ids_fora(cenario, [*parametros_modelo, 'rigidez']))
    modelo_cenario = sm.OLS(np.log(dados_df.loc[ficam, 'rigidez']), sm.add_constant(np.log(dados_df.loc[ficam, parametros_modelo]))).fit()
    partes_expoentes.append(modelo_cenario.params[parametros_modelo].rename(nome_cenario))
correlacoes_cenarios_df = pd.concat(partes_correlacao, ignore_index=True)
base = correlacoes_cenarios_df[correlacoes_cenarios_df['cenario'] == 'todas as molas'].set_index(['resultado', 'parametro'])
comparacao = correlacoes_cenarios_df.merge(base[['rho', 'significativo']].reset_index(), on=['resultado', 'parametro'], suffixes=('', '_todas'))
comparacao['muda'] = (comparacao['significativo'] != comparacao['significativo_todas']) | ((np.sign(comparacao['rho']) * np.sign(comparacao['rho_todas']) < 0) & comparacao['significativo_todas'])
mudancas_df = comparacao[comparacao['muda']][['cenario', 'resultado', 'parametro', 'molas', 'rho_todas', 'rho', 'p_valor', 'significativo']]
expoentes_cenarios_df = pd.DataFrame(partes_expoentes)
expoentes_cenarios_df.loc['teoria de viga'] = list(TEORIA_VIGA.values())

teste_rigidez_dutilidade = stats.spearmanr(dados_df['rigidez'], dados_df['dutilidade'])
correlacao_rigidez_dutilidade = pd.DataFrame({'parametro': ['rigidez'], 'rho': [teste_rigidez_dutilidade.statistic],
                                              'p_valor': [teste_rigidez_dutilidade.pvalue],
                                              'significativo': [teste_rigidez_dutilidade.pvalue < NIVEL_SIGNIFICANCIA]})

#conversão de dataframe
dados_dict = dados_df.to_dict('list')


#processamento de dados
#calculo de correlação entre os parametros
#na matriz de  correlação devemos colocar as variaves que queremos estudar
# As variáveis estudadas são escolhidas pelo nome da coluna, e o .corr() do pandas calcula a matriz de
# Pearson já com os nomes nas linhas e colunas (o heatmap passa a mostrar os nomes nos eixos).
matriz_pearson = dados_df[['h', 'L', 'm', 'rigidez', 'dutilidade']]
#Calculo
resultado_pearson = matriz_pearson.corr()


#apresentação
print(dados_df)
# Console enxuto: avisos, outliers, o que muda com e sem eles e o resumo do comportamento geral. As tabelas
# completas (estatísticas, correlações, matrizes, elasticidades, modelos) vão para a planilha de resultados.
print("AVISO: molas sem IC e sem R² por ter poucos pontos:", por_mola_df.loc[por_mola_df['pontos'] < MINIMO_PONTOS_POR_MOLA, 'ID'].tolist())
print("AVISO: C suspeito nas molas", IDS_C_SUSPEITO, "(afeta a dutilidade; ver CLAUDE.md)")
print(f"\nOutliers (boxplot: {FATOR_IQR}·IQR; resíduo: {LIMITE_RESIDUO} desvios-padrão; linearidade: R² < {LIMITE_R2} ou variação > {LIMITE_VARIACAO_PONTOS:.0%})")
print(outliers_df[['ID', 'origem', 'metodo', 'motivo']].sort_values(['ID', 'metodo']).to_string(index=False))
print("\nResumo por mola, com a interpretação")
for _, mola in resumo_outliers_df.iterrows():
    print(f"  Mola {mola['ID']} ({mola['origem']}; {mola['metodos']}): {mola['interpretacao']}")
print("Com a variável de origem no modelo, nenhuma mola passa de 2 desvios-padrão no resíduo (maior: mola "
      f"{dados_df['ID'][residuo_com_origem.abs().idxmax()]}, {residuo_com_origem[residuo_com_origem.abs().idxmax()]:+.2f})."
      if residuo_com_origem.abs().max() <= LIMITE_RESIDUO else "Com a variável de origem, ainda há molas acima de 2 desvios-padrão.")
print("Nenhuma mola é descartada: elas só são marcadas.")
print("\nCom e sem outliers: conclusões de correlação que mudam (significância ou sinal) em relação a todas as molas")
print(mudancas_df.to_string(index=False, float_format='{:.4f}'.format) if len(mudancas_df) else "  nenhuma")
mudancas_conclusivas = mudancas_df[~mudancas_df['cenario'].str.startswith('teste de estresse')]  # o teste de estresse não entra na conclusão
pares_estaveis = base[base['significativo']].index.difference(mudancas_conclusivas.set_index(['resultado', 'parametro']).index)
print("Pares significativos que não mudam em nenhum cenário (fora o teste de estresse):", ', '.join(f"{resultado} × {parametro}" for resultado, parametro in pares_estaveis))
print(f"Atenção: com menos molas o |ρ| exigido sobe (no teste de estresse ficam {len(dados_df) - len(ids_marcados)} molas e o limite vai a {rho_minimo_significativo(len(dados_df) - len(ids_marcados)):.2f}); parte das perdas vem só disso.")

# Resumo do comportamento geral, em linguagem simples. Todos os números vêm das variáveis calculadas acima,
# então o texto se atualiza sozinho se a planilha mudar. As molas marcadas são agrupadas pelo começo da
# interpretação (comportamento real, provável erro de dado, efeito de medida ou incerto).
mais_rigida, menos_rigida = dados_df.loc[dados_df['rigidez'].idxmax()], dados_df.loc[dados_df['rigidez'].idxmin()]
mais_dutil, menos_dutil = dados_df.loc[dados_df['dutilidade'].idxmax()], dados_df.loc[dados_df['dutilidade'].idxmin()]
expoente = elasticidades_df.set_index('parametro')['expoente']
cenarios_conclusivos = [nome for nome in cenarios if not nome.startswith('teste de estresse')]
pares_frageis = mudancas_df[mudancas_df['cenario'].isin(cenarios_conclusivos)][['resultado', 'parametro']].drop_duplicates()
grupos_interpretacao = {rotulo: resumo_outliers_df.loc[resumo_outliers_df['interpretacao'].str.startswith(inicio), 'ID'].tolist()
                        for rotulo, inicio in (('comportamento real', 'Comportamento real'), ('provável erro de dado', 'Provável erro'),
                                               ('efeito de medida', 'Efeito de medida'), ('incerto', 'Incerto'))}
molas_sem_r2 = por_mola_df.loc[por_mola_df['R2'].isna(), 'ID'].tolist()
dois_que_mais_pesam = elasticidades_df.reindex(elasticidades_df['expoente'].abs().sort_values(ascending=False).index)['descricao'].head(2).str.split(' ').str[0].str.lower().tolist()
robustos_dutilidade = [parametro for resultado, parametro in pares_estaveis if resultado == 'Dutilidade']
significativo_com_rigidez = correlacoes_principais_df.query("resultado == 'Rigidez K'").set_index('parametro')['significativo']
efeito_claro = (elasticidades_df['IC_inf'] > 0) | (elasticidades_df['IC_sup'] < 0) #IC do expoente não passa pelo zero
escondidos_no_rho = elasticidades_df[~elasticidades_df['parametro'].map(significativo_com_rigidez) & efeito_claro]
alongamento_falha = dados_df['dutilidade'] * dados_df['C'] #alongamento na falha em mm, sem dividir por C
alongamento_com_h = stats.spearmanr(dados_df['h'], alongamento_falha)
alongamento_com_fio = stats.spearmanr(derivadas_df['fio_desenvolvido'], alongamento_falha)
compara_origens = {coluna: stats.mannwhitneyu(*[dados_df.loc[dados_df['origem'] == origem, coluna] for origem in ('professor', 'grupo')]).pvalue for coluna in ('rigidez', 'dutilidade')}
medianas_por_espiras = dados_df.groupby('n').agg(molas=('ID', 'size'), rigidez_mediana=('rigidez', 'median'), dutilidade_mediana=('dutilidade', 'median'))
print("\n" + "=" * 100)
print("COMPORTAMENTO GERAL DO SISTEMA")
print("=" * 100)
print(f"1. Sob carga, as molas se comportam como molas lineares até a falha: R² entre {por_mola_df['R2'].min():.4f} e {por_mola_df['R2'].max():.4f}"
      + (f" (a mola {', '.join(map(str, molas_sem_r2))} tem só 2 pontos)." if molas_sem_r2 else "."))
print(f"2. A rigidez vai de {menos_rigida['rigidez']:.4f} (mola {menos_rigida['ID']}) a {mais_rigida['rigidez']:.4f} kgf/mm (mola {mais_rigida['ID']}), "
      f"{mais_rigida['rigidez'] / menos_rigida['rigidez']:.0f} vezes; a dutilidade vai de {menos_dutil['dutilidade']:.0%} (mola {menos_dutil['ID']}) "
      f"a {mais_dutil['dutilidade']:.0%} (mola {mais_dutil['ID']}). Como os valores são assimétricos, a mediana representa melhor: "
      f"{dados_df['rigidez'].median():.4f} kgf/mm e {dados_df['dutilidade'].median():.0%}.")
print(f"3. A rigidez segue a teoria de viga (K proporcional a t·h³ / (n·Ltot³)): o modelo explica {modelo_rigidez.rsquared:.1%} da variação. "
      f"Com os outros fixos, +1% em h dá {expoente['h']:+.1f}% em K, +1% em Ltot dá {expoente['Ltot']:+.1f}%, +1% em n dá {expoente['n']:+.1f}% "
      f"e +1% em t dá {expoente['t']:+.1f}%. O que mais pesa: {' e '.join(dois_que_mais_pesam)}.")
print("4. A dutilidade (alongamento na falha ÷ C) cai quando h aumenta e sobe com a esbeltez (Ltot/h). Cuidado: o alongamento em mm")
print(f"   quase não depende de h (ρ = {alongamento_com_h.statistic:+.2f}); ele cresce com o comprimento de fio que trabalha (fio desenvolvido, "
      f"ρ = {alongamento_com_fio.statistic:+.2f}). O efeito de h na dutilidade vem em boa parte de C, que está no denominador e cresce com h.")
print("   Relações com a dutilidade que valem em todos os cenários com e sem outliers (fora o teste de estresse): " + (', '.join(robustos_dutilidade) or 'nenhuma') + ".")
print(f"5. Rigidez e dutilidade andam em sentidos opostos (ρ = {teste_rigidez_dutilidade.statistic:+.2f}): a geometria que deixa a mola rígida")
print("   é a mesma que limita o alongamento, e a dutilidade ainda divide por C. Não é o material ficando frágil.")
print("6. Correlação simples engana: " + ("; ".join(f"{linha.parametro} não é significativo no ρ ({linha.rho_simples:+.2f}), mas com os outros fixos "
      f"tem efeito claro (expoente {linha.expoente:+.2f})" for linha in escondidos_no_rho.itertuples()) or "nenhum parâmetro fica escondido no ρ")
      + ". Parâmetros que andam juntos no desenho das molas (L e Ltot; h e C) não se separam.")
print(f"7. Em valores absolutos, professor e grupo não diferem (Mann-Whitney: rigidez p = {compara_origens['rigidez']:.2f}, dutilidade "
      f"p = {compara_origens['dutilidade']:.2f}), mas com a mesma geometria as molas do professor ficam "
      f"{np.exp(modelo_com_origem.params['professor']) - 1:.0%} mais rígidas. Possíveis causas: h do grupo medido acima do que trabalha")
print("   (mediana de várias medidas em peças com falhas) ou diferença de impressão; estes dados não separam as duas.")
print(f"8. Outliers: {len(ids_marcados)} molas marcadas, nenhuma descartada. " + "; ".join(
      f"{rotulo}: {', '.join(map(str, ids))}" for rotulo, ids in grupos_interpretacao.items() if ids) + ".")
print("   Conclusões que dependem de uma mola ou das prováveis de erro: "
      + (', '.join(f"{resultado} × {parametro}" for resultado, parametro in pares_frageis.itertuples(index=False)) or 'nenhuma') + ".")
print("=" * 100)

#construção de algumas dispersões
#a) rigidez em função da espessura do fio "h"
plt.scatter(dados_dict['h'], dados_dict['rigidez'])
for id_mola, h, rigidez in zip(dados_dict['ID'], dados_dict['h'], dados_dict['rigidez']):
    plt.annotate(str(id_mola), (h, rigidez), xytext=(5, 5), textcoords="offset points")   #rotilos nas bolinhas dos graficos pa a grafico hxrigidez

# Rótulos dos eixos corrigidos, com o nome certo da grandeza e a unidade.
plt.xlabel("Espessura do fio - h (mm)")
plt.ylabel("Rigidez da mola - K (kgf/mm)")
plt.show()

#b) dutilidade em função do passo "L"
plt.scatter(dados_dict['L'], dados_dict['dutilidade'])
for id_mola, L, dutilidade in zip(dados_dict['ID'], dados_dict['L'], dados_dict['dutilidade']):
    plt.annotate(str(id_mola), (L, dutilidade), xytext=(5, 5), textcoords="offset points")
# Rótulos dos eixos corrigidos, com o nome certo da grandeza e a unidade.
plt.xlabel("Passo da mola - L (mm)")
plt.ylabel("Dutilidade (fração de C)")
plt.show()

#c) dutilidade em funcao da massa "m"
plt.scatter(dados_dict['m'], dados_dict['dutilidade'])
for id_mola, m, dutilidade in zip(dados_dict['ID'], dados_dict['m'], dados_dict['dutilidade']):
    plt.annotate(str(id_mola), (m, dutilidade), xytext=(5, 5), textcoords="offset points")
# Rótulos dos eixos corrigidos, com o nome certo da grandeza e a unidade.
plt.xlabel("Massa da mola - m (g)")
plt.ylabel("Dutilidade (fração de C)")
plt.show()

#cprrelação de pearson
print(resultado_pearson)
sb.heatmap(resultado_pearson)
plt.show()

# Cria a pasta saidas/ (se ainda não existir) e gera os gráficos da análise de correlação: barras de
# influência dos parâmetros principais e das derivadas, e as dispersões dos parâmetros principais contra a
# rigidez e contra a dutilidade (as derivadas aparecem nas barras de influência).
os.makedirs(PASTA_SAIDAS, exist_ok=True)
grafico_barras_correlacao(correlacoes_principais_df, rho_minimo, "Influência dos parâmetros principais", "correlacao_parametros_principais.png")
grafico_barras_correlacao(correlacoes_derivadas_df, rho_minimo, "Influência das dimensões derivadas", "correlacao_dimensoes_derivadas.png")
com_rigidez = "resultado == 'Rigidez K'"
com_dutilidade = "resultado == 'Dutilidade'"
grafico_dispersoes(dados_df, PARAMETROS_PRINCIPAIS, 'rigidez', 'Rigidez K (kgf/mm)', correlacoes_principais_df.query(com_rigidez), "dispersao_principais_rigidez.png")
grafico_dispersoes(dados_df, PARAMETROS_PRINCIPAIS, 'dutilidade', 'Dutilidade (fração de C)', correlacoes_principais_df.query(com_dutilidade), "dispersao_principais_dutilidade.png")
# Gráficos da relação entre parâmetros: as duas matrizes de correlação e a dispersão rigidez × dutilidade.
grafico_matriz(matriz_principais_rho, matriz_principais_p, list(PARAMETROS_PRINCIPAIS.values()), "Correlação entre os parâmetros principais", "matriz_parametros_principais.png")
grafico_matriz(matriz_derivadas_rho, matriz_derivadas_p, list(PARAMETROS_DERIVADOS.values()), "Correlação entre as dimensões derivadas", "matriz_dimensoes_derivadas.png")
grafico_dispersoes(dados_df, {'rigidez': 'Rigidez K (kgf/mm)'}, 'dutilidade', 'Dutilidade (fração de C)', correlacao_rigidez_dutilidade, "rigidez_x_dutilidade.png")
# Gráficos das elasticidades: expoentes medidos contra a teoria de viga e rigidez prevista × medida.
grafico_elasticidades(elasticidades_df, "elasticidades_rigidez.png")
grafico_previsto_medido(dados_df, rigidez_prevista, modelo_rigidez.rsquared, "rigidez_prevista_x_medida.png")
# Gráficos dos outliers: boxplots (parâmetros principais e derivadas), resíduo do modelo por mola e as curvas
# carga × alongamento de cada ensaio.
grafico_boxplots(dados_df, rotulos_boxplot, outliers_df[outliers_df['metodo'] == '1. boxplot'], "Rigidez, dutilidade e parâmetros principais", "boxplots_principais.png")
grafico_boxplots(derivadas_df, PARAMETROS_DERIVADOS, outliers_df[outliers_df['metodo'] == '1. boxplot (derivadas)'], "Dimensões derivadas", "boxplots_derivadas.png")
grafico_residuos(dados_df, residuo_padronizado, "residuos_modelo_rigidez.png")
grafico_curvas_ensaio(ensaios_df, por_mola_df, "curvas_carga_alongamento.png")
# Gráfico da comparação com e sem as molas marcadas (parâmetros principais).
grafico_com_sem_outliers(correlacoes_cenarios_df, PARAMETROS_PRINCIPAIS, rho_minimo, "Correlações com e sem outliers: parâmetros principais", "com_sem_outliers_principais.png")
# Gráfico final: rigidez e dutilidade por número de espiras.
grafico_por_espiras(dados_df, "por_numero_de_espiras.png")

# Planilha de resultados em saidas/ (a planilha de entrada não é alterada). Uma aba por tabela calculada:
# rigidez e dutilidade de cada mola com as dimensões, derivadas, estatísticas, correlações, matrizes,
# pares significativos, elasticidades, modelos, outliers, comparação com e sem outliers e medianas por n.
arquivo_resultados = os.path.join(PASTA_SAIDAS, "experimento_resultados.xlsx")
with pd.ExcelWriter(arquivo_resultados) as planilha_resultados:
    dados_df.drop(columns=['rigidez', 'dutilidade']).merge(por_mola_df, on=['ID', 'origem']).to_excel(planilha_resultados, sheet_name='resultados_por_mola', index=False)
    derivadas_df.to_excel(planilha_resultados, sheet_name='dimensoes_derivadas', index=False)
    pd.concat([estatisticas_df, estatisticas_derivadas_df]).to_excel(planilha_resultados, sheet_name='estatisticas', index=False)
    pd.concat([correlacoes_principais_df, correlacoes_derivadas_df]).to_excel(planilha_resultados, sheet_name='correlacoes', index=False)
    matriz_principais_rho.to_excel(planilha_resultados, sheet_name='matriz_principais')
    matriz_derivadas_rho.to_excel(planilha_resultados, sheet_name='matriz_derivadas')
    elasticidades_df.to_excel(planilha_resultados, sheet_name='elasticidades', index=False)
    outliers_df.to_excel(planilha_resultados, sheet_name='outliers', index=False)
    resumo_outliers_df.to_excel(planilha_resultados, sheet_name='outliers_resumo', index=False)
    mudancas_df.to_excel(planilha_resultados, sheet_name='com_sem_outliers', index=False)
    expoentes_cenarios_df.to_excel(planilha_resultados, sheet_name='expoentes_por_cenario')
    pd.concat({'principais': pares_significativos(matriz_principais_rho, matriz_principais_p),
               'derivadas': pares_significativos(matriz_derivadas_rho, matriz_derivadas_p)}).to_excel(planilha_resultados, sheet_name='pares_significativos')
    pd.concat([modelo_com_origem.params.rename('coeficiente'), intervalo_com_origem.rename(columns={0: 'IC_inf', 1: 'IC_sup'})],
              axis=1).to_excel(planilha_resultados, sheet_name='modelo_com_origem')
    pd.DataFrame({'b1': [modelo_indice.params['indice_rigidez']], 'IC_inf': [intervalo_indice[0]], 'IC_sup': [intervalo_indice[1]],
                  'R2': [modelo_indice.rsquared]}).to_excel(planilha_resultados, sheet_name='modelo_indice', index=False)
    razao_rigidez_indice.to_excel(planilha_resultados, sheet_name='razao_K_indice')
    medianas_por_espiras.to_excel(planilha_resultados, sheet_name='por_numero_de_espiras')
print(f"\nPlanilha de resultados gravada em {arquivo_resultados}")
