import pandas as pd # leitura de excel
import openpyxl #biblioteca auxiliar para ler excel
import seaborn as sb
# os cria a pasta saidas/ e monta o caminho dos arquivos de gráfico.
import os
import matplotlib.pyplot as plt
import numpy as np
# scipy.stats faz o intervalo de confiança (t de Student), a média geométrica, a regressão linear (R²) e a
# correlação de Spearman. adjust_text afasta os
# rótulos dos pontos para um não cobrir o outro. Line2D, Patch, NullLocator e LinearSegmentedColormap só
# montam legendas, eixos e a escala de cor dos gráficos.
from scipy import stats
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
# Limites da marcação de outliers: regra do boxplot (1,5 × intervalo entre quartis) e linearidade do ensaio (R²
# abaixo de 0,98 ou rigidez ponto a ponto variando mais de 20% dentro da mesma mola, medida como maior ÷ menor − 1
# sem o primeiro ponto, porque o primeiro ponto tem carga baixa e uma pequena folga nele infla a variação).
FATOR_IQR = 1.5
LIMITE_R2 = 0.98
LIMITE_VARIACAO_PONTOS = 0.20
# Rigidez principal de cada mola: 'media_pontos' = média de carga ÷ alongamento ponto a ponto (fórmula da
# planilha do professor); 'regressao' = inclinação da reta carga × alongamento pela origem. As duas são sempre
# calculadas e gravadas na planilha de resultados; esta opção só escolhe qual segue para as análises.
RIGIDEZ_PRINCIPAL = 'media_pontos'
# Correlação de Pearson: |r| a partir de 0,70 é tratado como relação forte (Akoglu, 2018), o mesmo limite do
# artigo do professor.
LIMITE_PEARSON = 0.70
# Regra de eliminação de outliers: a mola sai da análise se ficar fora do limite em pelo menos este número de
# boxplots das grandezas de conferência (regra do artigo do professor).
MINIMO_BOXPLOTS_PARA_ELIMINAR = 2
# Espessura de projeto (nominal) de cada mola, em mm. Foi deduzida do comprimento por espira (C ÷ n), que cai em
# três valores: cerca de 41, 27 e 21 mm. Serve só para comparar com o h medido; os cálculos usam o h medido.
H_NOMINAL = {3.0: [107, 101, 113, 122, 73, 71, 60, 36], 2.0: [111, 82, 94, 12, 80], 1.5: [105, 112, 56, 65]}
# Planilha com as leituras individuais das molas do grupo (aba 'dimensoes'). Se o arquivo não estiver na pasta,
# o programa roda sem as colunas de leituras. Leituras eliminadas por erro de registro: (mola, grandeza, posição).
ARQUIVO_LEITURAS = "molas_dados_brutos.xlsx"
LEITURAS_ELIMINADAS = {(60, 'h', 1): "6,25 mm entre leituras de 3,20 a 3,26 mm: erro de digitação"}
# Interpretação de cada mola apontada como outlier, escrita depois de olhar os dados e conferida no loop de
# conferência. Mola apontada que não estiver aqui aparece como "sem interpretação registrada".
INTERPRETACAO_OUTLIERS = {
    107: "Comportamento real: junta tudo o que aumenta K (fio grosso, mola alta, só 3 espiras, largura pequena). A rigidez ÷ índice de rigidez dela fica dentro do limite, então é mola rígida por projeto, não defeito.",
    56: "Comportamento real, com ressalva: é a mola mais esbelta (Ltot/h = 32,3), que alonga muito. A folha do ensaio registra plastificação em 1,1 kgf sem quebra, então os 150 mm não são fratura como nas outras molas e a dutilidade dela não é diretamente comparável.",
    71: "Provável erro de dado: C = 260,1 idêntico ao da 36 e fora do padrão C/n das molas de h parecido. Afeta só a dutilidade.",
    36: "Provável erro de dado: C = 260,1 idêntico ao da 71 e fora do padrão C/n das molas de h parecido. Afeta só a dutilidade.",
    94: "Provável erro de dado na massa: o índice de densidade é cerca do dobro do das demais e a massa ÷ t fica muito acima da linha das outras molas com fio parecido, enquanto a rigidez segue o modelo.",
    112: "Efeito de medida em carga baixa: a variação vem do primeiro ponto (0,1 kgf), onde uma pequena folga pesa muito.",
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
    'densidade_aparente': 'Índice de densidade (g/cm³)',
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
ensaio_df['rigidez_regressao'] = pontos_por_mola.apply(lambda pontos: np.linalg.lstsq(pontos[['alongamento']], pontos['carga'], rcond=None)[0][0])
# Rigidez principal da mola, conforme RIGIDEZ_PRINCIPAL, e a diferença relativa entre as duas fórmulas (mostra
# que a escolha não muda as conclusões).
ensaio_df['rigidez'] = ensaio_df['media_pontos'] if RIGIDEZ_PRINCIPAL == 'media_pontos' else ensaio_df['rigidez_regressao']
ensaio_df['diferenca_formulas'] = ensaio_df['media_pontos'] / ensaio_df['rigidez_regressao'] - 1
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
    ensaio_df[['pontos', 'R2', 'media_pontos', 'rigidez_regressao', 'diferenca_formulas', 'desvio_pontos', 'IC_pontos_inf', 'IC_pontos_sup', 'variacao_pontos', 'variacao_sem_primeiro']], left_on='ID', right_index=True)

# Calcula as 5 dimensões derivadas a partir das medidas da tabela1, com as fórmulas escritas por extenso
# (pandas eval). Elas ficam numa tabela separada (derivadas_df, junto com origem, rigidez e
# dutilidade de cada mola) para não se misturarem com os
# parâmetros principais: índice de rigidez (rigidez esperada pela teoria de viga), esbeltez do braço (o índice
# de dutilidade saiu por ser praticamente a mesma medida que a esbeltez, ρ 0,99), fio desenvolvido (comprimento aproximado de fio), momento de inércia da seção e
# índice de densidade (massa ÷ volume aproximado; serve para checar se massa e geometria combinam). O volume usa só
# os braços retos (ignora os arcos e os blocos das pontas), então o índice fica acima da densidade real do PLA
# (cerca de 1,24 g/cm³) e só vale para comparar uma mola com as outras.
derivadas_df = dados_df[['ID', 'origem', 'rigidez', 'dutilidade']].assign(
    indice_rigidez=dados_df.eval('t * h**3 / (n * Ltot**3)'),
    esbeltez=dados_df.eval('Ltot / h'),
    fio_desenvolvido=dados_df.eval('2 * n * Ltot'),
    momento_inercia=dados_df.eval('t * h**3 / 12'),
    densidade_aparente=dados_df.eval('m / (t * h * 2 * n * Ltot) * 1000'),
)
# Grandezas de conferência da regra de eliminação de outliers. Todas deveriam ser parecidas entre molas sem
# defeito: rigidez e dutilidade medidas no ensaio, rigidez ÷ índice de rigidez (tira o efeito da geometria; o que
# sobra depende do material) e o índice de densidade. As dimensões de projeto (h, t, L, n, Ltot, C, m) ficam de
# fora porque variam de propósito: um boxplot delas marcaria o projeto da mola, não um defeito.
ROTULOS_CONFERENCIA = {
    'rigidez': 'Rigidez K (kgf/mm)',
    'dutilidade': 'Dutilidade (fração de C)',
    'rigidez_normalizada': 'Rigidez ÷ índice de rigidez (kgf/mm²)',
    'densidade_aparente': 'Índice de densidade (g/cm³)',
}
conferencia_df = dados_df[['ID', 'origem', 'rigidez', 'dutilidade']].assign(
    rigidez_normalizada=dados_df['rigidez'] / derivadas_df['indice_rigidez'],
    densidade_aparente=derivadas_df['densidade_aparente'],
)

# Medidas de h de cada mola: o valor usado nos cálculos (mediana das leituras nas molas do grupo, valor único
# nas do professor), o h nominal e o desvio entre os dois. Quando a planilha de leituras está na pasta, entram
# também a quantidade de leituras, o mínimo, o máximo e a amplitude (máximo − mínimo), que mostra o quanto o h
# varia ao longo da mola. As molas do professor têm uma leitura só, então ficam sem amplitude.
nominal_por_mola = {id_mola: nominal for nominal, ids_molas in H_NOMINAL.items() for id_mola in ids_molas}
medidas_h_df = dados_df[['ID', 'origem', 'h']].assign(h_nominal=dados_df['ID'].map(nominal_por_mola))
medidas_h_df['h_desvio_nominal'] = medidas_h_df['h'] - medidas_h_df['h_nominal']
if os.path.exists(ARQUIVO_LEITURAS):
    leituras_df = pd.read_excel(ARQUIVO_LEITURAS, sheet_name='dimensoes')
    eliminada = [chave in LEITURAS_ELIMINADAS for chave in zip(leituras_df['id'], leituras_df['grandeza'], leituras_df['posicao'])]
    leituras_h = leituras_df[(leituras_df['grandeza'] == 'h') & ~np.array(eliminada)]
    resumo_leituras = leituras_h.groupby('id')['valor_mm'].agg(h_leituras='size', h_minimo='min', h_maximo='max')
    resumo_leituras['h_amplitude'] = resumo_leituras['h_maximo'] - resumo_leituras['h_minimo']
    medidas_h_df = medidas_h_df.merge(resumo_leituras, left_on='ID', right_index=True, how='left')
    medidas_h_df['h_leituras'] = medidas_h_df['h_leituras'].fillna(1).astype(int)
else:
    print(f"AVISO: {ARQUIVO_LEITURAS} não está na pasta; a tabela de medidas de h sai sem as colunas de leituras.")
desvio_h_por_origem = medidas_h_df.groupby('origem')['h_desvio_nominal'].agg(['mean', 'min', 'max'])
# Mesmo índice de rigidez, calculado com o h nominal no lugar do medido. Se a reta K × índice ajusta melhor com
# o nominal, o h medido está carregando um erro que não existe na mola.
indice_rigidez_nominal = dados_df['t'] * medidas_h_df['h_nominal']**3 / (dados_df['n'] * dados_df['Ltot']**3)
r2_com_h_nominal = np.corrcoef(indice_rigidez_nominal, dados_df['rigidez'])[0, 1] ** 2


# Termina todo gráfico novo do mesmo jeito: ajusta as margens (margem_inferior reserva espaço para a legenda
# embaixo), salva o PNG em saidas/ e mostra na tela.
def salvar_grafico(figura, nome_arquivo, margem_inferior=0):
    figura.tight_layout(rect=(0, margem_inferior, 1, 1))
    figura.savefig(os.path.join(PASTA_SAIDAS, nome_arquivo), dpi=150)
    plt.show()
    plt.close(figura) #libera a figura depois de salva e mostrada


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
def grafico_matriz(rho, p_valor, rotulos, titulo, nome_arquivo, rotulo_escala='ρ de Spearman', nota='* = significativo, p < 0,05', destaque=None):
    destaque = (p_valor < NIVEL_SIGNIFICANCIA) if destaque is None else destaque #quais células levam o asterisco
    textos = rho.map(lambda valor: f"{valor:+.2f}".replace('.', ',')) + destaque.map(lambda marcado: '*' if marcado else '')
    rho, textos = rho.iloc[1:, :-1], textos.iloc[1:, :-1] #tira a 1ª linha e a última coluna, que só teriam a diagonal
    acima_da_diagonal = np.triu(np.ones(rho.shape, dtype=bool), k=1)
    figura, eixo = plt.subplots(figsize=(1.2 * len(rotulos) + 3, 0.9 * len(rotulos) + 2))
    sb.heatmap(rho, mask=acima_da_diagonal, annot=textos, fmt='', cmap=ESCALA_CORRELACAO, vmin=-1, vmax=1, center=0,
               linewidths=2, linecolor='white', cbar_kws={'label': rotulo_escala}, ax=eixo,
               xticklabels=rotulos[:-1], yticklabels=rotulos[1:])
    eixo.tick_params(axis='x', rotation=35)
    plt.setp(eixo.get_xticklabels(), ha='right')
    eixo.set_title(f"{titulo}\n({nota})")
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
# os prováveis erros de dado e sem as eliminadas pela regra (quando houver). Marcadores próximos = conclusão estável; marcador que
# cruza a faixa cinza (não significativo) = conclusão que depende das molas marcadas.
def grafico_com_sem_outliers(correlacoes_cenarios, parametros, rho_minimo, titulo, nome_arquivo):
    cenarios_mostrados = {'todas as molas': ('o', '#222222'), 'sem prováveis erros de dado': ('s', '#777777'), 'sem as eliminadas pela regra dos 2 boxplots': ('^', '#aaaaaa')}
    figura, eixos = plt.subplots(1, 2, figsize=(11, 1.9 + 0.5 * len(parametros)), sharey=True)
    posicoes = np.arange(len(parametros))[::-1]
    for eixo, resultado in zip(eixos, ['Rigidez K', 'Dutilidade']):
        eixo.axvspan(-rho_minimo, rho_minimo, color='#eeeeee', zorder=0)
        for deslocamento, (cenario, (marcador, cor)) in zip((0.2, 0, -0.2), cenarios_mostrados.items()):
            linhas = correlacoes_cenarios[(correlacoes_cenarios['cenario'] == cenario) & (correlacoes_cenarios['resultado'] == resultado)].set_index('parametro')
            if linhas.empty: #cenário que não existe nesta rodada (ex.: nenhuma mola eliminada pela regra)
                continue
            linhas = linhas.loc[list(parametros)]
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


# Matriz de Pearson entre as colunas pedidas: o r de cada par (pandas .corr) e o p-valor (scipy pearsonr).
# Pearson mede o quanto dois valores andam juntos em linha reta; é o coeficiente do método do professor.
def matriz_pearson(tabela, colunas):
    r = tabela[colunas].corr(method='pearson')
    p_valor = tabela[colunas].corr(method=lambda x, y: stats.pearsonr(x, y).pvalue)
    return r, p_valor


# Correlação de Pearson de cada parâmetro com a rigidez e com a dutilidade. A coluna 'forte' marca |r| ≥ 0,70.
def correlacoes_pearson(tabela, parametros):
    colunas = [*parametros, 'rigidez', 'dutilidade']
    r, p_valor = matriz_pearson(tabela, colunas)
    partes = [pd.DataFrame({'resultado': resultado, 'parametro': list(parametros), 'descricao': list(parametros.values()),
                            'r': r.loc[list(parametros), coluna].to_numpy(), 'p_valor': p_valor.loc[list(parametros), coluna].to_numpy()})
              for resultado, coluna in (('Rigidez K', 'rigidez'), ('Dutilidade', 'dutilidade'))]
    correlacoes = pd.concat(partes, ignore_index=True)
    correlacoes['forte'] = correlacoes['r'].abs() >= LIMITE_PEARSON
    return correlacoes


# Histogramas: um painel por grandeza, com as barras empilhadas por origem (azul = grupo, laranja = professor).
# A linha cheia é a média e a tracejada é a mediana; quando as duas se afastam, a distribuição é assimétrica.
def grafico_histogramas(tabela, rotulos, titulo, nome_arquivo):
    colunas_figura = 3
    linhas_figura = int(np.ceil(len(rotulos) / colunas_figura))
    figura, eixos = plt.subplots(linhas_figura, colunas_figura, figsize=(12, 3.1 * linhas_figura + 1))
    eixos = eixos.flatten()
    for eixo, (coluna, rotulo) in zip(eixos, rotulos.items()):
        sb.histplot(tabela, x=coluna, hue='origem', palette=COR_ORIGEM, multiple='stack', bins='sturges', alpha=1, edgecolor='white', legend=False, ax=eixo)
        eixo.axvline(tabela[coluna].mean(), color='#222222', linewidth=1.3)
        eixo.axvline(tabela[coluna].median(), color='#222222', linewidth=1.3, linestyle='--')
        eixo.set(xlabel=rotulo, ylabel='Molas')
    for eixo in eixos[len(rotulos):]:
        eixo.set_visible(False)
    legenda = [Patch(color=cor, label=f"Molas: {origem}") for origem, cor in COR_ORIGEM.items()] + [
        Line2D([0], [0], color='#222222', label='Média'), Line2D([0], [0], color='#222222', linestyle='--', label='Mediana')]
    figura.legend(handles=legenda, loc='lower center', ncol=4, frameon=False)
    figura.suptitle(titulo)
    salvar_grafico(figura, nome_arquivo, 0.05)


# Dispersão da rigidez contra o índice de rigidez, com a reta ajustada por mínimos quadrados (np.polyfit) e o
# R². Eixos em escala comum (não log): é a regressão linear simples do método do professor.
def grafico_regressao_simples(tabela, inclinacao, intercepto, r2, nome_arquivo):
    figura, eixo = plt.subplots(figsize=(7.5, 5.5))
    sb.scatterplot(tabela, x='indice_rigidez', y='rigidez', hue='origem', palette=COR_ORIGEM, s=45, zorder=3, ax=eixo)
    valores_x = np.array([0, tabela['indice_rigidez'].max() * 1.05])
    eixo.plot(valores_x, inclinacao * valores_x + intercepto, color='#777777', linewidth=1.5, label='Reta ajustada (mínimos quadrados)')
    eixo.set(xlabel='Índice de rigidez t·h³ / (n·Ltot³) (mm)', ylabel='Rigidez K (kgf/mm)',
             title=f"Rigidez em função do índice de rigidez (R² = {r2:.3f})".replace('.', ','))
    eixo.legend(frameon=False, title=None)
    figura.tight_layout()
    rotulos = [eixo.text(x, y, str(id_mola), fontsize=7, color='#444444') for id_mola, x, y in zip(tabela['ID'], tabela['indice_rigidez'], tabela['rigidez'])]
    adjust_text(rotulos, x=tabela['indice_rigidez'].to_numpy(), y=tabela['rigidez'].to_numpy(), ax=eixo)
    salvar_grafico(figura, nome_arquivo)


# Desvio do h medido em relação ao nominal, uma barra por mola (azul = grupo, laranja = professor). Barras para
# cima = h medido acima do nominal. Se as duas origens ficam de lados diferentes, a diferença é da medição.
def grafico_desvio_h(medidas_h, nome_arquivo):
    desvios = medidas_h.assign(mola=medidas_h['ID'].astype(str)).sort_values('h_desvio_nominal')
    figura, eixo = plt.subplots(figsize=(10, 4.5))
    sb.barplot(desvios, x='mola', y='h_desvio_nominal', hue='origem', palette=COR_ORIGEM, saturation=1, dodge=False, ax=eixo)
    eixo.axhline(0, color='#555555', linewidth=1)
    eixo.set(xlabel='ID da mola', ylabel='h medido − h nominal (mm)', title='Desvio do h medido em relação ao nominal, por mola')
    eixo.legend(title=None, frameon=False)
    salvar_grafico(figura, nome_arquivo)


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

# Pearson (método do professor): matriz entre os parâmetros medidos, a rigidez e a dutilidade, e a correlação de
# cada parâmetro (medido ou derivado) com os dois resultados. Fica ao lado do Spearman, que não exige linha reta.
colunas_pearson = [*PARAMETROS_PRINCIPAIS, 'rigidez', 'dutilidade']
matriz_pearson_r, matriz_pearson_p = matriz_pearson(dados_df, colunas_pearson)
pearson_df = pd.concat([correlacoes_pearson(dados_df, PARAMETROS_PRINCIPAIS).assign(grupo='principais'),
                        correlacoes_pearson(derivadas_df, PARAMETROS_DERIVADOS).assign(grupo='derivadas')], ignore_index=True)
# Regressão linear simples da rigidez contra o índice de rigidez (np.polyfit, mínimos quadrados) e o R², em
# escala comum. O índice foi escolhido pela teoria de viga, como a área foi escolhida no artigo do professor.
inclinacao_simples, intercepto_simples = np.polyfit(derivadas_df['indice_rigidez'], derivadas_df['rigidez'], 1)
r2_simples = np.corrcoef(derivadas_df['indice_rigidez'], derivadas_df['rigidez'])[0, 1] ** 2


def pares_significativos(rho, p_valor):
    abaixo_da_diagonal = np.tril(np.ones(rho.shape, dtype=bool), k=-1)
    pares = pd.DataFrame({'rho': rho.where(abaixo_da_diagonal).stack(), 'p_valor': p_valor.where(abaixo_da_diagonal).stack()})
    return pares[pares['p_valor'] < NIVEL_SIGNIFICANCIA].sort_values('rho', key=abs, ascending=False)


# Marcação de outliers por dois métodos. Método 1: regra do boxplot em rigidez, dutilidade e cada parâmetro (e,
# em separado, nas derivadas). Método 2: ensaio pouco linear (R² baixo ou rigidez ponto a ponto variando mais de
# 20%). Cada mola apontada recebe o motivo e a interpretação (erro de dado ou comportamento real). Estes métodos
# só marcam; quem decide a eliminação é a regra dos 2 boxplots, logo abaixo.
rotulos_boxplot = {'rigidez': 'Rigidez K (kgf/mm)', 'dutilidade': 'Dutilidade (fração de C)', **PARAMETROS_PRINCIPAIS}
fora_linearidade = por_mola_df[(por_mola_df['R2'] < LIMITE_R2) | (por_mola_df['variacao_sem_primeiro'] > LIMITE_VARIACAO_PONTOS)]
outliers_df = pd.concat([
    outliers_boxplot(dados_df, rotulos_boxplot, '1. boxplot'),
    outliers_boxplot(derivadas_df, PARAMETROS_DERIVADOS, '1. boxplot (derivadas)'),
    pd.DataFrame({'ID': fora_linearidade['ID'], 'metodo': '2. linearidade do ensaio', 'grandeza': 'ensaio',
                  'motivo': [f"sem o 1º ponto, a rigidez ponto a ponto varia {variacao:.0%}" + ('' if pd.isna(r2) or r2 >= LIMITE_R2 else f", R² {r2:.4f}")
                             for variacao, r2 in zip(fora_linearidade['variacao_sem_primeiro'], fora_linearidade['R2'])]}),
], ignore_index=True)
outliers_df['origem'] = outliers_df['ID'].map(dados_df.set_index('ID')['origem'])
resumo_outliers_df = outliers_df.groupby('ID').agg(origem=('origem', 'first'), metodos=('metodo', lambda metodos: ', '.join(sorted(set(metodos))))).reset_index()
resumo_outliers_df['interpretacao'] = resumo_outliers_df['ID'].map(INTERPRETACAO_OUTLIERS).fillna('Sem interpretação registrada: analisar.')

# Regra de eliminação (a do artigo do professor): a mola é eliminada se ficar fora do limite do boxplot em 2 ou
# mais grandezas de conferência. Uma marcação só não elimina, porque com poucas molas o boxplot marca valores
# extremos verdadeiros. A mola eliminada continua na tabela de dados; sai só das estatísticas e correlações.
fora_conferencia_df = outliers_boxplot(conferencia_df, ROTULOS_CONFERENCIA, 'regra dos 2 boxplots')
contagem_conferencia_df = fora_conferencia_df.groupby('ID')['grandeza'].agg(boxplots='size', grandezas=', '.join).reset_index()
contagem_conferencia_df['eliminada'] = contagem_conferencia_df['boxplots'] >= MINIMO_BOXPLOTS_PARA_ELIMINAR
ids_eliminados = contagem_conferencia_df.loc[contagem_conferencia_df['eliminada'], 'ID'].tolist()

# Comparação com e sem outliers. Cada cenário tira um conjunto de molas e refaz as correlações de Spearman
# (principais e derivadas). Cenários: sem cada mola marcada, uma de cada vez; sem os prováveis erros de dado
# (interpretação que começa com "Provável erro de dado"), tirados só das análises que usam a grandeza com erro; e
# sem as molas eliminadas pela regra dos 2 boxplots, quando houver. Uma conclusão "muda" quando o par troca de
# significativo para não significativo (ou o contrário), ou quando um par significativo troca de sinal. Com
# menos molas o |ρ| exigido sobe, então parte das perdas de significância vem só de ter menos dados.
ids_marcados = resumo_outliers_df['ID'].tolist()
ids_erro_de_dado = [id_mola for id_mola in ids_marcados if INTERPRETACAO_OUTLIERS.get(id_mola, '').startswith('Provável erro de dado')]
# Cada cenário diz, para cada mola tirada, de quais grandezas ela sai (None = sai de tudo).
cenarios = {'todas as molas': {}, **{f'sem a {id_mola}': {id_mola: None} for id_mola in ids_marcados},
            'sem prováveis erros de dado': {id_mola: GRANDEZAS_AFETADAS_POR_ERRO.get(id_mola) for id_mola in ids_erro_de_dado},
            **({'sem as eliminadas pela regra dos 2 boxplots': {id_mola: None for id_mola in ids_eliminados}} if ids_eliminados else {}),
            }


def ids_fora(cenario, colunas_usadas):
    """Molas que o cenário tira de uma análise que usa estas colunas."""
    return [id_mola for id_mola, afetadas in cenario.items() if afetadas is None or set(afetadas) & set(colunas_usadas)]


# Todos os pares (grupo de parâmetros, parâmetro, resultado) que entram na comparação.
pares_analisados = [(tabela, parametro, descricao, resultado, coluna_resultado)
                    for tabela, parametros in ((dados_df, PARAMETROS_PRINCIPAIS), (derivadas_df, PARAMETROS_DERIVADOS))
                    for parametro, descricao in parametros.items()
                    for resultado, coluna_resultado in (('Rigidez K', 'rigidez'), ('Dutilidade', 'dutilidade'))]
partes_correlacao = []
for nome_cenario, cenario in cenarios.items():
    for tabela, parametro, descricao, resultado, coluna_resultado in pares_analisados:
        restantes = tabela[~tabela['ID'].isin(ids_fora(cenario, [parametro, coluna_resultado]))]
        linha = correlacoes_spearman(restantes, {parametro: descricao}).query('resultado == @resultado')
        partes_correlacao.append(linha.assign(cenario=nome_cenario, molas=len(restantes)))
correlacoes_cenarios_df = pd.concat(partes_correlacao, ignore_index=True)
base = correlacoes_cenarios_df[correlacoes_cenarios_df['cenario'] == 'todas as molas'].set_index(['resultado', 'parametro'])
comparacao = correlacoes_cenarios_df.merge(base[['rho', 'significativo']].reset_index(), on=['resultado', 'parametro'], suffixes=('', '_todas'))
comparacao['muda'] = (comparacao['significativo'] != comparacao['significativo_todas']) | ((np.sign(comparacao['rho']) * np.sign(comparacao['rho_todas']) < 0) & comparacao['significativo_todas'])
mudancas_df = comparacao[comparacao['muda']][['cenario', 'resultado', 'parametro', 'molas', 'rho_todas', 'rho', 'p_valor', 'significativo']]
teste_rigidez_dutilidade = stats.spearmanr(dados_df['rigidez'], dados_df['dutilidade'])
correlacao_rigidez_dutilidade = pd.DataFrame({'parametro': ['rigidez'], 'rho': [teste_rigidez_dutilidade.statistic],
                                              'p_valor': [teste_rigidez_dutilidade.pvalue],
                                              'significativo': [teste_rigidez_dutilidade.pvalue < NIVEL_SIGNIFICANCIA]})

#conversão de dataframe
dados_dict = dados_df.to_dict('list')


#processamento de dados
#calculo de correlação entre os parametros
#na matriz de  correlação devemos colocar as variaves que queremos estudar
# A matriz de Pearson inclui todos os parâmetros medidos, a rigidez e a dutilidade (matriz_pearson_r, calculada
# acima pelo .corr() do pandas), e não só h, L e m.
resultado_pearson = matriz_pearson_r


#apresentação
print(dados_df)
# Console enxuto: avisos, outliers, o que muda com e sem eles e o resumo do comportamento geral. As tabelas
# completas (estatísticas, correlações, matrizes, elasticidades, modelos) vão para a planilha de resultados.
print("AVISO: molas sem IC e sem R² por ter poucos pontos:", por_mola_df.loc[por_mola_df['pontos'] < MINIMO_PONTOS_POR_MOLA, 'ID'].tolist())
print("AVISO: C suspeito nas molas", IDS_C_SUSPEITO, "(afeta a dutilidade; valor não corrigido)")
print(f"\nMarcação de outliers (boxplot: {FATOR_IQR}·IQR; linearidade: R² < {LIMITE_R2} ou variação > {LIMITE_VARIACAO_PONTOS:.0%})")
print(outliers_df[['ID', 'origem', 'metodo', 'motivo']].sort_values(['ID', 'metodo']).to_string(index=False))
print("\nResumo por mola, com a interpretação")
for _, mola in resumo_outliers_df.iterrows():
    print(f"  Mola {mola['ID']} ({mola['origem']}; {mola['metodos']}): {mola['interpretacao']}")
print(f"\nRegra de eliminação: fora do limite em {MINIMO_BOXPLOTS_PARA_ELIMINAR} ou mais boxplots de conferência ({', '.join(ROTULOS_CONFERENCIA)})")
print(contagem_conferencia_df.to_string(index=False) if len(contagem_conferencia_df) else "  nenhuma mola fora do limite")
print("Molas eliminadas pela regra:", ids_eliminados if ids_eliminados else "nenhuma (cada mola marcada aparece em um boxplot só)")
print("Valores com provável erro de dado saem só das grandezas afetadas:", GRANDEZAS_AFETADAS_POR_ERRO)
print("\nCom e sem outliers: conclusões de correlação que mudam (significância ou sinal) em relação a todas as molas")
print(mudancas_df.to_string(index=False, float_format='{:.4f}'.format) if len(mudancas_df) else "  nenhuma")
pares_estaveis = base[base['significativo']].index.difference(mudancas_df.set_index(['resultado', 'parametro']).index)
print("Pares significativos que não mudam em nenhum cenário:", ', '.join(f"{resultado} × {parametro}" for resultado, parametro in pares_estaveis))

# Resumo do comportamento geral, em linguagem simples. Todos os números vêm das variáveis calculadas acima,
# então o texto se atualiza sozinho se a planilha mudar. As molas marcadas são agrupadas pelo começo da
# interpretação (comportamento real, provável erro de dado ou efeito de medida).
mais_rigida, menos_rigida = dados_df.loc[dados_df['rigidez'].idxmax()], dados_df.loc[dados_df['rigidez'].idxmin()]
mais_dutil, menos_dutil = dados_df.loc[dados_df['dutilidade'].idxmax()], dados_df.loc[dados_df['dutilidade'].idxmin()]
pares_frageis = mudancas_df[['resultado', 'parametro']].drop_duplicates()
grupos_interpretacao = {rotulo: resumo_outliers_df.loc[resumo_outliers_df['interpretacao'].str.startswith(inicio), 'ID'].tolist()
                        for rotulo, inicio in (('comportamento real', 'Comportamento real'), ('provável erro de dado', 'Provável erro'),
                                               ('efeito de medida', 'Efeito de medida'), ('sem interpretação', 'Sem interpretação'))}
molas_sem_r2 = por_mola_df.loc[por_mola_df['R2'].isna(), 'ID'].tolist()
robustos_dutilidade = [parametro for resultado, parametro in pares_estaveis if resultado == 'Dutilidade']
alongamento_falha = dados_df['dutilidade'] * dados_df['C'] #alongamento na falha em mm, sem dividir por C
alongamento_com_h = stats.spearmanr(dados_df['h'], alongamento_falha)
alongamento_com_fio = stats.spearmanr(derivadas_df['fio_desenvolvido'], alongamento_falha)
medianas_por_espiras = dados_df.groupby('n').agg(molas=('ID', 'size'), rigidez_mediana=('rigidez', 'median'), dutilidade_mediana=('dutilidade', 'median'))
pearson_por_par = pearson_df.set_index(['resultado', 'parametro'])['r']
maiores_medidos = {resultado: grupo.loc[grupo['r'].abs().idxmax()] for resultado, grupo in pearson_df[pearson_df['grupo'] == 'principais'].groupby('resultado')}
fortes_derivadas = pearson_df[(pearson_df['grupo'] == 'derivadas') & pearson_df['forte']]
print("\n" + "=" * 100)
print("COMPORTAMENTO GERAL DO SISTEMA")
print("=" * 100)
print(f"1. Sob carga, as molas se comportam como molas lineares até a falha: R² entre {por_mola_df['R2'].min():.4f} e {por_mola_df['R2'].max():.4f}"
      + (f" (a mola {', '.join(map(str, molas_sem_r2))} tem só 2 pontos)." if molas_sem_r2 else "."))
print(f"2. A rigidez vai de {menos_rigida['rigidez']:.4f} (mola {menos_rigida['ID']}) a {mais_rigida['rigidez']:.4f} kgf/mm (mola {mais_rigida['ID']}), "
      f"{mais_rigida['rigidez'] / menos_rigida['rigidez']:.0f} vezes; a dutilidade vai de {menos_dutil['dutilidade']:.0%} (mola {menos_dutil['ID']}) "
      f"a {mais_dutil['dutilidade']:.0%} (mola {mais_dutil['ID']}). Como os valores são assimétricos, a mediana representa melhor: "
      f"{dados_df['rigidez'].median():.4f} kgf/mm e {dados_df['dutilidade'].median():.0%}.")
print(f"3. A rigidez segue a teoria de viga: o índice de rigidez t·h³ / (n·Ltot³) tem r = {pearson_por_par[('Rigidez K', 'indice_rigidez')]:+.2f} com K, "
      f"e a reta K = {inclinacao_simples:.1f} × índice {intercepto_simples:+.4f} explica {r2_simples:.1%} da variação.")
print("4. A dutilidade (alongamento na falha ÷ C) cai quando h aumenta e sobe com a esbeltez (Ltot/h). Cuidado: o alongamento em mm")
print(f"   quase não depende de h (ρ = {alongamento_com_h.statistic:+.2f}); ele cresce com o comprimento de fio que trabalha (fio desenvolvido, "
      f"ρ = {alongamento_com_fio.statistic:+.2f}). O efeito de h na dutilidade vem em boa parte de C, que está no denominador e cresce com h.")
print("   Relações com a dutilidade que valem em todos os cenários com e sem outliers: " + (', '.join(robustos_dutilidade) or 'nenhuma') + ".")
print(f"5. Rigidez e dutilidade andam em sentidos opostos (ρ = {teste_rigidez_dutilidade.statistic:+.2f}): a geometria que deixa a mola rígida")
print("   é a mesma que limita o alongamento, e a dutilidade ainda divide por C. Não é o material ficando frágil.")
print(f"6. Nenhum parâmetro medido sozinho chega a |r| = {LIMITE_PEARSON:.2f} (maiores: " + "; ".join(
      f"{resultado}: {linha['parametro']} {linha['r']:+.2f}" for resultado, linha in maiores_medidos.items())
      + "). As relações fortes aparecem nas dimensões derivadas: " + "; ".join(
      f"{linha.parametro} × {linha.resultado} {linha.r:+.2f}" for linha in fortes_derivadas.itertuples()) + ".")
print(f"7. Medida de h: nas molas do professor o h medido fica em média {desvio_h_por_origem.loc['professor', 'mean']:+.2f} mm em relação ao nominal; "
      f"nas do grupo, {desvio_h_por_origem.loc['grupo', 'mean']:+.2f} mm (de {desvio_h_por_origem.loc['grupo', 'min']:+.2f} a {desvio_h_por_origem.loc['grupo', 'max']:+.2f}).")
print(f"   Com o h nominal, o R² da reta K × índice de rigidez vai de {r2_simples:.3f} para {r2_com_h_nominal:.3f}: as molas do grupo se comportam como se")
print("   tivessem o h nominal, o que indica h medido acima do que trabalha. As molas foram descartadas após o ensaio, então não foi possível remedir.")
print(f"8. Outliers: {len(ids_marcados)} molas marcadas; eliminadas pela regra dos 2 boxplots: {', '.join(map(str, ids_eliminados)) or 'nenhuma'}. " + "; ".join(
      f"{rotulo}: {', '.join(map(str, ids))}" for rotulo, ids in grupos_interpretacao.items() if ids) + ".")
print("   Conclusões que dependem de uma mola ou dos prováveis erros de dado: "
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
print(resultado_pearson.round(2))
# Resumo do método do professor: diferença entre as duas fórmulas de rigidez, relações fortes pelo Pearson e a
# regressão linear simples.
print(f"\nRigidez principal: {RIGIDEZ_PRINCIPAL}. Diferença entre a média ponto a ponto e a reta pela origem: até "
      f"{por_mola_df['diferenca_formulas'].abs().max():.1%} (mola {por_mola_df.loc[por_mola_df['diferenca_formulas'].abs().idxmax(), 'ID']}).")
print(f"Pearson, relações fortes (|r| ≥ {LIMITE_PEARSON}):")
print(pearson_df[pearson_df['forte']][['grupo', 'resultado', 'parametro', 'r', 'p_valor']].to_string(index=False, float_format='{:.3f}'.format))
print("Parâmetros medidos, maior |r| com cada resultado: " + "; ".join(
    f"{resultado}: {linha['parametro']} ({linha['r']:+.2f})" for resultado, linha in
    ((resultado, grupo.loc[grupo['r'].abs().idxmax()]) for resultado, grupo in pearson_df[pearson_df['grupo'] == 'principais'].groupby('resultado'))))
print(f"Regressão linear simples: K = {inclinacao_simples:.1f} × índice de rigidez {intercepto_simples:+.4f}, R² = {r2_simples:.3f}")

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
# Histogramas de cada grandeza (medidas, resultados e derivadas), exigidos na proposta do projeto.
grafico_histogramas(dados_df, rotulos_boxplot, "Histogramas: rigidez, dutilidade e parâmetros medidos", "histogramas_principais.png")
grafico_histogramas(derivadas_df, PARAMETROS_DERIVADOS, "Histogramas: dimensões derivadas", "histogramas_derivadas.png")
# Matriz de Pearson (com * onde |r| ≥ 0,70), a regressão linear simples e os boxplots da regra de eliminação.
grafico_matriz(matriz_pearson_r, matriz_pearson_p, [*PARAMETROS_PRINCIPAIS.values(), 'Rigidez K (kgf/mm)', 'Dutilidade (fração de C)'],
               "Correlação de Pearson entre parâmetros, rigidez e dutilidade", "matriz_pearson.png",
               rotulo_escala='r de Pearson', nota=f"* = relação forte, |r| ≥ {LIMITE_PEARSON:.2f}".replace('.', ','), destaque=matriz_pearson_r.abs() >= LIMITE_PEARSON)
grafico_regressao_simples(derivadas_df, inclinacao_simples, intercepto_simples, r2_simples, "regressao_rigidez_indice.png")
grafico_boxplots(conferencia_df, ROTULOS_CONFERENCIA, fora_conferencia_df, "Grandezas de conferência da regra de eliminação", "boxplots_conferencia.png")
# Gráficos da relação entre parâmetros: as duas matrizes de correlação e a dispersão rigidez × dutilidade.
grafico_matriz(matriz_principais_rho, matriz_principais_p, list(PARAMETROS_PRINCIPAIS.values()), "Correlação entre os parâmetros principais", "matriz_parametros_principais.png")
grafico_matriz(matriz_derivadas_rho, matriz_derivadas_p, list(PARAMETROS_DERIVADOS.values()), "Correlação entre as dimensões derivadas", "matriz_dimensoes_derivadas.png")
grafico_dispersoes(dados_df, {'rigidez': 'Rigidez K (kgf/mm)'}, 'dutilidade', 'Dutilidade (fração de C)', correlacao_rigidez_dutilidade, "rigidez_x_dutilidade.png")
# Gráficos dos outliers: boxplots (parâmetros principais e derivadas) e as curvas carga × alongamento de cada
# ensaio. Depois, o desvio do h medido em relação ao nominal.
grafico_boxplots(dados_df, rotulos_boxplot, outliers_df[outliers_df['metodo'] == '1. boxplot'], "Rigidez, dutilidade e parâmetros principais", "boxplots_principais.png")
grafico_boxplots(derivadas_df, PARAMETROS_DERIVADOS, outliers_df[outliers_df['metodo'] == '1. boxplot (derivadas)'], "Dimensões derivadas", "boxplots_derivadas.png")
grafico_curvas_ensaio(ensaios_df, por_mola_df, "curvas_carga_alongamento.png")
grafico_desvio_h(medidas_h_df, "desvio_h_nominal.png")
# Gráfico da comparação com e sem as molas marcadas (parâmetros principais).
grafico_com_sem_outliers(correlacoes_cenarios_df, PARAMETROS_PRINCIPAIS, rho_minimo, "Correlações com e sem outliers: parâmetros principais", "com_sem_outliers_principais.png")
# Gráfico final: rigidez e dutilidade por número de espiras.
grafico_por_espiras(dados_df, "por_numero_de_espiras.png")

# Planilha de resultados em saidas/ (a planilha de entrada não é alterada). Uma aba por tabela calculada:
# rigidez e dutilidade de cada mola com as dimensões, derivadas, estatísticas, correlações, matrizes,
# pares significativos, Pearson, medidas de h, outliers, comparação com e sem outliers e medianas por n.
arquivo_resultados = os.path.join(PASTA_SAIDAS, "experimento_resultados.xlsx")
with pd.ExcelWriter(arquivo_resultados) as planilha_resultados:
    dados_df.drop(columns=['rigidez', 'dutilidade']).merge(por_mola_df, on=['ID', 'origem']).to_excel(planilha_resultados, sheet_name='resultados_por_mola', index=False)
    derivadas_df.to_excel(planilha_resultados, sheet_name='dimensoes_derivadas', index=False)
    pd.concat([estatisticas_df, estatisticas_derivadas_df]).to_excel(planilha_resultados, sheet_name='estatisticas', index=False)
    pd.concat([correlacoes_principais_df, correlacoes_derivadas_df]).to_excel(planilha_resultados, sheet_name='correlacoes', index=False)
    matriz_principais_rho.to_excel(planilha_resultados, sheet_name='matriz_principais')
    matriz_derivadas_rho.to_excel(planilha_resultados, sheet_name='matriz_derivadas')
    medidas_h_df.to_excel(planilha_resultados, sheet_name='medidas_h', index=False)
    pearson_df.to_excel(planilha_resultados, sheet_name='pearson', index=False)
    matriz_pearson_r.to_excel(planilha_resultados, sheet_name='matriz_pearson')
    pd.DataFrame({'inclinacao': [inclinacao_simples], 'intercepto': [intercepto_simples], 'R2': [r2_simples]}).to_excel(planilha_resultados, sheet_name='regressao_simples', index=False)
    conferencia_df.to_excel(planilha_resultados, sheet_name='conferencia', index=False)
    contagem_conferencia_df.to_excel(planilha_resultados, sheet_name='regra_2_boxplots', index=False)
    outliers_df.to_excel(planilha_resultados, sheet_name='outliers', index=False)
    resumo_outliers_df.to_excel(planilha_resultados, sheet_name='outliers_resumo', index=False)
    mudancas_df.to_excel(planilha_resultados, sheet_name='com_sem_outliers', index=False)
    pd.concat({'principais': pares_significativos(matriz_principais_rho, matriz_principais_p),
               'derivadas': pares_significativos(matriz_derivadas_rho, matriz_derivadas_p)}).to_excel(planilha_resultados, sheet_name='pares_significativos')
    medianas_por_espiras.to_excel(planilha_resultados, sheet_name='por_numero_de_espiras')
print(f"\nPlanilha de resultados gravada em {arquivo_resultados}")
