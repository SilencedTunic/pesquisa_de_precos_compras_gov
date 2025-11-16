"""
Coletor automático de itens de contratações PNCP (Lei 14.133/2021)
Versão 3.3 – Excel + Relatório HTML (sem seção de gráficos).

Como usar no Jupyter:
1. Vá até a seção de CONFIGURAÇÕES BÁSICAS logo abaixo.
2. Preencha os filtros que você quiser (ou deixe como None / "" para ignorar).
   - Inclusive o COD_ITEM_CATALOGO é opcional.
3. Rode a célula inteira.
4. Ao final, serão gerados:
   - Um arquivo .xlsx com as abas:
       • 'dados'            → registros completos
       • 'resumo_unidade'   → estatísticas por unidadeMedida
           (resultado + média saneada + limites)
       • 'preco_referencia' → média, mediana e média saneada por unidade de medida
   - Um arquivo .html contendo uma nota técnica explicativa da pesquisa de preços,
     com introdução, filtros, estatísticas, metodologia, resultados
     e um quadro-resumo de preço de referência por unidade de medida.

A janela temporal é sempre: hoje até 1 ano atrás (365 dias).
"""

# ============================================================
# 🔧 CONFIGURAÇÕES BÁSICAS (EDITE AQUI)
# ============================================================

# Informe aqui o código do item de catálogo (CATMAT/CATSER).
# Deixe como None se não quiser filtrar por codItemCatalogo.
COD_ITEM_CATALOGO = None  # ex.: 279727 ou None

# Filtros opcionais (defina os valores desejados ou deixe como None/"" para ignorar)
ORGAO_ENTIDADE_CNPJ = ""                 # string ou "" para ignorar
UNIDADE_ORGAO_CODIGO_UNIDADE = None      # int ou None
SITUACAO_COMPRA_ITEM = ""                # string (ex.: "4") ou "" para ignorar

# MATERIAL_OU_SERVICO:
#   "M"  → Material
#   "S"  → Serviço
#   None ou "" → não envia o parâmetro (pega tudo)
MATERIAL_OU_SERVICO = ""                 # "M", "S" ou None/""

CODIGO_CLASSE = None                     # int ou None (permite consulta só por classe)
CODIGO_GRUPO = None                      # int ou None
COD_FORNECEDOR = ""                      # string ou "" para ignorar
FILTRAR_TEM_RESULTADO = None             # True, False ou None
FILTRAR_BPS = None                       # True, False ou None
FILTRAR_MARGEM_PREFERENCIA_NORMAL = None # True, False ou None
CODIGO_NCM = ""                          # string ou "" para ignorar

# Opcional: nome base dos arquivos de saída (sem extensão).
# Se deixar None, será gerado automaticamente.
NOME_BASE_SAIDA = None  # ex.: "pesquisa_preco_catmat_279727"


# ============================================================
# 📦 IMPORTAÇÕES
# ============================================================

try:
    import requests
except ImportError as exc:
    print("❌ Erro: a biblioteca 'requests' não está instalada.")
    print("   Instale com: pip install requests")
    raise exc

try:
    import pandas as pd
except ImportError as exc:
    print("❌ Erro: a biblioteca 'pandas' não está instalada.")
    print("   Instale com: pip install pandas")
    raise exc

try:
    import matplotlib.pyplot as plt
except ImportError as exc:
    print("❌ Erro: a biblioteca 'matplotlib' não está instalada.")
    print("   Instale com: pip install matplotlib")
    raise exc

try:
    import openpyxl  # garante engine do Excel
except ImportError as exc:
    print("❌ Erro: a biblioteca 'openpyxl' não está instalada.")
    print("   Instale com: pip install openpyxl")
    raise exc

import base64
from io import BytesIO
from datetime import date, timedelta
import numpy as np


# ============================================================
# 🗓️ INTERVALO DE 1 ANO
# ============================================================

def calcular_intervalo_ultimo_ano():
    """
    Retorna (data_inicial, data_final) em formato 'YYYY-MM-DD',
    considerando 'hoje' e 'hoje - 365 dias'.
    """
    data_final = date.today()
    data_inicial = data_final - timedelta(days=365)
    return data_inicial.strftime("%Y-%m-%d"), data_final.strftime("%Y-%m-%d")


# ============================================================
# 🔄 AJUDANTES PARA FILTROS OPCIONAIS
# ============================================================

def bool_to_api_flag(value):
    """
    Converte True/False em 'true'/'false' para a API.
    Retorna None se value não for booleano.
    """
    if isinstance(value, bool):
        return "true" if value else "false"
    return None


def montar_filtros_opcionais():
    """
    Lê as variáveis de configuração no topo e monta o dicionário
    de parâmetros opcionais a ser enviado para a API.
    Só inclui parâmetros que não forem None/vazios.
    """
    filtros = {}

    if ORGAO_ENTIDADE_CNPJ:
        filtros["orgaoEntidadeCnpj"] = ORGAO_ENTIDADE_CNPJ

    if UNIDADE_ORGAO_CODIGO_UNIDADE is not None:
        filtros["unidadeOrgaoCodigoUnidade"] = int(UNIDADE_ORGAO_CODIGO_UNIDADE)

    if SITUACAO_COMPRA_ITEM:
        filtros["situacaoCompraItem"] = SITUACAO_COMPRA_ITEM

    if MATERIAL_OU_SERVICO:
        filtros["materialOuServico"] = MATERIAL_OU_SERVICO

    if CODIGO_CLASSE is not None:
        filtros["codigoClasse"] = int(CODIGO_CLASSE)

    if CODIGO_GRUPO is not None:
        filtros["codigoGrupo"] = int(CODIGO_GRUPO)

    if COD_FORNECEDOR:
        filtros["codFornecedor"] = COD_FORNECEDOR

    flag_tr = bool_to_api_flag(FILTRAR_TEM_RESULTADO)
    if flag_tr is not None:
        filtros["temResultado"] = flag_tr

    flag_bps = bool_to_api_flag(FILTRAR_BPS)
    if flag_bps is not None:
        filtros["bps"] = flag_bps

    flag_mpn = bool_to_api_flag(FILTRAR_MARGEM_PREFERENCIA_NORMAL)
    if flag_mpn is not None:
        filtros["margemPreferenciaNormal"] = flag_mpn

    if CODIGO_NCM:
        filtros["codigoNCM"] = CODIGO_NCM

    return filtros


# ============================================================
# 🌐 CHAMADA PAGINADA À API
# ============================================================

def buscar_itens_pncp(cod_item_catalogo, data_inicial, data_final,
                      filtros_opcionais=None, tamanho_pagina=500):
    """
    Faz chamadas paginadas ao endpoint:
      /modulo-contratacoes/2_consultarItensContratacoes_PNCP_14133

    Parâmetros mínimos:
      - dataInclusaoPncpInicial
      - dataInclusaoPncpFinal

    'cod_item_catalogo' é opcional (pode ser None).
    'filtros_opcionais' pode conter qualquer outro parâmetro aceito pela API.

    Retorna:
      - Lista de dicionários (cada dicionário é um item retornado pela API).
    """
    base_url = (
        "https://dadosabertos.compras.gov.br/"
        "modulo-contratacoes/2_consultarItensContratacoes_PNCP_14133"
    )

    pagina = 1
    todos_resultados = []
    filtros_opcionais = filtros_opcionais or {}

    print("==============================================")
    print(" Iniciando coleta na API Compras.gov.br (v3.3)")
    print(" Intervalo de inclusão PNCP:", data_inicial, "até", data_final)
    if cod_item_catalogo is not None:
        print(" codItemCatalogo:", cod_item_catalogo)
    else:
        print(" codItemCatalogo: não informado (consulta sem filtro de item).")
    print(" Filtros opcionais:",
          filtros_opcionais if filtros_opcionais else "nenhum")
    print("==============================================")

    while True:
        # Parâmetros obrigatórios
        params = {
            "pagina": pagina,
            "tamanhoPagina": tamanho_pagina,
            "dataInclusaoPncpInicial": data_inicial,
            "dataInclusaoPncpFinal": data_final,
        }

        # Parâmetro opcional codItemCatalogo
        if cod_item_catalogo is not None:
            params["codItemCatalogo"] = cod_item_catalogo

        # Demais filtros opcionais
        for k, v in filtros_opcionais.items():
            params[k] = v

        print(f"▶ Buscando página {pagina}...")
        try:
            resp = requests.get(base_url, params=params, timeout=60)
        except Exception as exc:
            print("❌ Erro de conexão ao chamar a API.")
            print("   Detalhes:", exc)
            break

        if resp.status_code != 200:
            print(f"❌ Erro HTTP {resp.status_code} na página {pagina}.")
            print("   Trecho da resposta:", resp.text[:500])
            break

        try:
            dados = resp.json()
        except ValueError:
            print("❌ Erro ao interpretar a resposta como JSON.")
            print("   Conteúdo recebido (início):")
            print(resp.text[:500])
            break

        resultados_pagina = dados.get("resultado", [])

        if not resultados_pagina:
            print("⚠ Nenhum registro nesta página. Encerrando paginação.")
            break

        todos_resultados.extend(resultados_pagina)

        total_paginas = dados.get("totalPaginas")
        paginas_restantes = dados.get("paginasRestantes")

        print(
            f"   → Página {pagina} retornou {len(resultados_pagina)} registros. "
            f"Total acumulado: {len(todos_resultados)}"
        )

        # Critérios de parada
        if paginas_restantes in (0, None):
            print("✅ Paginação concluída (sem páginas restantes).")
            break

        if total_paginas is not None and pagina >= total_paginas:
            print("✅ Paginação concluída (atingido totalPaginas informado).")
            break

        pagina += 1

    print("----------------------------------------------")
    print(f" Coleta finalizada com {len(todos_resultados)} registros.")
    print("----------------------------------------------")

    return todos_resultados


# ============================================================
# 📊 MÉDIA SANEADA, RESUMO E PREÇO DE REFERÊNCIA
# ============================================================

def calcular_media_sanada_serie(serie: pd.Series, cv_limite: float = 25.0) -> float:
    """
    Calcula a média saneada de uma série numérica.
    (expurgo iterativo por desvio-padrão até CV <= limite, ou devolve média simples)
    """
    s = pd.to_numeric(serie.dropna(), errors="coerce").dropna()
    if s.empty:
        return float("nan")

    while True:
        m = s.mean()
        dp = s.std(ddof=0)
        if m == 0 or pd.isna(m) or pd.isna(dp) or len(s) < 3:
            return m

        cv = abs(dp / m) * 100.0
        if cv <= cv_limite:
            return m

        li = m - dp
        ls = m + dp
        filtrado = s[(s >= li) & (s <= ls)]

        if len(filtrado) == len(s) or filtrado.empty:
            return m

        s = filtrado


def calcular_resumo_por_unidade(df: pd.DataFrame) -> pd.DataFrame:
    """
    Considera apenas 'valorUnitarioResultado' para o resumo estatístico;
    inclui:
      - media_sanada
      - limite_inferior_intervalo
      - limite_superior_intervalo
    """
    if df.empty or "unidadeMedida" not in df.columns:
        return pd.DataFrame()

    df_local = df.copy()

    if "valorUnitarioResultado" not in df_local.columns:
        return pd.DataFrame()

    df_local["valorUnitarioResultado"] = pd.to_numeric(
        df_local["valorUnitarioResultado"], errors="coerce"
    )

    grp = df_local.groupby("unidadeMedida")["valorUnitarioResultado"]

    resumo_base = (
        grp.agg(["count", "mean", "median", "std", "min", "max"])
        .rename(
            columns={
                "count": "resultado_qtde",
                "mean": "resultado_media",
                "median": "resultado_mediana",
                "std": "resultado_desvio_padrao",
                "min": "resultado_minimo",
                "max": "resultado_maximo",
            }
        )
    )

    media_sanada = grp.apply(calcular_media_sanada_serie).rename("media_sanada")

    resumo = resumo_base.join(media_sanada, how="left")

    for col in ["resultado_desvio_padrao", "media_sanada",
                "resultado_media", "resultado_mediana"]:
        if col in resumo.columns:
            resumo[col] = pd.to_numeric(resumo[col], errors="coerce")

    base = resumo["media_sanada"].copy()
    mask_nan = base.isna()
    if "resultado_media" in resumo.columns:
        base[mask_nan] = resumo.loc[mask_nan, "resultado_media"]
        mask_nan = base.isna()
    if "resultado_mediana" in resumo.columns:
        base[mask_nan] = resumo.loc[mask_nan, "resultado_mediana"]

    dp = resumo["resultado_desvio_padrao"].fillna(0)
    resumo["limite_inferior_intervalo"] = (base - dp).clip(lower=0)
    resumo["limite_superior_intervalo"] = (base + dp).clip(lower=0)

    resumo = resumo.reset_index().sort_values("unidadeMedida")
    return resumo


def montar_preco_referencia(resumo_df: pd.DataFrame) -> pd.DataFrame:
    """
    Monta aba 'preco_referencia' com:
      unidadeMedida, media, mediana, media_sanada
    """
    if resumo_df is None or resumo_df.empty:
        return pd.DataFrame()

    df = resumo_df.copy()

    for col in ["resultado_media", "resultado_mediana", "media_sanada"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    colunas_necessarias = [
        "unidadeMedida",
        "resultado_media",
        "resultado_mediana",
        "media_sanada",
    ]
    colunas_existentes = [c for c in colunas_necessarias if c in df.columns]
    df_out = df[colunas_existentes].copy()

    renomear = {}
    if "resultado_media" in df_out.columns:
        renomear["resultado_media"] = "media"
    if "resultado_mediana" in df_out.columns:
        renomear["resultado_mediana"] = "mediana"

    df_out = df_out.rename(columns=renomear)
    df_out = df_out.sort_values("unidadeMedida")

    return df_out


# ============================================================
# 💾 PREPARAR DATAFRAMES + SALVAR EM EXCEL
# ============================================================

def preparar_dataframes(dados: list) -> tuple:
    """
    A partir da lista de dicionários retornada pela API,
    monta:
      - df_dados         → DataFrame completo
      - resumo_df        → resumo por unidadeMedida
      - preco_ref_df     → tabela de preço de referência (resumida)
    """
    df = pd.DataFrame(dados)
    if df.empty:
        return df, pd.DataFrame(), pd.DataFrame()

    colunas_prioritarias = [
        "idContratacaoPNCP",
        "idCompra",
        "idCompraItem",
        "orgaoEntidadeCnpj",
        "unidadeOrgaoCodigoUnidade",
        "descricaoResumida",
        "descricaodetalhada",
        "materialOuServicoNome",
        "codigoClasse",
        "codigoGrupo",
        "codItemCatalogo",
        "unidadeMedida",
        "quantidade",
        "valorUnitarioEstimado",
        "valorTotal",
        "quantidadeResultado",
        "valorUnitarioResultado",
        "valorTotalResultado",
        "situacaoCompraItemNome",
        "nomeFornecedor",
        "dataInclusaoPncp",
        "dataAtualizacaoPncp",
        "dataResultado",
        "codigoNCM",
        "descricaoNCM",
    ]
    colunas_existentes = [c for c in colunas_prioritarias if c in df.columns]
    outras_colunas = [c for c in df.columns if c not in colunas_existentes]
    df = df[colunas_existentes + outras_colunas]

    resumo_df = calcular_resumo_por_unidade(df)
    preco_ref_df = montar_preco_referencia(resumo_df) if not resumo_df.empty else pd.DataFrame()

    return df, resumo_df, preco_ref_df


def salvar_resultados_em_excel(df_dados, resumo_df, preco_ref_df, caminho_arquivo):
    """
    Salva em Excel:
      - Aba 'dados'            → registros detalhados
      - Aba 'resumo_unidade'   → estatísticas por unidadeMedida
      - Aba 'preco_referencia' → média, mediana e média saneada
    """
    if df_dados is None or df_dados.empty:
        print("⚠ Nenhum dado para salvar em Excel.")
        return

    print(f"💾 Salvando arquivo Excel em: {caminho_arquivo}")
    with pd.ExcelWriter(caminho_arquivo, engine="openpyxl") as writer:
        df_dados.to_excel(writer, index=False, sheet_name="dados")
        if resumo_df is not None and not resumo_df.empty:
            resumo_df.to_excel(writer, index=False, sheet_name="resumo_unidade")
        if preco_ref_df is not None and not preco_ref_df.empty:
            preco_ref_df.to_excel(writer, index=False, sheet_name="preco_referencia")

    print("✅ Arquivo Excel gerado com sucesso.")


# ============================================================
# 📝 RELATÓRIO HTML (NOTA TÉCNICA, SEM SEÇÃO DE GRÁFICOS)
# ============================================================

def _fig_to_base64img(fig):
    """
    Helper mantido apenas por compatibilidade (não é usado nesta versão).
    """
    buf = BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight")
    plt.close(fig)
    buf.seek(0)
    b64 = base64.b64encode(buf.read()).decode("utf-8")
    return f"data:image/png;base64,{b64}"


def gerar_relatorio_html(df_dados: pd.DataFrame,
                         resumo_df: pd.DataFrame,
                         preco_ref_df: pd.DataFrame,
                         meta: dict,
                         caminho_html: str):
    """
    Gera relatório HTML em formato de nota técnica, SEM a seção 5 (gráficos).
    Mantém:
      1. Introdução
      2. Período e filtros
      3. Estatísticas descritivas
      4. Metodologia
      6. Resultados e uso recomendado
      7. Quadro-resumo de preço de referência
    """
    print(f"📝 Gerando relatório HTML em: {caminho_html}")

    total_registros = len(df_dados)
    if "unidadeMedida" in df_dados.columns:
        unidades_distintas = int(df_dados["unidadeMedida"].nunique())
    else:
        unidades_distintas = 0

    # Estatísticas de valorUnitarioResultado
    estat_resultado = {}
    if "valorUnitarioResultado" in df_dados.columns:
        serie = pd.to_numeric(df_dados["valorUnitarioResultado"], errors="coerce").dropna()
        if not serie.empty:
            estat_resultado = {
                "min": float(serie.min()),
                "max": float(serie.max()),
                "mean": float(serie.mean()),
                "median": float(serie.median()),
                "std": float(serie.std(ddof=0)),
            }

    # Tabela de filtros
    filtros_html_rows = ""
    for chave, valor in meta.get("filtros_efetivos", {}).items():
        filtros_html_rows += f"<tr><td>{chave}</td><td>{valor}</td></tr>\n"

    # Estatísticas globais
    estat_html_rows = ""
    for k, v in estat_resultado.items():
        estat_html_rows += f"<tr><td>{k}</td><td>{v:.4f}</td></tr>\n"

    hoje_str = date.today().strftime("%d/%m/%Y")

    # Quadro-resumo de preço de referência
    quadro_html_rows = ""
    if preco_ref_df is not None and not preco_ref_df.empty:
        quadro_df = preco_ref_df.copy()
        for col in ["media", "mediana", "media_sanada"]:
            if col in quadro_df.columns:
                quadro_df[col] = pd.to_numeric(quadro_df[col], errors="coerce")

        quadro_df["preco_referencia"] = quadro_df.get("media_sanada")
        if "mediana" in quadro_df.columns:
            mask_nan = quadro_df["preco_referencia"].isna()
            quadro_df.loc[mask_nan, "preco_referencia"] = quadro_df.loc[mask_nan, "mediana"]
        if "media" in quadro_df.columns:
            mask_nan = quadro_df["preco_referencia"].isna()
            quadro_df.loc[mask_nan, "preco_referencia"] = quadro_df.loc[mask_nan, "media"]

        if resumo_df is not None and not resumo_df.empty:
            limites = resumo_df[[
                "unidadeMedida",
                "limite_inferior_intervalo",
                "limite_superior_intervalo"
            ]].copy()
            quadro_df = quadro_df.merge(limites, on="unidadeMedida", how="left")

        quadro_df = quadro_df.sort_values("unidadeMedida")

        def fmt(x):
            try:
                return f"{float(x):.4f}"
            except Exception:
                return ""

        for _, row in quadro_df.iterrows():
            um = row.get("unidadeMedida", "")
            media = row.get("media", float("nan"))
            mediana = row.get("mediana", float("nan"))
            media_sanada = row.get("media_sanada", float("nan"))
            pr = row.get("preco_referencia", float("nan"))
            li = row.get("limite_inferior_intervalo", float("nan"))
            ls = row.get("limite_superior_intervalo", float("nan"))

            quadro_html_rows += (
                "<tr>"
                f"<td>{um}</td>"
                f"<td>{fmt(media)}</td>"
                f"<td>{fmt(mediana)}</td>"
                f"<td>{fmt(media_sanada)}</td>"
                f"<td>{fmt(pr)}</td>"
                f"<td>{fmt(li)}</td>"
                f"<td>{fmt(ls)}</td>"
                "</tr>\n"
            )

    # HTML
    html = f"""
<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<title>Relatório de Pesquisa de Preços – PNCP</title>
<style>
body {{ font-family: Arial, sans-serif; margin: 20px; }}
h1, h2, h3 {{ color: #333; }}
table {{ border-collapse: collapse; width: 100%; margin-bottom: 20px; }}
th, td {{ border: 1px solid #ccc; padding: 8px; text-align: left; }}
th {{ background-color: #f0f0f0; }}
.section {{ margin-bottom: 30px; }}
small {{ color: #555; }}
</style>
</head>
<body>

<h1>Relatório de Pesquisa de Preços – PNCP (Lei 14.133/2021)</h1>
<p><small>Relatório gerado em {hoje_str}</small></p>

<div class="section">
<h2>1. Introdução</h2>
<p>
Este relatório apresenta os resultados de uma pesquisa de preços realizada a partir de dados
extraídos do Portal Nacional de Contratações Públicas (PNCP), utilizando o serviço de dados
abertos do Compras.gov.br. O objetivo é subsidiar a estimativa de preços para contratações
públicas, de forma transparente, reprodutível e alinhada às boas práticas de planejamento das
contratações previstas na Lei nº 14.133/2021.
</p>
</div>

<div class="section">
<h2>2. Período e filtros utilizados</h2>
<p>Período de inclusão no PNCP considerado na amostra:</p>
<ul>
  <li><strong>Data inicial:</strong> {meta.get("data_inicial", "")}</li>
  <li><strong>Data final:</strong> {meta.get("data_final", "")}</li>
</ul>

<p>Resumo dos filtros aplicados na consulta:</p>
<table>
  <thead>
    <tr><th>Parâmetro</th><th>Valor</th></tr>
  </thead>
  <tbody>
    {filtros_html_rows}
  </tbody>
</table>
</div>

<div class="section">
<h2>3. Estatísticas descritivas da amostra</h2>
<p>
A amostra utilizada contém <strong>{total_registros}</strong> registros de itens de contratação
e <strong>{unidades_distintas}</strong> unidade(s) de medida distinta(s).
</p>
"""

    if estat_resultado:
        html += f"""
<p>Para o campo <code>valorUnitarioResultado</code>, as estatísticas descritivas globais são:</p>
<table>
  <thead>
    <tr><th>Medida</th><th>Valor</th></tr>
  </thead>
  <tbody>
    {estat_html_rows}
  </tbody>
</table>
"""
    else:
        html += "<p>Não foi possível calcular estatísticas descritivas para <code>valorUnitarioResultado</code>.</p>"

    html += """
</div>

<div class="section">
<h2>4. Metodologia de cálculo</h2>
<p>
Os dados foram extraídos diretamente da API oficial do PNCP, considerando o período informado
e os filtros aplicados. Após a consolidação dos registros em base única, procedeu-se ao cálculo
de estatísticas descritivas por unidade de medida, com destaque para a <strong>média saneada</strong>,
obtida a partir da seguinte lógica:
</p>
<ol>
  <li>Para cada unidade de medida, são considerados os valores de <code>valorUnitarioResultado</code> válidos.</li>
  <li>Calculam-se a média (M) e o desvio-padrão (DP) da amostra.</li>
  <li>É obtido o coeficiente de variação (CV = DP / M * 100). Se o CV for menor ou igual ao limite pré-definido (25%), a média simples é adotada como média saneada.</li>
  <li>Caso o CV seja superior ao limite, são expurgados os valores considerados outliers, isto é, aqueles abaixo de M - DP ou acima de M + DP.</li>
  <li>O procedimento é repetido iterativamente enquanto houver exclusão de valores e o CV permanecer acima do limite.</li>
  <li>Ao final do processo, a média calculada sobre o conjunto remanescente é definida como <strong>média saneada</strong>.</li>
</ol>
<p>
A partir da média saneada e do desvio-padrão por unidade de medida, foram também construídos
intervalos de referência (limite inferior e superior), utilizados como apoio à análise crítica
dos valores de mercado.
</p>
</div>

<div class="section">
<h2>6. Resultados e uso recomendado</h2>
<p>
Os resultados consolidados encontram-se detalhados nas planilhas eletrônicas geradas em paralelo
a este relatório, contendo:
</p>
<ul>
  <li>Aba <strong>dados</strong>: base completa de registros extraídos da API.</li>
  <li>Aba <strong>resumo_unidade</strong>: estatísticas descritivas por unidade de medida, com destaque para a média saneada e para os intervalos de referência.</li>
  <li>Aba <strong>preco_referencia</strong>: visão resumida das medidas centrais (média, mediana e média saneada) por unidade de medida.</li>
</ul>
<p>
Recomenda-se que o <strong>preço de referência</strong> para fins de estimativa seja definido a partir
da análise conjunta da média saneada, da mediana e do contexto de mercado, podendo ser adotada,
por exemplo, a própria média saneada como valor de referência, desde que tecnicamente justificada
no estudo preliminar e na nota técnica de pesquisa de preços.
</p>
<p>
Este relatório pode ser anexado integralmente ao processo administrativo (por exemplo, em sistema
eletrônico de informações), como evidência da metodologia e dos parâmetros utilizados na pesquisa de preços.
</p>
</div>

<div class="section">
<h2>7. Quadro-resumo de preço de referência por unidade de medida</h2>
<p>
O quadro abaixo apresenta, para cada unidade de medida, as principais medidas de tendência central
e o preço de referência sugerido, bem como o intervalo de referência construído a partir da média
saneada (ou de seu valor substituto, quando aplicável).
</p>
"""

    if quadro_html_rows:
        html += f"""
<table>
  <thead>
    <tr>
      <th>Unidade de medida</th>
      <th>Média</th>
      <th>Mediana</th>
      <th>Média saneada</th>
      <th>Preço de referência sugerido</th>
      <th>Limite inferior (intervalo)</th>
      <th>Limite superior (intervalo)</th>
    </tr>
  </thead>
  <tbody>
    {quadro_html_rows}
  </tbody>
</table>
<p>
Este quadro pode ser transcrito ou referenciado diretamente na nota técnica ou no despacho
decisório como síntese dos parâmetros adotados para a estimativa de preços.
</p>
"""
    else:
        html += "<p>Não foi possível montar o quadro-resumo por falta de dados consolidados.</p>"

    html += """
</div>

</body>
</html>
"""

    with open(caminho_html, "w", encoding="utf-8") as f:
        f.write(html)

    print("✅ Relatório HTML gerado com sucesso.")
    return caminho_html


# ============================================================
# 🚀 FUNÇÃO PRINCIPAL
# ============================================================

def main():
    cod_item = COD_ITEM_CATALOGO if COD_ITEM_CATALOGO is not None else None

    data_inicial, data_final = calcular_intervalo_ultimo_ano()

    filtros = montar_filtros_opcionais()

    filtros_efetivos = {
        "codItemCatalogo": cod_item if cod_item is not None else "",
        "orgaoEntidadeCnpj": ORGAO_ENTIDADE_CNPJ,
        "unidadeOrgaoCodigoUnidade": UNIDADE_ORGAO_CODIGO_UNIDADE,
        "situacaoCompraItem": SITUACAO_COMPRA_ITEM,
        "materialOuServico": MATERIAL_OU_SERVICO,
        "codigoClasse": CODIGO_CLASSE,
        "codigoGrupo": CODIGO_GRUPO,
        "codFornecedor": COD_FORNECEDOR,
        "temResultado": FILTRAR_TEM_RESULTADO,
        "bps": FILTRAR_BPS,
        "margemPreferenciaNormal": FILTRAR_MARGEM_PREFERENCIA_NORMAL,
        "codigoNCM": CODIGO_NCM,
    }
    filtros_efetivos = {
        k: v for k, v in filtros_efetivos.items()
        if v not in (None, "", [])
    }

    resultados = buscar_itens_pncp(
        cod_item_catalogo=cod_item,
        data_inicial=data_inicial,
        data_final=data_final,
        filtros_opcionais=filtros,
        tamanho_pagina=500,
    )

    df_dados, resumo_df, preco_ref_df = preparar_dataframes(resultados)

    if NOME_BASE_SAIDA:
        base = NOME_BASE_SAIDA
    else:
        cod_str = str(cod_item) if cod_item is not None else "sem_item"
        base = f"pncp_itens_param_{cod_str}_{data_inicial}_a_{data_final}"

    caminho_excel = f"{base}.xlsx"
    caminho_html = f"{base}.html"

    salvar_resultados_em_excel(df_dados, resumo_df, preco_ref_df, caminho_excel)

    meta = {
        "data_inicial": data_inicial,
        "data_final": data_final,
        "filtros_efetivos": filtros_efetivos,
    }

    gerar_relatorio_html(df_dados, resumo_df, preco_ref_df, meta, caminho_html)

    print("==============================================")
    print(" Processo concluído (v3.3 – Excel + HTML sem seção de gráficos).")
    print(f" Arquivo Excel: {caminho_excel}")
    print(f" Relatório HTML: {caminho_html}")
    print("==============================================")


# ============================================================
# 🏁 PONTO DE ENTRADA
# ============================================================

if __name__ == "__main__":
    # Opcional: teste local
    # main()
    pass

