"""
Coletor automático de itens de contratações PNCP (Lei 14.133/2021)
Versão 4.2 – backend focado em correção de requisição API (Headers, Paginação e Chunking de Datas).

Este módulo é pensado para ser usado pelo frontend em Streamlit.
A função pública principal é `executar_pesquisa_e_gerar_arquivos`.
"""

from __future__ import annotations

import base64
import io
import os
from dataclasses import dataclass
from datetime import date, timedelta
from typing import Any, Dict, List, Optional, Tuple

import numpy as np
import openpyxl  # noqa: F401  # usado como engine do Excel
import pandas as pd
import requests
from io import BytesIO


# ============================================================
# 🔧 UTILITÁRIOS GERAIS
# ============================================================

def gerar_intervalos_mensais(data_inicial: date, data_final: date) -> List[Tuple[str, str]]:
    """
    Fraciona um grande intervalo de datas em blocos menores (ex: 30 dias).
    Garante que a API não recuse requisições por limite de dias estourado.
    """
    intervalos = []
    atual_inicio = data_inicial
    while atual_inicio <= data_final:
        atual_fim = atual_inicio + timedelta(days=30)
        if atual_fim > data_final:
            atual_fim = data_final
        intervalos.append((atual_inicio.strftime("%Y-%m-%d"), atual_fim.strftime("%Y-%m-%d")))
        atual_inicio = atual_fim + timedelta(days=1)
    return intervalos

def bool_to_api_flag(value: Optional[bool]) -> Optional[str]:
    """Converte True/False/None em 'true'/'false'/None para a API."""
    if isinstance(value, bool):
        return "true" if value else "false"
    return None

def limpar_valor(valor: Any) -> Any:
    """Normaliza valores opcionais."""
    if valor is None:
        return None
    if isinstance(valor, str):
        s = valor.strip()
        return s if s else None
    return valor


# ============================================================
# 🧮 MONTAGEM DOS PARÂMETROS DA API
# ============================================================

def montar_parametros_consulta(
    data_inicial: str,
    data_final: str,
    cod_item_catalogo: Optional[int] = None,
    orgao_cnpj: Optional[str] = None,
    unidade_orgao: Optional[int] = None,
    situacao_item: Optional[str] = None,
    material_ou_servico: Optional[str] = None,
    codigo_classe: Optional[int] = None,
    codigo_grupo: Optional[int] = None,
    cod_fornecedor: Optional[str] = None,
    tem_resultado: Optional[bool] = None,
    bps: Optional[bool] = None,
    margem_pref_normal: Optional[bool] = None,
    codigo_ncm: Optional[str] = None,
) -> Dict[str, Any]:
    
    cod_item_catalogo = limpar_valor(cod_item_catalogo)
    orgao_cnpj = limpar_valor(orgao_cnpj)
    unidade_orgao = limpar_valor(unidade_orgao)
    situacao_item = limpar_valor(situacao_item)
    material_ou_servico = limpar_valor(material_ou_servico)
    codigo_classe = limpar_valor(codigo_classe)
    codigo_grupo = limpar_valor(codigo_grupo)
    cod_fornecedor = limpar_valor(cod_fornecedor)
    tem_resultado_flag = bool_to_api_flag(tem_resultado)
    bps_flag = bool_to_api_flag(bps)
    margem_pref_flag = bool_to_api_flag(margem_pref_normal)
    codigo_ncm = limpar_valor(codigo_ncm)

    # Padroniza para material se grupo/classe for especificado
    if material_ou_servico is None and (codigo_classe is not None or codigo_grupo is not None):
        material_ou_servico = "M"

    params: Dict[str, Any] = {
        "dataInclusaoPncpInicial": data_inicial,
        "dataInclusaoPncpFinal": data_final,
    }

    if cod_item_catalogo is not None:
        params["codItemCatalogo"] = cod_item_catalogo
    if orgao_cnpj is not None:
        params["orgaoEntidadeCnpj"] = orgao_cnpj
    if unidade_orgao is not None:
        params["unidadeOrgaoCodigoUnidade"] = unidade_orgao
    if situacao_item is not None:
        params["situacaoCompraItem"] = situacao_item
    if material_ou_servico is not None:
        params["materialOuServico"] = material_ou_servico
    if codigo_classe is not None:
        params["codigoClasse"] = codigo_classe
    if codigo_grupo is not None:
        params["codigoGrupo"] = codigo_grupo
    if cod_fornecedor is not None:
        params["codFornecedor"] = cod_fornecedor
    if tem_resultado_flag is not None:
        params["temResultado"] = tem_resultado_flag
    if bps_flag is not None:
        params["bps"] = bps_flag
    if margem_pref_flag is not None:
        params["margemPreferenciaNormal"] = margem_pref_flag
    if codigo_ncm is not None:
        params["codigoNCM"] = codigo_ncm

    return params


# ============================================================
# 🌐 CONSULTA PAGINADA À API
# ============================================================

BASE_URL_ITENS_PNCP = (
    "https://dadosabertos.compras.gov.br/"
    "modulo-contratacoes/2_consultarItensContratacoes_PNCP_14133"
)


def buscar_itens_pncp(params_base: Dict[str, Any], tamanho_pagina: int = 100) -> List[Dict[str, Any]]:
    pagina = 1
    todos_resultados: List[Dict[str, Any]] = []

    print(f"📡 Iniciando consulta PNCP (período {params_base.get('dataInclusaoPncpInicial')} a {params_base.get('dataInclusaoPncpFinal')})...")
    
    # Headers adicionados para evitar recusa de conexão (WAF/Firewall do Governo)
    headers = {
        "Accept": "*/*",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36"
    }

    while True:
        params = dict(params_base)
        params["pagina"] = pagina
        params["tamanhoPagina"] = tamanho_pagina

        try:
            resp = requests.get(BASE_URL_ITENS_PNCP, params=params, headers=headers, timeout=60)
        except Exception as exc:  # pragma: no cover
            print(f"❌ Erro de conexão na página {pagina}: {exc}")
            break

        if resp.status_code != 200:
            print(f"❌ Erro HTTP {resp.status_code} na página {pagina}.")
            print(f"   URL chamada: {resp.url}")
            print(f"   Retorno: {resp.text}")  # Exibe erro detalhado da API, se houver
            break

        try:
            dados = resp.json()
        except ValueError:
            print("❌ Erro ao decodificar JSON da resposta.")
            break

        resultados = dados.get("resultado") or []
        total_paginas = dados.get("totalPaginas") or 1

        if not resultados:
            print(f"ℹ️ Página {pagina} veio vazia. Encerrando paginação deste bloco.")
            break

        todos_resultados.extend(resultados)
        print(
            f"   ✅ Página {pagina}/{total_paginas} carregada "
            f"({len(resultados)} itens). Total no bloco: {len(todos_resultados)}"
        )

        if pagina >= total_paginas:
            break

        pagina += 1

    return todos_resultados


# ============================================================
# 📊 ESTATÍSTICAS – MÉDIA SANEADA E OUTROS
# ============================================================

def calcular_media_sanada_serie(serie: pd.Series, cv_limite: float = 25.0) -> float:
    s = pd.to_numeric(serie.dropna(), errors="coerce").dropna()
    if s.empty:
        return float("nan")

    while True:
        m = s.mean()
        dp = s.std(ddof=0)

        if len(s) < 3 or m == 0 or pd.isna(m) or pd.isna(dp):
            return float(m)

        cv = abs(dp / m) * 100.0
        if cv <= cv_limite:
            return float(m)

        li = m - dp
        ls = m + dp
        filtrado = s[(s >= li) & (s <= ls)]

        if len(filtrado) == len(s) or filtrado.empty:
            return float(m)

        s = filtrado


def calcular_resumo_por_unidade(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty or "unidadeMedida" not in df.columns or "valorUnitarioResultado" not in df.columns:
        return pd.DataFrame()

    df_local = df.copy()
    df_local["valorUnitarioResultado"] = pd.to_numeric(
        df_local["valorUnitarioResultado"], errors="coerce"
    )

    grp = df_local.groupby("unidadeMedida")["valorUnitarioResultado"]

    resumo = grp.agg(["count", "mean", "median", "std", "min", "max"]).rename(
        columns={
            "count": "resultado_qtde",
            "mean": "resultado_media",
            "median": "resultado_mediana",
            "std": "resultado_desvio_padrao",
            "min": "resultado_minimo",
            "max": "resultado_maximo",
        }
    )

    media_sanada = grp.apply(calcular_media_sanada_serie).rename("media_sanada")
    resumo = resumo.join(media_sanada, how="left")

    base_calc = resumo["media_sanada"].fillna(resumo["resultado_media"]).fillna(
        resumo["resultado_mediana"]
    )
    dp_calc = resumo["resultado_desvio_padrao"].fillna(0)

    resumo["limite_inferior_intervalo"] = (base_calc - dp_calc).clip(lower=0)
    resumo["limite_superior_intervalo"] = (base_calc + dp_calc).clip(lower=0)

    return resumo.reset_index().sort_values("unidadeMedida")


def montar_preco_referencia(resumo_df: pd.DataFrame) -> pd.DataFrame:
    if resumo_df.empty:
        return pd.DataFrame()

    cols = ["unidadeMedida", "resultado_media", "resultado_mediana", "media_sanada"]
    cols_existentes = [c for c in cols if c in resumo_df.columns]

    df_out = resumo_df[cols_existentes].copy()
    rename_map = {
        "resultado_media": "media",
        "resultado_mediana": "mediana",
    }
    df_out = df_out.rename(columns=rename_map)
    return df_out


def preparar_dataframes(lista_dados: List[Dict[str, Any]]) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    df = pd.DataFrame(lista_dados)
    if df.empty:
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

    cols_order = [
        "idContratacaoPNCP", "idCompra", "idCompraItem", "orgaoEntidadeCnpj",
        "unidadeOrgaoCodigoUnidade", "descricaoResumida", "materialOuServicoNome",
        "codigoClasse", "codigoGrupo", "codItemCatalogo", "unidadeMedida",
        "quantidade", "valorUnitarioEstimado", "valorTotal", "quantidadeResultado",
        "valorUnitarioResultado", "valorTotalResultado", "situacaoCompraItemNome",
        "nomeFornecedor", "dataInclusaoPncp",
    ]

    cols_existentes = [c for c in cols_order if c in df.columns]
    cols_extras = [c for c in df.columns if c not in cols_existentes]
    df = df[cols_existentes + cols_extras]

    resumo_df = calcular_resumo_por_unidade(df)
    preco_ref_df = montar_preco_referencia(resumo_df)

    return df, resumo_df, preco_ref_df


# ============================================================
# 📝 GERAÇÃO DO RELATÓRIO HTML
# ============================================================

def gerar_relatorio_html(
    df_dados: pd.DataFrame,
    resumo_df: pd.DataFrame,
    preco_ref_df: pd.DataFrame,
    meta: Dict[str, Any],
    caminho_saida: str,
) -> None:

    total_registros = len(df_dados)
    unidades_distintas = (
        df_dados["unidadeMedida"].nunique() if "unidadeMedida" in df_dados.columns else 0
    )
    hoje_str = date.today().strftime("%d/%m/%Y")

    filtros_rows = ""
    for k, v in meta.get("filtros_efetivos", {}).items():
        filtros_rows += f"<tr><td>{k}</td><td>{v}</td></tr>"

    quadro_rows = ""
    if not preco_ref_df.empty and not resumo_df.empty:
        full_ref = preco_ref_df.merge(
            resumo_df[["unidadeMedida", "limite_inferior_intervalo", "limite_superior_intervalo"]],
            on="unidadeMedida",
            how="left",
        )

        for _, row in full_ref.iterrows():
            um = row.get("unidadeMedida", "-")
            media = row.get("media", 0.0) or 0.0
            mediana = row.get("mediana", 0.0) or 0.0
            saneada = row.get("media_sanada", 0.0) or 0.0
            li = row.get("limite_inferior_intervalo", 0.0) or 0.0
            ls = row.get("limite_superior_intervalo", 0.0) or 0.0

            quadro_rows += f"""
            <tr>
                <td>{um}</td>
                <td>{media:.4f}</td>
                <td>{mediana:.4f}</td>
                <td><strong>{saneada:.4f}</strong></td>
                <td>{saneada:.4f}</td>
                <td>{li:.4f}</td>
                <td>{ls:.4f}</td>
            </tr>
            """

    html_content = f"""
<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <title>Nota Técnica - Pesquisa de Preços PNCP</title>
    <style>
        body {{
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            margin: 40px;
            color: #333;
        }}
        h1 {{
            color: #003366;
            border-bottom: 2px solid #003366;
            padding-bottom: 10px;
        }}
        h2 {{
            color: #00509e;
            margin-top: 30px;
        }}
        table {{
            width: 100%;
            border-collapse: collapse;
            margin-top: 10px;
            font-size: 0.9em;
        }}
        th, td {{
            border: 1px solid #ddd;
            padding: 8px;
            text-align: left;
        }}
        th {{
            background-color: #f2f2f2;
            color: #003366;
        }}
        tr:nth-child(even) {{
            background-color: #f9f9f9;
        }}
        .box {{
            background: #eef;
            padding: 15px;
            border-radius: 5px;
            border: 1px solid #ccf;
        }}
    </style>
</head>
<body>
    <h1>Nota Técnica de Pesquisa de Preços (Lei 14.133/2021)</h1>
    <p><strong>Data de geração:</strong> {hoje_str}</p>

    <h2>1. Parâmetros da pesquisa</h2>
    <div class="box">
        <p><strong>Intervalo temporal total pesquisado:</strong> {meta.get("data_inicial")} a {meta.get("data_final")}</p>
        <table>
            <tr><th>Filtro</th><th>Valor aplicado</th></tr>
            {filtros_rows}
        </table>
    </div>

    <h2>2. Metodologia</h2>
    <p>
        Os dados foram extraídos da API de Dados Abertos do PNCP.
        A <strong>Média Saneada</strong> é calculada por expurgo iterativo de
        outliers com base no coeficiente de variação (CV &gt; 25%).
    </p>

    <h2>3. Resumo da amostra</h2>
    <p>
        Foram encontrados <strong>{total_registros}</strong> registros válidos,
        distribuídos em <strong>{unidades_distintas}</strong> unidade(s) de medida.
    </p>

    <h2>4. Quadro de referência de preços</h2>
    <table>
        <thead>
            <tr>
                <th>Unidade</th>
                <th>Média</th>
                <th>Mediana</th>
                <th>Média saneada</th>
                <th>Preço ref. sugerido</th>
                <th>Limite inf.</th>
                <th>Limite sup.</th>
            </tr>
        </thead>
        <tbody>
            {quadro_rows}
        </tbody>
    </table>

    <p>
        <em>
        Este relatório documenta a pesquisa de preços realizada em dados do PNCP,
        podendo ser anexado ao Estudo Técnico Preliminar ou ao Documento de
        Formalização de Demanda, em consonância com a Lei 14.133/2021.
        </em>
    </p>
</body>
</html>
"""

    with open(caminho_saida, "w", encoding="utf-8") as f:
        f.write(html_content)


# ============================================================
# 🚀 FUNÇÃO PRINCIPAL – INTERFACE COM O FRONTEND
# ============================================================

def executar_pesquisa_e_gerar_arquivos(
    cod_item_catalogo: Optional[int] = None,
    orgao_cnpj: Optional[str] = "",
    unidade_orgao: Optional[int] = None,
    situacao_item: Optional[str] = "",
    material_ou_servico: Optional[str] = "",
    codigo_classe: Optional[int] = None,
    codigo_grupo: Optional[int] = None,
    cod_fornecedor: Optional[str] = "",
    tem_resultado: Optional[bool] = None,
    bps: Optional[bool] = None,
    margem_pref_normal: Optional[bool] = None,
    codigo_ncm: Optional[str] = "",
    nome_base_saida: Optional[str] = None,
):
    
    # 1) Intervalo base temporal - Últimos 365 dias, divididos em blocos de até 30 dias
    data_fim_obj = date.today()
    data_ini_obj = data_fim_obj - timedelta(days=365)
    
    intervalos = gerar_intervalos_mensais(data_ini_obj, data_fim_obj)
    lista_resultados = []
    params_info = {}

    for ini, fim in intervalos:
        # 2) Monta parâmetros para a API para cada bloco
        params = montar_parametros_consulta(
            data_inicial=ini,
            data_final=fim,
            cod_item_catalogo=cod_item_catalogo,
            orgao_cnpj=orgao_cnpj,
            unidade_orgao=unidade_orgao,
            situacao_item=situacao_item,
            material_ou_servico=material_ou_servico,
            codigo_classe=codigo_classe,
            codigo_grupo=codigo_grupo,
            cod_fornecedor=cod_fornecedor,
            tem_resultado=tem_resultado,
            bps=bps,
            margem_pref_normal=margem_pref_normal,
            codigo_ncm=codigo_ncm,
        )

        if not params_info: # Guarda filtros globais apenas uma vez para o relatório
            params_info = {k: v for k, v in params.items() if "data" not in k}

        # 3) Consulta iterativa com tamanho de página seguro (100)
        resultados_chunk = buscar_itens_pncp(params, tamanho_pagina=100)
        lista_resultados.extend(resultados_chunk)

    print(f"📦 Total GERAL de registros obtidos: {len(lista_resultados)}")

    # 4) Consolida em DataFrames
    df_dados, resumo_df, preco_ref_df = preparar_dataframes(lista_resultados)

    data_ini_str = data_ini_obj.strftime("%Y-%m-%d")
    data_fim_str = data_fim_obj.strftime("%Y-%m-%d")

    # 5) Define nome-base dos arquivos
    if nome_base_saida:
        base_name = nome_base_saida
    else:
        if params_info.get("codItemCatalogo"):
            ident = f"item_{params_info['codItemCatalogo']}"
        elif params_info.get("codigoClasse"):
            ident = f"classe_{params_info['codigoClasse']}"
        elif params_info.get("codigoGrupo"):
            ident = f"grupo_{params_info['codigoGrupo']}"
        elif params_info.get("orgaoEntidadeCnpj"):
            ident = f"cnpj_{params_info['orgaoEntidadeCnpj']}"
        else:
            ident = "geral"
        base_name = f"pesquisa_pncp_{ident}_{data_ini_str}"

    # 6) Gera Excel em memória
    excel_bytes: bytes = b""
    if not df_dados.empty:
        buffer = BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            df_dados.to_excel(writer, sheet_name="dados", index=False)
            if not resumo_df.empty:
                resumo_df.to_excel(writer, sheet_name="resumo_unidade", index=False)
            if not preco_ref_df.empty:
                preco_ref_df.to_excel(writer, sheet_name="preco_referencia", index=False)
        excel_bytes = buffer.getvalue()

    # 7) Gera HTML da Nota Técnica
    caminho_html = f"{base_name}.html"

    meta = {
        "data_inicial": data_ini_str,
        "data_final": data_fim_str,
        "filtros_efetivos": params_info,
        "nome_base": base_name,
    }

    gerar_relatorio_html(df_dados, resumo_df, preco_ref_df, meta, caminho_html)

    html_string = ""
    if os.path.exists(caminho_html):
        with open(caminho_html, "r", encoding="utf-8") as f:
            html_string = f.read()
        try:
            os.remove(caminho_html)
        except Exception:
            pass

    return excel_bytes, html_string, meta
