"""
Coletor automático de itens de contratações PNCP (Lei 14.133/2021)
Versão 4.0 – Refatorada para flexibilidade total de filtros.

Este módulo é o backend de processamento. Ele não depende de configurações globais
hardcoded. Todos os parâmetros devem ser passados via função 'executar_pesquisa_e_gerar_arquivos'.
"""

import requests
import pandas as pd
import numpy as np
import openpyxl  # Engine do Excel
import base64
import os
import io
from datetime import date, timedelta
from io import BytesIO

# ============================================================
# 🛠️ UTILITÁRIOS GERAIS
# ============================================================

def calcular_intervalo_ultimo_ano():
    """
    Retorna (data_inicial, data_final) em formato 'YYYY-MM-DD',
    considerando 'hoje' e 'hoje - 365 dias'.
    """
    data_final = date.today()
    data_inicial = data_final - timedelta(days=365)
    return data_inicial.strftime("%Y-%m-%d"), data_final.strftime("%Y-%m-%d")

def bool_to_api_flag(value):
    """
    Converte True/False/None em 'true'/'false'/None para a API.
    """
    if isinstance(value, bool):
        return "true" if value else "false"
    return None

def limpar_valor(valor):
    """
    Retorna o valor se ele for válido (não vazio, não None), senão retorna None.
    Usado para garantir que strings vazias não sejam enviadas à API.
    """
    if valor is None:
        return None
    if isinstance(valor, str):
        s = valor.strip()
        return s if s else None
    return valor

def montar_parametros_consulta(
    data_inicial,
    data_final,
    cod_item_catalogo=None,
    orgao_cnpj=None,
    unidade_orgao=None,
    situacao_item=None,
    material_ou_servico=None,
    codigo_classe=None,
    codigo_grupo=None,
    cod_fornecedor=None,
    tem_resultado=None,
    bps=None,
    margem_pref_normal=None,
    codigo_ncm=None
):
    """
    Constrói o dicionário de parâmetros limpo para enviar à API.
    Remove chaves com valores None ou vazios para não quebrar a consulta.
    """
    # Dicionário bruto com todos os possíveis filtros
    raw_params = {
        "dataInclusaoPncpInicial": data_inicial,
        "dataInclusaoPncpFinal": data_final,
        "codItemCatalogo": limpar_valor(cod_item_catalogo),
        "orgaoEntidadeCnpj": limpar_valor(orgao_cnpj),
        "unidadeOrgaoCodigoUnidade": limpar_valor(unidade_orgao),
        "situacaoCompraItem": limpar_valor(situacao_item),
        "materialOuServico": limpar_valor(material_ou_servico),
        "codigoClasse": limpar_valor(codigo_classe),
        "codigoGrupo": limpar_valor(codigo_grupo),
        "codFornecedor": limpar_valor(cod_fornecedor),
        "temResultado": bool_to_api_flag(tem_resultado),
        "bps": bool_to_api_flag(bps),
        "margemPreferenciaNormal": bool_to_api_flag(margem_pref_normal),
        "codigoNCM": limpar_valor(codigo_ncm),
    }

    # Remove chaves que ficaram com valor None
    params_limpos = {k: v for k, v in raw_params.items() if v is not None}
    
    return params_limpos

# ============================================================
# 🌐 CONSULTA À API (CORE)
# ============================================================

def buscar_itens_pncp(params_base, tamanho_pagina=500):
    """
    Faz a paginação automática na API do PNCP usando os parâmetros fornecidos.
    
    Args:
        params_base (dict): Dicionário contendo datas e filtros já limpos.
        tamanho_pagina (int): Itens por página (padrão 500).
    
    Returns:
        list: Lista consolidada de dicionários (registros).
    """
    base_url = "https://dadosabertos.compras.gov.br/modulo-contratacoes/2_consultarItensContratacoes_PNCP_14133"
    
    pagina = 1
    todos_resultados = []
    
    print(f"📡 Iniciando consulta ao PNCP...")
    print(f"   Parâmetros efetivos: {params_base}")

    while True:
        # Atualiza a página atual nos parâmetros
        params_request = params_base.copy()
        params_request["pagina"] = pagina
        params_request["tamanhoPagina"] = tamanho_pagina

        try:
            resp = requests.get(base_url, params=params_request, timeout=60)
        except Exception as exc:
            print(f"❌ Erro de conexão na página {pagina}: {exc}")
            break

        if resp.status_code != 200:
            print(f"❌ Erro HTTP {resp.status_code} na página {pagina}.")
            print(f"   URL chamada: {resp.url}")
            break

        try:
            dados = resp.json()
        except ValueError:
            print("❌ Erro ao decodificar JSON da resposta.")
            break

        resultados_pagina = dados.get("resultado", [])
        total_paginas = dados.get("totalPaginas", 1)
        
        if not resultados_pagina:
            print(f"ℹ️ Página {pagina} retornou vazia. Encerrando.")
            break

        todos_resultados.extend(resultados_pagina)
        print(f"   ✅ Página {pagina}/{total_paginas} carregada ({len(resultados_pagina)} itens). Total acumulado: {len(todos_resultados)}")

        # Critérios de parada
        if pagina >= total_paginas:
            break
            
        pagina += 1

    return todos_resultados

# ============================================================
# 📊 CÁLCULOS ESTATÍSTICOS (MÉDIA SANEADA)
# ============================================================

def calcular_media_sanada_serie(serie: pd.Series, cv_limite: float = 25.0) -> float:
    """
    Calcula a média saneada (expurgo iterativo de outliers).
    """
    s = pd.to_numeric(serie.dropna(), errors="coerce").dropna()
    if s.empty:
        return float("nan")

    # Loop de saneamento
    while True:
        m = s.mean()
        dp = s.std(ddof=0)
        
        # Se só tem 1 ou 2 itens, ou média zero, não há como sanear mais
        if len(s) < 3 or m == 0 or pd.isna(m) or pd.isna(dp):
            return m

        cv = abs(dp / m) * 100.0
        if cv <= cv_limite:
            return m

        # Define limites para corte
        li = m - dp
        ls = m + dp
        
        # Filtra (mantém o que está dentro de M +/- DP)
        filtrado = s[(s >= li) & (s <= ls)]

        # Se não excluiu ninguém ou excluiu todo mundo (caso raro), para
        if len(filtrado) == len(s) or filtrado.empty:
            return m

        s = filtrado

def calcular_resumo_por_unidade(df: pd.DataFrame) -> pd.DataFrame:
    """
    Gera estatísticas agrupadas por unidade de medida.
    """
    if df.empty or "unidadeMedida" not in df.columns or "valorUnitarioResultado" not in df.columns:
        return pd.DataFrame()

    df_local = df.copy()
    df_local["valorUnitarioResultado"] = pd.to_numeric(df_local["valorUnitarioResultado"], errors="coerce")

    # Agrupamento
    grp = df_local.groupby("unidadeMedida")["valorUnitarioResultado"]

    # Estatísticas básicas
    resumo = grp.agg(["count", "mean", "median", "std", "min", "max"]).rename(columns={
        "count": "resultado_qtde",
        "mean": "resultado_media",
        "median": "resultado_mediana",
        "std": "resultado_desvio_padrao",
        "min": "resultado_minimo",
        "max": "resultado_maximo",
    })

    # Média Saneada
    media_sanada = grp.apply(calcular_media_sanada_serie).rename("media_sanada")
    resumo = resumo.join(media_sanada, how="left")

    # Intervalos de confiança simples (Base +/- DP)
    # A base preferencial é a média saneada; se nula, usa média ou mediana
    base_calc = resumo["media_sanada"].fillna(resumo["resultado_media"]).fillna(resumo["resultado_mediana"])
    dp_calc = resumo["resultado_desvio_padrao"].fillna(0)

    resumo["limite_inferior_intervalo"] = (base_calc - dp_calc).clip(lower=0)
    resumo["limite_superior_intervalo"] = (base_calc + dp_calc).clip(lower=0)

    return resumo.reset_index().sort_values("unidadeMedida")

def montar_preco_referencia(resumo_df: pd.DataFrame) -> pd.DataFrame:
    """
    Cria uma tabela simplificada para consulta rápida.
    """
    if resumo_df.empty:
        return pd.DataFrame()
        
    cols = ["unidadeMedida", "resultado_media", "resultado_mediana", "media_sanada"]
    # Seleciona apenas as que existem
    cols_existentes = [c for c in cols if c in resumo_df.columns]
    
    df_out = resumo_df[cols_existentes].copy()
    
    # Renomeia para ficar mais amigável
    rename_map = {
        "resultado_media": "media",
        "resultado_mediana": "mediana"
    }
    df_out = df_out.rename(columns=rename_map)
    return df_out

def preparar_dataframes(lista_dados):
    """
    Transforma lista de dicts em DataFrames processados.
    """
    df = pd.DataFrame(lista_dados)
    
    if df.empty:
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

    # Colunas de interesse para ordenação visual no Excel
    cols_order = [
        "idContratacaoPNCP", "idCompra", "orgaoEntidadeCnpj", "unidadeOrgaoCodigoUnidade",
        "descricaoResumida", "materialOuServicoNome", "codigoClasse", "codItemCatalogo",
        "unidadeMedida", "quantidade", "valorUnitarioEstimado", 
        "valorUnitarioResultado", "valorTotalResultado", 
        "situacaoCompraItemNome", "nomeFornecedor", "dataInclusaoPncp"
    ]
    # Mantém colunas existentes, adiciona as extras no final
    cols_existentes = [c for c in cols_order if c in df.columns]
    cols_extras = [c for c in df.columns if c not in cols_existentes]
    df = df[cols_existentes + cols_extras]

    resumo_df = calcular_resumo_por_unidade(df)
    preco_ref_df = montar_preco_referencia(resumo_df)

    return df, resumo_df, preco_ref_df

# ============================================================
# 📝 GERAÇÃO DE RELATÓRIO HTML
# ============================================================

def gerar_relatorio_html(df_dados, resumo_df, preco_ref_df, meta, caminho_saida):
    """
    Gera o HTML da Nota Técnica.
    """
    total_registros = len(df_dados)
    unidades_distintas = df_dados["unidadeMedida"].nunique() if "unidadeMedida" in df_dados.columns else 0
    hoje_str = date.today().strftime("%d/%m/%Y")

    # Monta linhas da tabela de filtros
    filtros_rows = ""
    for k, v in meta.get("filtros_efetivos", {}).items():
        filtros_rows += f"<tr><td>{k}</td><td>{v}</td></tr>"

    # Monta linhas do Quadro Resumo
    quadro_rows = ""
    if not preco_ref_df.empty and not resumo_df.empty:
        # Merge para pegar os limites
        full_ref = preco_ref_df.merge(
            resumo_df[["unidadeMedida", "limite_inferior_intervalo", "limite_superior_intervalo"]],
            on="unidadeMedida", how="left"
        )
        
        for _, row in full_ref.iterrows():
            um = row.get("unidadeMedida", "-")
            media = f"{row.get('media', 0):.4f}"
            mediana = f"{row.get('mediana', 0):.4f}"
            saneada = f"{row.get('media_sanada', 0):.4f}"
            li = f"{row.get('limite_inferior_intervalo', 0):.4f}"
            ls = f"{row.get('limite_superior_intervalo', 0):.4f}"
            
            quadro_rows += f"""
            <tr>
                <td>{um}</td>
                <td>{media}</td>
                <td>{mediana}</td>
                <td><strong>{saneada}</strong></td>
                <td>{saneada}</td>
                <td>{li}</td>
                <td>{ls}</td>
            </tr>
            """

    html_content = f"""
    <!DOCTYPE html>
    <html lang="pt-BR">
    <head>
        <meta charset="UTF-8">
        <title>Nota Técnica - Pesquisa de Preços PNCP</title>
        <style>
            body {{ font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; margin: 40px; color: #333; }}
            h1 {{ color: #003366; border-bottom: 2px solid #003366; padding-bottom: 10px; }}
            h2 {{ color: #00509e; margin-top: 30px; }}
            table {{ width: 100%; border-collapse: collapse; margin-top: 10px; font-size: 0.9em; }}
            th, td {{ border: 1px solid #ddd; padding: 8px; text-align: left; }}
            th {{ background-color: #f2f2f2; color: #003366; }}
            tr:nth-child(even) {{ background-color: #f9f9f9; }}
            .box {{ background: #eef; padding: 15px; border-radius: 5px; border: 1px solid #ccf; }}
        </style>
    </head>
    <body>
        <h1>Nota Técnica de Pesquisa de Preços (Lei 14.133/2021)</h1>
        <p><strong>Data de geração:</strong> {hoje_str}</p>
        
        <h2>1. Parâmetros da Pesquisa</h2>
        <div class="box">
            <p><strong>Intervalo Temporal:</strong> {meta.get('data_inicial')} a {meta.get('data_final')}</p>
            <table>
                <tr><th>Filtro</th><th>Valor Aplicado</th></tr>
                {filtros_rows}
            </table>
        </div>

        <h2>2. Metodologia</h2>
        <p>Os dados foram extraídos da API de Dados Abertos do PNCP. A <strong>Média Saneada</strong> foi calculada utilizando um método iterativo para exclusão de outliers (Coefficient of Variation > 25%).</p>
        
        <h2>3. Resumo da Amostra</h2>
        <p>Foram encontrados <strong>{total_registros}</strong> registros válidos distribuídos em <strong>{unidades_distintas}</strong> unidades de medida.</p>

        <h2>4. Quadro de Referência de Preços</h2>
        <table>
            <thead>
                <tr>
                    <th>Unidade</th>
                    <th>Média</th>
                    <th>Mediana</th>
                    <th>Média Saneada</th>
                    <th>Preço Ref. Sugerido</th>
                    <th>Limite Inf.</th>
                    <th>Limite Sup.</th>
                </tr>
            </thead>
            <tbody>
                {quadro_rows}
            </tbody>
        </table>
        
        <p><em>Este relatório serve como evidência de pesquisa de mercado conforme IN SEGES/ME nº 65/2021 e Lei 14.133/2021.</em></p>
    </body>
    </html>
    """
    
    with open(caminho_saida, "w", encoding="utf-8") as f:
        f.write(html_content)

# ============================================================
# 🚀 FUNÇÃO PRINCIPAL (INTERFACE COM O STREAMLIT)
# ============================================================

def executar_pesquisa_e_gerar_arquivos(
    cod_item_catalogo=None,
    orgao_cnpj=None,
    unidade_orgao=None,
    situacao_item=None,
    material_ou_servico=None,
    codigo_classe=None,
    codigo_grupo=None,
    cod_fornecedor=None,
    tem_resultado=None,
    bps=None,
    margem_pref_normal=None,
    codigo_ncm=None,
    nome_base_saida=None,
):
    """
    Função principal chamada pelo frontend (Streamlit).
    """
    
    # 1. Define intervalo de datas (obrigatório e automático)
    data_ini, data_fim = calcular_intervalo_ultimo_ano()

    # 2. Monta parâmetros limpos (remove Nones e Strings vazias)
    params = montar_parametros_consulta(
        data_inicial=data_ini,
        data_final=data_fim,
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
        codigo_ncm=codigo_ncm
    )

    # 3. Executa a busca
    lista_resultados = buscar_itens_pncp(params)

    # 4. Processa os dados
    df_dados, resumo_df, preco_ref_df = preparar_dataframes(lista_resultados)

    # 5. Gera nome base do arquivo
    if not nome_base_saida:
        # Tenta pegar algum identificador para o nome
        id_nome = "geral"
        if params.get("codItemCatalogo"):
            id_nome = f"item_{params['codItemCatalogo']}"
        elif params.get("codigoClasse"):
            id_nome = f"classe_{params['codigoClasse']}"
        nome_base_saida = f"pesquisa_pncp_{id_nome}_{data_ini}"

    # 6. Gera Excel em memória (BytesIO) para download
    excel_bytes = b""
    if not df_dados.empty:
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df_dados.to_excel(writer, sheet_name="dados", index=False)
            if not resumo_df.empty:
                resumo_df.to_excel(writer, sheet_name="resumo_unidade", index=False)
            if not preco_ref_df.empty:
                preco_ref_df.to_excel(writer, sheet_name="preco_referencia", index=False)
        excel_bytes = output.getvalue()

    # 7. Gera HTML (Temporário -> String -> Deleta)
    # Precisamos salvar em disco primeiro para reaproveitar a função de HTML, ou refatorar ela.
    # Vamos salvar temp e ler.
    caminho_html = f"{nome_base_saida}.html"
    
    # Metadados para o relatório
    meta = {
        "data_inicial": data_ini,
        "data_final": data_fim,
        "filtros_efetivos": {k: v for k, v in params.items() if "data" not in k} # Mostra só os filtros extras
    }
    
    gerar_relatorio_html(df_dados, resumo_df, preco_ref_df, meta, caminho_html)
    
    html_string = ""
    if os.path.exists(caminho_html):
        with open(caminho_html, "r", encoding="utf-8") as f:
            html_string = f.read()
        try:
            os.remove(caminho_html) # Limpa arquivo temporário
        except:
            pass

    return excel_bytes, html_string, meta

if __name__ == "__main__":
    print("Execute via Streamlit ou importe como módulo.")
