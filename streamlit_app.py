import streamlit as st
import pncp_backend  # nosso backend com toda a lógica

# ============================================================
# ⚙️ CONFIGURAÇÃO BÁSICA DA PÁGINA
# ============================================================

st.set_page_config(
    page_title="Pesquisa de Preços PNCP – Lei 14.133/2021",
    layout="wide",
)

# Pequeno ajuste visual no fundo e nos títulos (CSS leve)
st.markdown(
    """
    <style>
    .main {
        background-color: #f7f9fc;
    }
    .pncp-header {
        padding: 0.5rem 0 1rem 0;
        border-bottom: 1px solid #e0e0e0;
        margin-bottom: 1rem;
    }
    .pncp-badge {
        display: inline-block;
        padding: 0.2rem 0.6rem;
        border-radius: 999px;
        background-color: #e6f2ff;
        color: #003366;
        font-size: 0.8rem;
        margin-top: 0.2rem;
    }
    .pncp-help-box {
        background-color: #ffffff;
        border-radius: 0.5rem;
        padding: 0.8rem 1rem;
        border: 1px solid #e0e0e0;
        font-size: 0.9rem;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# ============================================================
# 🧩 CABEÇALHO
# ============================================================

st.markdown('<div class="pncp-header">', unsafe_allow_html=True)

st.title("Pesquisa de Preços PNCP – Lei 14.133/2021")

st.markdown(
    "Aplicação para pesquisa de preços no PNCP, "
    "com geração de planilha Excel e nota técnica em HTML."
)

st.markdown(
    "_Projeto desenvolvido pela Coordenação de Assuntos Estratégicos de Proteção Ambiental - Copes/Dipro/Ibama_"
)

st.markdown(
    '<span class="pncp-badge">Janela temporal: últimos 12 meses de inclusão no PNCP</span>',
    unsafe_allow_html=True,
)

st.markdown("</div>", unsafe_allow_html=True)

# ============================================================
# 🧱 LAYOUT PRINCIPAL – DUAS COLUNAS
# ============================================================

col_filtros, col_ajuda = st.columns([2, 1])

# ------------------------------------------------------------
# 🧮 COLUNA ESQUERDA – FORMULÁRIO DE FILTROS
# ------------------------------------------------------------
with col_filtros:
    with st.form("filtros_pncp"):
        st.subheader("Configuração da pesquisa")

        # ---------------- Filtros principais ----------------
        with st.expander("Filtros principais", expanded=True):
            cod_item_str = st.text_input(
                "Código do item de catálogo (CATMAT/CATSER) – opcional",
                placeholder="Ex.: 279727",
                help="Informe o código do item de catálogo, se quiser restringir a pesquisa a um item específico.",
            )
            orgao_cnpj = st.text_input(
                "CNPJ do órgão – opcional",
                value="",
                placeholder="Ex.: 00394494000136",
                help="Número do CNPJ da entidade do órgão (somente números).",
            )
            unidade_orgao = st.text_input(
                "Código da unidade do órgão – opcional",
                value="",
                placeholder="Ex.: 200350",
                help="Código da unidade do órgão responsável pela contratação.",
            )
            situacao_item = st.text_input(
                "Situação do item – opcional",
                value="",
                placeholder="Ex.: 4 para 'Deserto'",
                help="Código da situação do item da compra (ex.: 4 = Deserto, 3 = Fracassado etc.).",
            )

            material_ou_servico = st.selectbox(
                "Material ou Serviço",
                options=["(sem filtro)", "Material (M)", "Serviço (S)"],
                index=0,
                help="Selecione se deseja restringir a pesquisa apenas a materiais (M) ou serviços (S).",
            )

        # ---------------- Filtros avançados ----------------
        with st.expander("Filtros avançados (opcionais)", expanded=False):
            col1, col2, col3 = st.columns(3)
            with col1:
                codigo_classe_str = st.text_input(
                    "Código da classe – opcional",
                    value="",
                    placeholder="Ex.: 6510",
                    help="Código da classe no catálogo de materiais/serviços.",
                )
            with col2:
                codigo_grupo_str = st.text_input(
                    "Código do grupo – opcional",
                    value="",
                    placeholder="Ex.: 01",
                    help="Código do grupo no catálogo.",
                )
            with col3:
                cod_fornecedor = st.text_input(
                    "Código do fornecedor – opcional",
                    value="",
                    placeholder="Ex.: código interno do fornecedor",
                )

            col4, col5, col6 = st.columns(3)
            with col4:
                tem_resultado_opt = st.selectbox(
                    "Filtrar por 'temResultado'?",
                    options=["(não filtrar)", "Somente com resultado", "Somente sem resultado"],
                    index=0,
                    help="Filtra itens que possuem (ou não) resultado registrado.",
                )
            with col5:
                bps_opt = st.selectbox(
                    "Filtrar BPS?",
                    options=["(não filtrar)", "Somente BPS verdadeiro", "Somente BPS falso"],
                    index=0,
                    help="Filtra se a compra segue (ou não) Boas Práticas de Suprimentos.",
                )
            with col6:
                mpn_opt = st.selectbox(
                    "Filtrar margem de preferência normal?",
                    options=["(não filtrar)", "Somente com margem", "Somente sem margem"],
                    index=0,
                    help="Filtra compras com aplicação de margem de preferência normal.",
                )

            codigo_ncm = st.text_input(
                "Código NCM – opcional",
                value="",
                placeholder="Ex.: 30049099",
                help="Código NCM – Nomenclatura Comum do Mercosul.",
            )

        # ---------------- Nome base dos arquivos ----------------
        nome_base = st.text_input(
            "Nome base dos arquivos de saída – opcional",
            value="",
            placeholder="Ex.: pesquisa_algodao_2024",
            help="Se informado, será usado como prefixo do nome da planilha e da nota técnica.",
        )

        executar = st.form_submit_button("🔎 Executar pesquisa")

# ------------------------------------------------------------
# 📖 COLUNA DIREITA – AJUDA, PASSO A PASSO, OBSERVAÇÕES
# ------------------------------------------------------------
with col_ajuda:
    st.subheader("Como utilizar")

    st.markdown(
        """
        <div class="pncp-help-box">
        <ol style="padding-left: 1.2rem;">
          <li>Preencha, se desejar, o <strong>item de catálogo</strong> ou a <strong>classe</strong>.</li>
          <li>Inclua filtros por <strong>CNPJ</strong>, unidade ou situação do item, se necessários.</li>
          <li>Clique em <strong>“Executar pesquisa”</strong>.</li>
          <li>Baixe a <strong>planilha Excel</strong> e/ou a <strong>nota técnica em HTML</strong>.</li>
          <li>Anexe os arquivos ao processo (SEI) como evidência da pesquisa de preços.</li>
        </ol>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.markdown("### Observações importantes")
    st.markdown(
        """
        - O período considerado é sempre de **12 meses para trás** a partir da data atual.
        - Se nenhum dado for encontrado, ainda assim será gerada uma **nota técnica**
          registrando a tentativa de pesquisa e os filtros utilizados.
        - Filtros deixados em branco **não são enviados** à API (não restringem a consulta).
        """
    )

# ============================================================
# 🧠 FUNÇÕES AUXILIARES (FRONTEND)
# ============================================================

def _opt_to_bool(opt, true_label, false_label):
    if opt == true_label:
        return True
    if opt == false_label:
        return False
    return None

def _parse_int_or_none(text, campo_nome):
    """
    Converte texto em int, ou retorna None.
    Em caso de erro, mostra um aviso leve na interface.
    """
    text = text.strip()
    if not text:
        return None
    try:
        return int(text)
    except ValueError:
        st.warning(f"Valor inválido em '{campo_nome}'. Ignorando este filtro.")
        return None

# ============================================================
# 🚀 EXECUÇÃO DA PESQUISA
# ============================================================

if executar:
    st.info("Executando consulta ao PNCP. Isso pode levar alguns segundos...")

    # Converte campos de texto para tipos adequados
    cod_item = _parse_int_or_none(cod_item_str, "Código do item de catálogo")
    unidade_orgao_int = _parse_int_or_none(unidade_orgao, "Código da unidade do órgão")
    codigo_classe = _parse_int_or_none(codigo_classe_str, "Código da classe")
    codigo_grupo = _parse_int_or_none(codigo_grupo_str, "Código do grupo")

    # Converte selects booleanos
    tem_resultado = _opt_to_bool(
        tem_resultado_opt,
        "Somente com resultado",
        "Somente sem resultado",
    )
    bps = _opt_to_bool(
        bps_opt,
        "Somente BPS verdadeiro",
        "Somente BPS falso",
    )
    mpn = _opt_to_bool(
        mpn_opt,
        "Somente com margem",
        "Somente sem margem",
    )

    # Mapeia Material/Serviço
    if material_ou_servico == "Material (M)":
        mos = "M"
    elif material_ou_servico == "Serviço (S)":
        mos = "S"
    else:
        mos = ""

    with st.spinner("Consultando API do PNCP e gerando arquivos..."):
        excel_bytes, html_string, meta = pncp_backend.executar_pesquisa_e_gerar_arquivos(
            cod_item_catalogo=cod_item,
            orgao_cnpj=orgao_cnpj,
            unidade_orgao=unidade_orgao_int,
            situacao_item=situacao_item,
            material_ou_servico=mos,
            codigo_classe=codigo_classe,
            codigo_grupo=codigo_grupo,
            cod_fornecedor=cod_fornecedor,
            tem_resultado=tem_resultado,
            bps=bps,
            margem_pref_normal=mpn,
            codigo_ncm=codigo_ncm,
            nome_base_saida=nome_base or None,
        )

    # ========================================================
    # 📊 APRESENTAÇÃO DOS RESULTADOS
    # ========================================================
    base = meta.get("nome_base", "pncp_pesquisa")

    # Resumo dos filtros efetivos aplicados
    st.markdown("### Resumo dos filtros aplicados")
    filtros_efetivos = meta.get("filtros_efetivos", {})
    if filtros_efetivos:
        st.json(filtros_efetivos)
    else:
        st.caption("Nenhum filtro adicional foi aplicado além do período de 12 meses.")

    if not excel_bytes:
        # Não há dados suficientes para montar Excel, mas o HTML ainda é útil
        st.warning(
            "Nenhum dado foi encontrado para os filtros informados no período considerado. "
            "Ainda assim, uma nota técnica foi gerada registrando a tentativa de pesquisa "
            "no PNCP e os filtros utilizados."
        )

        # Download do HTML mesmo sem Excel
        st.download_button(
            label="⬇️ Baixar nota técnica em HTML",
            data=html_string.encode("utf-8"),
            file_name=f"{base}.html",
            mime="text/html",
        )

        st.subheader("Pré-visualização da nota técnica")
        st.components.v1.html(html_string, height=700, scrolling=True)

    else:
        st.success("Pesquisa concluída com sucesso!")

        # Abas para organizar área de resultados
        tab_downloads, tab_preview = st.tabs(["📂 Downloads", "📝 Nota técnica (visualização)"])

        with tab_downloads:
            st.markdown("#### Arquivos gerados")

            # Download da planilha
            st.download_button(
                label="⬇️ Baixar planilha Excel",
                data=excel_bytes,
                file_name=f"{base}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

            # Download do HTML
            st.download_button(
                label="⬇️ Baixar nota técnica em HTML",
                data=html_string.encode("utf-8"),
                file_name=f"{base}.html",
                mime="text/html",
            )

            st.caption("Anexe esses arquivos à instrução processual (por exemplo, no SEI).")

        with tab_preview:
            st.subheader("Visualização da nota técnica")
            st.components.v1.html(html_string, height=700, scrolling=True)
