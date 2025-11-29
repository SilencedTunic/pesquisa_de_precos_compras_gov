import random
import streamlit as st
import pncp_backend  # backend com a lógica de consulta e geração de arquivos

# ============================================================
# ⚙️ CONFIGURAÇÃO BÁSICA DA PÁGINA
# ============================================================

st.set_page_config(
    page_title="Pesquisa de Preços PNCP – Lei 14.133/2021",
    page_icon="💸",
    layout="wide",
)

LOGO_PATH = "logos_fortfisc_fundoamazonia.png"

# ============================================================
# 🎨 ESTILO GLOBAL (TEMA CLARO, MODERNO)
# ============================================================

st.markdown(
    """
    <style>
    /* Fundo geral em tom claro com gradiente suave */
    [data-testid="stAppViewContainer"] {
        background: linear-gradient(180deg, #eef2ff 0%, #ecfdf5 55%, #f9fafb 100%);
        color: #111827;
    }

    /* Área principal mais estreita e com respiro */
    .main .block-container {
        padding-top: 1.5rem;
        padding-bottom: 2.5rem;
        max-width: 1200px;
    }

    /* Hero do cabeçalho em cartão branco */
    .pncp-hero {
        padding: 1.5rem 2rem;
        border-radius: 1.5rem;
        background: #ffffff;
        border: 1px solid #e5e7eb;
        box-shadow: 0 18px 40px rgba(15,23,42,0.08);
        margin-bottom: 1.5rem;
    }

    .pncp-hero-title {
        font-size: 2.1rem;
        font-weight: 800;
        letter-spacing: 0.03em;
        color: #111827;
        margin-bottom: 0.25rem;
    }

    .pncp-hero-subtitle {
        font-size: 0.96rem;
        color: #4b5563;
        margin-bottom: 0.4rem;
    }

    .pncp-hero-credit {
        font-size: 0.8rem;
        color: #6b7280;
        font-style: italic;
    }

    .pncp-badge {
        display: inline-flex;
        align-items: center;
        gap: 0.4rem;
        padding: 0.18rem 0.75rem;
        border-radius: 999px;
        background: #eff6ff;
        color: #1d4ed8;
        font-size: 0.76rem;
        margin-top: 0.45rem;
        text-transform: uppercase;
        letter-spacing: 0.08em;
    }

    .pncp-badge-dot {
        width: 8px;
        height: 8px;
        border-radius: 999px;
        background: #22c55e;
        box-shadow: 0 0 0 4px rgba(34,197,94,0.25);
    }

    /* Box de ajuda à direita */
    .pncp-help-box {
        background: #ffffff;
        border-radius: 1rem;
        padding: 1rem 1.1rem;
        border: 1px solid #e5e7eb;
        font-size: 0.9rem;
        color: #111827;
        box-shadow: 0 10px 25px rgba(15,23,42,0.04);
    }

    .pncp-help-box ol {
        margin: 0;
    }

    .pncp-help-box li {
        margin-bottom: 0.25rem;
    }

    /* Títulos das seções */
    .stMarkdown h2, .stMarkdown h3 {
        color: #111827;
    }

    /* Botão principal estilo pill */
    div.stButton > button {
        background: linear-gradient(135deg, #22c55e, #16a34a);
        color: #0b1120;
        border-radius: 999px;
        border: none;
        padding: 0.55rem 1.4rem;
        font-weight: 700;
        letter-spacing: 0.03em;
        text-transform: uppercase;
        font-size: 0.86rem;
        box-shadow: 0 10px 24px rgba(34,197,94,0.35);
    }
    div.stButton > button:hover {
        background: linear-gradient(135deg, #4ade80, #22c55e);
        box-shadow: 0 14px 30px rgba(34,197,94,0.5);
    }

    /* Inputs claros */
    .stTextInput > div > div > input,
    .stSelectbox > div > div > select {
        background-color: #ffffff;
        color: #111827;
        border-radius: 0.6rem;
        border: 1px solid #d1d5db;
    }

    /* Expander com visual clean */
    .streamlit-expanderHeader {
        background: #f9fafb;
        color: #111827;
        border-radius: 0.7rem;
        border: 1px solid #e5e7eb;
        font-weight: 600;
    }

    /* Tabs */
    .stTabs [data-baseweb="tab-list"] {
        gap: 0.4rem;
    }
    .stTabs [data-baseweb="tab"] {
        background: #f3f4f6;
        border-radius: 999px;
        color: #111827;
    }

    /* Mensagens info/sucesso/aviso */
    .stAlert {
        border-radius: 0.9rem;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# ============================================================
# 🧩 CABEÇALHO COM LOGO E TÍTULO
# ============================================================

st.markdown('<div class="pncp-hero">', unsafe_allow_html=True)
col_logo, col_titulo = st.columns([1, 3])

with col_logo:
    st.image(LOGO_PATH, use_column_width=True)

with col_titulo:
    st.markdown(
        '<div class="pncp-hero-title">Pesquisa de Preços PNCP – Lei 14.133/2021</div>',
        unsafe_allow_html=True,
    )
    st.markdown(
        '<div class="pncp-hero-subtitle">'
        'Ferramenta digital para apoiar pesquisas de mercado em compras públicas, '
        'com base em contratações registradas no PNCP.'
        '</div>',
        unsafe_allow_html=True,
    )
    st.markdown(
        '<div class="pncp-hero-credit">'
        'Projeto da Coordenação de Assuntos Estratégicos de Proteção Ambiental – Copes/Dipro/Ibama'
        '</div>',
        unsafe_allow_html=True,
    )
    st.markdown(
        """
        <div class="pncp-badge">
            <span class="pncp-badge-dot"></span>
            <span>Janela temporal: últimos 12 meses de inclusão no PNCP</span>
        </div>
        """,
        unsafe_allow_html=True,
    )
st.markdown("</div>", unsafe_allow_html=True)

# ============================================================
# 🧱 LAYOUT PRINCIPAL – FORMULÁRIO + AJUDA
# ============================================================

col_filtros, col_ajuda = st.columns([2, 1])

# ------------------------------------------------------------
# 🧮 FORMULÁRIO DE FILTROS
# ------------------------------------------------------------
with col_filtros:
    with st.form("filtros_pncp"):
        st.subheader("Configuração da pesquisa")

        # ---------------- Filtros principais ----------------
        with st.expander("🎯 Filtros principais", expanded=True):
            cod_item_str = st.text_input(
                "Código do item de catálogo (CATMAT/CATSER)",
                placeholder="Ex.: 630754",
                help="Informe o código CATMAT/CATSER, quando já souber exatamente o item que deseja analisar.",
            )

            material_ou_servico = st.selectbox(
                "Tipo do objeto",
                options=["(todos)", "Material (M)", "Serviço (S)"],
                index=0,
                help="Você pode limitar a busca apenas a materiais (M) ou a serviços (S).",
            )

        # ---------------- Filtros avançados ----------------
        with st.expander("🧪 Filtros avançados (opcionais)", expanded=False):
            st.caption("Use estes campos para refinar a pesquisa quando necessário.")

            col_a1, col_a2, col_a3 = st.columns(3)
            with col_a1:
                orgao_cnpj = st.text_input(
                    "CNPJ do órgão – opcional",
                    value="",
                    placeholder="Ex.: 00394494000136",
                    help="Número do CNPJ da entidade do órgão (somente números).",
                )
            with col_a2:
                unidade_orgao = st.text_input(
                    "Código da unidade – opcional",
                    value="",
                    placeholder="Ex.: 200350",
                    help="Código da unidade do órgão responsável pela contratação.",
                )
            with col_a3:
                situacao_item = st.text_input(
                    "Situação do item – opcional",
                    value="",
                    placeholder="Ex.: 4 para 'Deserto'",
                    help="Código da situação do item (ex.: 4 = Deserto, 3 = Fracassado etc.).",
                )

            col_b1, col_b2, col_b3 = st.columns(3)
            with col_b1:
                codigo_classe_str = st.text_input(
                    "Código da classe – opcional",
                    value="",
                    placeholder="Ex.: 8145",
                    help="Código da classe no catálogo de materiais/serviços.",
                )
            with col_b2:
                codigo_grupo_str = st.text_input(
                    "Código do grupo – opcional",
                    value="",
                    placeholder="Ex.: 01",
                    help="Código do grupo no catálogo.",
                )
            with col_b3:
                cod_fornecedor = st.text_input(
                    "Código do fornecedor – opcional",
                    value="",
                    placeholder="Ex.: código interno do fornecedor",
                    help="Utilize quando quiser focar em um fornecedor específico.",
                )

            col_c1, col_c2, _ = st.columns(3)
            with col_c1:
                tem_resultado_opt = st.selectbox(
                    "Resultado da compra",
                    options=["(todos)", "Somente com resultado", "Somente sem resultado"],
                    index=0,
                    help="Permite mostrar apenas itens que já têm resultado registrado ou ainda não têm.",
                )
            with col_c2:
                mpn_opt = st.selectbox(
                    "Margem de preferência normal",
                    options=["(todos)", "Somente com margem", "Somente sem margem"],
                    index=0,
                    help="Filtra compras com aplicação de margem de preferência normal.",
                )

        # ---------------- Nome base dos arquivos ----------------
        nome_base = st.text_input(
            "Nome base dos arquivos de saída – opcional",
            value="",
            placeholder="Ex.: pesquisa_container_2025",
            help="Se informado, será usado como prefixo do nome da planilha e da nota técnica.",
        )

        executar = st.form_submit_button("🔎 Executar pesquisa")

# ------------------------------------------------------------
# 💡 COLUNA DIREITA – COMO USAR
# ------------------------------------------------------------
with col_ajuda:
    st.subheader("Guia rápido")

    st.markdown(
        """
        <div class="pncp-help-box">
        <ol style="padding-left: 1.1rem;">
          <li>Comece pelo <strong>código CATMAT/CATSER</strong> ou pela <strong>classe</strong>.</li>
          <li>Use os filtros avançados apenas quando precisar afinar a busca.</li>
          <li>Clique em <strong>“Executar pesquisa”</strong> e aguarde a coleta dos dados.</li>
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
        - Campos deixados em branco **não restringem a consulta**.
        """
    )

# ============================================================
# 🧠 FUNÇÕES AUXILIARES (FRONTEND)
# ============================================================

def _opt_to_bool(opt: str, true_label: str, false_label: str):
    if opt == true_label:
        return True
    if opt == false_label:
        return False
    return None


def _parse_int_or_none(text: str, campo_nome: str):
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


MOTIVATIONAL_MESSAGES = [
    "Cada pesquisa bem feita é um processo a menos para dar dor de cabeça lá na frente. Você está jogando no time da prevenção. 💼✨",
    "Compras públicas com critério são política pública na veia. Obrigado por segurar essa linha de frente. 💚",
    "Enquanto os dados chegam, lembra: transparência também é inovação – e você está empurrando o sistema para frente. 🚀",
    "Servidor de compras raiz sabe: planilha bem feita é escudo contra questionamento. Você está reforçando esse escudo agora. 🛡️",
    "Seu trabalho aqui vira escola, veículo, fiscalização, política ambiental. Não é ‘só’ pesquisa de preços. 🌳",
    "O futuro da Lei 14.133 são pessoas como você, que não têm medo de dado nem de processo. Segue firme. 🔍",
]

# ============================================================
# 🚀 EXECUÇÃO DA PESQUISA
# ============================================================

if executar:
    st.info("Sua pesquisa está sendo preparada. Respira fundo, pega um café e deixa o sistema trabalhar por você. ☕")

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

    mensagem_spinner = random.choice(MOTIVATIONAL_MESSAGES)
    with st.spinner(f"{mensagem_spinner}\n\n(Aguarde: buscando registros no PNCP e montando os arquivos...)"):
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
            margem_pref_normal=mpn,
            nome_base_saida=nome_base or None,
        )

    # ========================================================
    # 📊 APRESENTAÇÃO DOS RESULTADOS
    # ========================================================
    base = meta.get("nome_base", "pncp_pesquisa")

    st.markdown("### Resumo dos filtros aplicados")
    filtros_efetivos = meta.get("filtros_efetivos", {})
    if filtros_efetivos:
        st.json(filtros_efetivos)
    else:
        st.caption("Nenhum filtro adicional foi aplicado além do período de 12 meses.")

    if not excel_bytes:
        st.warning(
            "Nenhum dado foi encontrado para os filtros informados no período considerado. "
            "Ainda assim, uma nota técnica foi gerada registrando a tentativa de pesquisa "
            "no PNCP e os filtros utilizados."
        )

        st.download_button(
            label="⬇️ Baixar nota técnica em HTML",
            data=html_string.encode("utf-8"),
            file_name=f"{base}.html",
            mime="text/html",
        )

        st.subheader("Pré-visualização da nota técnica")
        st.components.v1.html(html_string, height=700, scrolling=True)

    else:
        st.success("Pesquisa concluída com sucesso! Bora usar esses dados a seu favor. ✅")

        tab_downloads, tab_preview = st.tabs(["📂 Downloads", "📝 Nota técnica (visualização)"])

        with tab_downloads:
            st.markdown("#### Arquivos gerados")

            st.download_button(
                label="⬇️ Baixar planilha Excel",
                data=excel_bytes,
                file_name=f"{base}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

            st.download_button(
                label="⬇️ Baixar nota técnica em HTML",
                data=html_string.encode("utf-8"),
                file_name=f"{base}.html",
                mime="text/html",
            )

            st.caption("Dica: anexe os arquivos ao processo (ex.: no SEI) junto com o ETP ou o TR.")

        with tab_preview:
            st.subheader("Visualização da nota técnica")
            st.components.v1.html(html_string, height=700, scrolling=True)
