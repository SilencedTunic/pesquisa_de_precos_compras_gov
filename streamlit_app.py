import streamlit as st
import pncp_backend  # nosso backend com toda a lógica

st.set_page_config(
    page_title="Pesquisa de Preços PNCP – Lei 14.133/2021",
    layout="wide",
)

st.title("Pesquisa de Preços PNCP – Lei 14.133/2021")
st.markdown(
    "Aplicação para pesquisa automatizada de preços no PNCP, "
    "com geração de planilha Excel e nota técnica em HTML."
)

with st.form("filtros_pncp"):
    st.subheader("Filtros da consulta")

    cod_item_str = st.text_input("Código do item de catálogo (CATMAT/CATSER) – opcional")
    orgao_cnpj = st.text_input("CNPJ do órgão (opcional)", value="")
    unidade_orgao = st.text_input("Código da unidade do órgão (opcional)", value="")
    situacao_item = st.text_input("Situação do item (ex.: 4 para 'Deserto') – opcional", value="")

    material_ou_servico = st.selectbox(
        "Material ou Serviço",
        options=["(sem filtro)", "Material (M)", "Serviço (S)"],
        index=0,
    )

    col1, col2, col3 = st.columns(3)
    with col1:
        codigo_classe_str = st.text_input("Código da classe – opcional", value="")
    with col2:
        codigo_grupo_str = st.text_input("Código do grupo – opcional", value="")
    with col3:
        cod_fornecedor = st.text_input("Código do fornecedor – opcional", value="")

    col4, col5, col6 = st.columns(3)
    with col4:
        tem_resultado_opt = st.selectbox(
            "Filtrar por 'temResultado'?",
            options=["(não filtrar)", "Somente com resultado", "Somente sem resultado"],
            index=0,
        )
    with col5:
        bps_opt = st.selectbox(
            "Filtrar BPS?",
            options=["(não filtrar)", "Somente BPS verdadeiro", "Somente BPS falso"],
            index=0,
        )
    with col6:
        mpn_opt = st.selectbox(
            "Filtrar margem de preferência normal?",
            options=["(não filtrar)", "Somente com margem", "Somente sem margem"],
            index=0,
        )

    codigo_ncm = st.text_input("Código NCM – opcional", value="")
    nome_base = st.text_input("Nome base dos arquivos de saída (opcional)", value="")

    executar = st.form_submit_button("Executar pesquisa")

if executar:
    st.info("Executando consulta ao PNCP. Isso pode levar alguns segundos...")

    # Converte campos de texto para tipos adequados
    cod_item = int(cod_item_str) if cod_item_str.strip() else None
    unidade_orgao_int = int(unidade_orgao) if unidade_orgao.strip() else None
    codigo_classe = int(codigo_classe_str) if codigo_classe_str.strip() else None
    codigo_grupo = int(codigo_grupo_str) if codigo_grupo_str.strip() else None

    # Converte selects booleanos
    def _opt_to_bool(opt, true_label, false_label):
        if opt == true_label:
            return True
        if opt == false_label:
            return False
        return None

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

    if not excel_bytes:
        st.error("Nenhum dado retornado ou erro na geração do Excel.")
    else:
        base = meta.get("nome_base", "pncp_pesquisa")

        st.success("Pesquisa concluída com sucesso!")

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

        # Visualização da nota técnica na própria aplicação
        st.subheader("Pré-visualização da nota técnica")
        st.components.v1.html(html_string, height=700, scrolling=True)
