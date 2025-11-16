# Pesquisa de Preços PNCP – Lei 14.133/2021

Aplicação web para **automatizar a pesquisa de preços** no Portal Nacional de Contratações Públicas (PNCP), utilizando a API de dados abertos do Compras.gov.br, com:

- Consulta paginada à API oficial do PNCP;
- Geração automática de:
  - **Planilha Excel** com base de dados e resumos estatísticos;
  - **Nota técnica em HTML** com metodologia, filtros e quadro-resumo de preços de referência;
- Interface web amigável (via **Streamlit**) para que qualquer servidor possa usar apenas com um navegador.

---

## 1. Objetivo da aplicação

A ferramenta foi criada para apoiar coordenações que realizam **pesquisa de preços** para contratações regidas pela **Lei nº 14.133/2021**, reduzindo trabalho manual e padronizando:

- Coleta de históricos de contratações no PNCP;
- Tratamento estatístico (média, mediana, média saneada, intervalos de referência);
- Registro transparente dos critérios utilizados, em formato de **nota técnica** anexável ao processo (SEI, por exemplo).

A lógica da **média saneada** e dos **limites inferior/superior** foi implementada para apoiar a análise crítica dos valores, evitando distorções causadas por outliers.

---

## 2. Visão geral da arquitetura

O projeto é composto por três arquivos principais:

- `pncp_backend.py`  
  Contém toda a lógica de:
  - chamada à API do PNCP;
  - paginação e consolidação dos resultados;
  - preparação dos DataFrames;
  - cálculo de estatísticas (incluindo média saneada);
  - geração da planilha Excel;
  - geração da nota técnica em HTML.

- `streamlit_app.py`  
  Contém a **interface web** (frontend + backend leve) feita em Streamlit:
  - exibe o formulário de filtros;
  - aciona a função de backend;
  - oferece botões de download do Excel e do HTML;
  - exibe a nota técnica diretamente na página.

- `requirements.txt`  
  Lista de bibliotecas Python necessárias para rodar a aplicação.

---

## 3. Funcionalidades principais

### 3.1. Consulta automatizada ao PNCP

A aplicação consome o endpoint:

> `modulo-contratacoes/2_consultarItensContratacoes_PNCP_14133`

com os parâmetros de filtro informados pelo usuário.  
A janela temporal é sempre:

- **dataInclusaoPncpInicial** = hoje – 365 dias  
- **dataInclusaoPncpFinal**   = hoje  

A aplicação:
- faz as chamadas paginadas automaticamente (`pagina`, `tamanhoPagina`);
- consolida todos os registros retornados.

### 3.2. Planilha Excel gerada automaticamente

O Excel gerado possui três abas:

1. **`dados`**  
   Base detalhada com os registros retornados pela API.  
   Exemplo de campos principais:
   - `idContratacaoPNCP`
   - `idCompra`, `idCompraItem`
   - `orgaoEntidadeCnpj`, `unidadeOrgaoCodigoUnidade`
   - `descricaoResumida`, `descricaodetalhada`
   - `materialOuServicoNome`
   - `codigoClasse`, `codigoGrupo`, `codItemCatalogo`
   - `unidadeMedida`
   - `quantidade`, `valorUnitarioEstimado`, `valorTotal`
   - `quantidadeResultado`, `valorUnitarioResultado`, `valorTotalResultado`
   - `situacaoCompraItemNome`
   - `nomeFornecedor`
   - `dataInclusaoPncp`, `dataAtualizacaoPncp`, `dataResultado`
   - `codigoNCM`, `descricaoNCM`

2. **`resumo_unidade`**  
   Estatísticas por `unidadeMedida`, considerando **valorUnitarioResultado**, incluindo:
   - `resultado_qtde`
   - `resultado_media`
   - `resultado_mediana`
   - `resultado_desvio_padrao`
   - `resultado_minimo`
   - `resultado_maximo`
   - `media_sanada`  
   - `limite_inferior_intervalo`
   - `limite_superior_intervalo`

3. **`preco_referencia`**  
   Tabela compacta por unidade de medida, com:
   - `unidadeMedida`
   - `media`
   - `mediana`
   - `media_sanada`

Essas abas podem ser usadas diretamente para instruir a **nota técnica de pesquisa de preços**.

### 3.3. Nota técnica em HTML

A aplicação gera um arquivo HTML com a estrutura de relatório, contendo:

1. **Introdução**  
   Contextualiza o uso do PNCP e da Lei 14.133/2021.

2. **Período e filtros utilizados**  
   - Datas de inclusão no PNCP consideradas;
   - Tabela com os filtros efetivamente aplicados (CNPJ, unidade, classe, codItemCatalogo, etc.).

3. **Estatísticas descritivas da amostra**  
   - Número total de registros;
   - Número de unidades de medida distintas;
   - Estatísticas globais de `valorUnitarioResultado` (mínimo, máximo, média, mediana, desvio-padrão), quando disponíveis.

4. **Metodologia de cálculo**  
   - Descrição da lógica de extração e tratamento;
   - Explicação da **média saneada** (expurgo iterativo por desvio-padrão e coeficiente de variação – CV).

6. **Resultados e uso recomendado**  
   - Orientações de uso das abas `dados`, `resumo_unidade` e `preco_referencia`;
   - Sugestão de uso da média saneada e da mediana como referência, a depender do contexto.

7. **Quadro-resumo de preço de referência por unidade de medida**  
   - Tabela com:
     - unidade de medida;
     - média, mediana, média saneada;
     - preço de referência sugerido;
     - limites inferior e superior do intervalo.

Esse HTML pode ser:
- baixado e anexado diretamente ao processo (SEI);
- visualizado dentro da própria aplicação web.

---

## 4. Parâmetros de filtro disponíveis

Na interface web (Streamlit), o usuário pode informar:

- **codItemCatalogo** (CATMAT/CATSER) – opcional  
- **orgaoEntidadeCnpj** – opcional  
- **unidadeOrgaoCodigoUnidade** – opcional  
- **situacaoCompraItem** – opcional (ex.: `4` = “Deserto” etc.)  
- **materialOuServico** – opcional:
  - Sem filtro (valor em branco)
  - `M` = Material
  - `S` = Serviço  
- **codigoClasse** – opcional  
- **codigoGrupo** – opcional  
- **codFornecedor** – opcional  
- **temResultado** – opcional:
  - não filtrar
  - somente com resultado
  - somente sem resultado  
- **bps** – opcional:
  - não filtrar
  - somente BPS = true
  - somente BPS = false  
- **margemPreferenciaNormal** – opcional:
  - não filtrar
  - somente com margem
  - somente sem margem  
- **codigoNCM** – opcional  
- **nome base dos arquivos de saída** – opcional (used como prefixo do Excel e do HTML)

Caso um campo seja deixado em branco, ele **não é enviado** à API (não filtra).

O intervalo de datas (últimos 365 dias) é preenchido automaticamente pelo backend.

---

## 5. Como rodar localmente

### 5.1. Requisitos

- **Python 3.10+** (recomendado 3.10 ou 3.11)
- Acesso à internet (para consumir a API do PNCP)

### 5.2. Instalação

Clone o repositório:

```bash
git clone https://github.com/SEU-USUARIO/pncp-pesquisa-precos.git
cd pncp-pesquisa-precos
