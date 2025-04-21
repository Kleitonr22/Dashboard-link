# --------------------------------------------------------------------------
#                             IMPORTAÇÕES
# --------------------------------------------------------------------------
import streamlit as st
import lxml.etree as ET
import pandas as pd
import os
import traceback
from io import BytesIO  # Para download do Excel

# --------------------------------------------------------------------------
#                       CONFIGURAÇÃO DA PÁGINA STREAMLIT
# --------------------------------------------------------------------------
st.set_page_config(layout="wide", page_title="Dashboard XML Vendas/Compras")
st.title("📊 Dashboard de Análise de XML (Vendas e Compras)")

# --------------------------------------------------------------------------
#                       FUNÇÕES DE PARSING XML
# --------------------------------------------------------------------------


def get_xml_text(element, xpath, namespaces):
    """Função auxiliar para obter texto de um elemento XML, tratando None."""
    if element is None:
        return None
    found = element.find(xpath, namespaces)
    return found.text if found is not None else None


def parse_xml_base(xml_file, tipo_doc):
    """Função base para analisar um arquivo XML (NF-e/NFC-e) e extrair itens."""
    try:
        tree = ET.parse(xml_file)
        root = tree.getroot()
        # Namespace padrão NF-e/NFC-e
        ns = {'ns': 'http://www.portalfiscal.inf.br/nfe'}

        # --- Extrair informações da TAG <ide> (Data/Hora) ---
        ide_element = root.find('.//ns:ide', namespaces=ns)
        dh_emi = None
        if ide_element is not None:
            dh_emi = get_xml_text(ide_element, './ns:dhEmi', ns)
        # Fallback simples se namespace falhar (menos comum)
        elif root.find('.//ide') is not None:
            ide_element = root.find('.//ide')
            dh_emi = get_xml_text(ide_element, './dhEmi', None)

        # --- Extrair informações dos Itens (<det>) ---
        itens = []
        det_elements = root.findall('.//ns:det', namespaces=ns)
        if not det_elements:  # Tenta sem namespace se não encontrar
            det_elements = root.findall('.//det')

        for det in det_elements:
            item = {}
            item['arquivo_origem'] = os.path.basename(xml_file)
            item['data_hora_emissao'] = dh_emi
            item['tipo_documento'] = tipo_doc  # 'Venda' ou 'Compra'

            # Informações do Produto (<prod>)
            prod_element = det.find('./ns:prod', namespaces=ns)
            if prod_element is None:  # Tenta sem namespace
                prod_element = det.find('./prod')
            if prod_element is None:
                continue  # Pula item se não achar <prod>

            # Usa a função auxiliar get_xml_text
            item['codigo_produto'] = get_xml_text(
                prod_element, './ns:cProd', ns) or get_xml_text(prod_element, './cProd', None)
            item['nome_produto'] = get_xml_text(
                prod_element, './ns:xProd', ns) or get_xml_text(prod_element, './xProd', None)
            item['NCM'] = get_xml_text(
                prod_element, './ns:NCM', ns) or get_xml_text(prod_element, './NCM', None)
            item['CFOP'] = get_xml_text(
                prod_element, './ns:CFOP', ns) or get_xml_text(prod_element, './CFOP', None)
            item['codigo_barras'] = get_xml_text(
                prod_element, './ns:cEAN', ns) or get_xml_text(prod_element, './cEAN', None)

            # Quantidade e Valor Unitário
            qcom_text = get_xml_text(
                prod_element, './ns:qCom', ns) or get_xml_text(prod_element, './qCom', None)
            vuncom_text = get_xml_text(
                prod_element, './ns:vUnCom', ns) or get_xml_text(prod_element, './vUnCom', None)

            try:
                item['quantidade'] = float(qcom_text) if qcom_text else 0.0
            except (ValueError, TypeError):
                item['quantidade'] = 0.0
            try:
                item['valor_unitario'] = float(
                    vuncom_text) if vuncom_text else 0.0
            except (ValueError, TypeError):
                item['valor_unitario'] = 0.0

            item['valor_total_item'] = item['quantidade'] * \
                item['valor_unitario']

            # Verifica se informações essenciais foram extraídas
            # Pelo menos um identificador
            if item.get('codigo_produto') or item.get('nome_produto'):
                itens.append(item)

        return itens

    except ET.XMLSyntaxError as e:
        # Erros de sintaxe são comuns, evita poluir demais a interface
        # print(f"Sintaxe XML inválida: {os.path.basename(xml_file)} - {e}") # Log no console
        return []
    except Exception as e:
        # print(f"Erro ao processar XML: {os.path.basename(xml_file)} - {e}") # Log no console
        # traceback.print_exc() # Para depuração profunda
        return []

# --------------------------------------------------------------------------
#                       FUNÇÃO PARA PROCESSAR DIRETÓRIO
# --------------------------------------------------------------------------


@st.cache_data  # Cacheia o resultado para performance
def processar_diretorio_xml(diretorio, tipo_doc):
    """Processa todos os arquivos XML em um diretório e retorna um DataFrame."""
    all_data = []
    if not diretorio or not os.path.isdir(diretorio):
        st.error(
            f"Erro: Diretório de {tipo_doc} não encontrado ou inválido: '{diretorio}'")
        return pd.DataFrame()

    st.write(f"Processando XMLs de {tipo_doc} em: `{diretorio}`...")
    try:
        files = [f for f in os.listdir(
            diretorio) if f.lower().endswith(".xml")]
    except FileNotFoundError:
        st.error(
            f"Erro: Diretório de {tipo_doc} não encontrado ao listar arquivos: '{diretorio}'")
        return pd.DataFrame()
    except Exception as e:
        st.error(f"Erro inesperado ao acessar o diretório {tipo_doc}: {e}")
        return pd.DataFrame()

    progress_bar = st.progress(0)
    total_files = len(files)
    processed_count = 0
    error_files = []

    if total_files == 0:
        st.warning(
            f"Nenhum arquivo .xml encontrado no diretório de {tipo_doc}: `{diretorio}`")
        return pd.DataFrame()

    for i, filename in enumerate(files):
        filepath = os.path.join(diretorio, filename)
        try:
            data = parse_xml_base(filepath, tipo_doc)
            if data:
                all_data.extend(data)
                processed_count += 1
        except Exception as e:
            error_files.append(filename)
            # print(f"Erro crítico ao processar o arquivo {filename}: {e}") # Log console
        progress_bar.progress((i + 1) / total_files)

    if error_files:
        st.warning(
            f"{len(error_files)} arquivo(s) XML de {tipo_doc} não puderam ser processados completamente (verifique o console para detalhes se o log estiver ativo). Arquivos: {', '.join(error_files[:5])}{'...' if len(error_files) > 5 else ''}")

    if not all_data:
        st.warning(
            f"Nenhum dado de item válido encontrado nos XMLs de {tipo_doc} processados com sucesso.")
        return pd.DataFrame()

    df = pd.DataFrame(all_data)
    # Conversões de tipo após criar o DataFrame completo
    df['quantidade'] = pd.to_numeric(
        df['quantidade'], errors='coerce').fillna(0.0)
    df['valor_unitario'] = pd.to_numeric(
        df['valor_unitario'], errors='coerce').fillna(0.0)
    df['valor_total_item'] = pd.to_numeric(
        df['valor_total_item'], errors='coerce').fillna(0.0)
    # Tenta converter data/hora, mas não falha se não conseguir
    if 'data_hora_emissao' in df.columns:
        df['data_hora_emissao'] = pd.to_datetime(
            df['data_hora_emissao'], errors='coerce')

    st.success(
        f"Processamento de {tipo_doc} concluído! {len(df)} itens extraídos de {processed_count} arquivos (de {total_files} encontrados).")
    return df

# --------------------------------------------------------------------------
#                       INTERFACE STREAMLIT E LÓGICA PRINCIPAL
# --------------------------------------------------------------------------


# --- CAMINHOS FIXOS DEFINIDOS AQUI ---
# Use raw strings (r"...") para caminhos do Windows
DIR_VENDAS_FIXO = r"C:\Users\kleit\Desktop\Projeto 3\21 - Sandro\XML Saidas"
DIR_COMPRAS_FIXO = r"C:\Users\kleit\Desktop\Projeto 3\21 - Sandro\XML Entrada"

# --- Barra Lateral ---
st.sidebar.header("Controles e Informações")
st.sidebar.markdown("**Diretórios de XML Configurados:**")
st.sidebar.info(f"**Vendas:** `{DIR_VENDAS_FIXO}`")
st.sidebar.info(f"**Compras:** `{DIR_COMPRAS_FIXO}`")

# Botão para iniciar análise na barra lateral
if st.sidebar.button("Analisar XMLs", key="analisar"):

    # --- Processamento ---
    df_vendas_detalhe = pd.DataFrame()
    df_compras_detalhe = pd.DataFrame()
    processou_vendas = False
    processou_compras = False

    # Processa Vendas
    if os.path.isdir(DIR_VENDAS_FIXO):  # Verifica se o diretório existe antes de processar
        with st.spinner('Processando XMLs de Vendas...'):
            df_vendas_detalhe = processar_diretorio_xml(
                DIR_VENDAS_FIXO, "Venda")
            if not df_vendas_detalhe.empty:
                processou_vendas = True
    else:
        st.sidebar.error(
            f"Diretório de Vendas não encontrado: {DIR_VENDAS_FIXO}")

    # Processa Compras
    # Verifica se o diretório existe antes de processar
    if os.path.isdir(DIR_COMPRAS_FIXO):
        with st.spinner('Processando XMLs de Compras...'):
            df_compras_detalhe = processar_diretorio_xml(
                DIR_COMPRAS_FIXO, "Compra")
            if not df_compras_detalhe.empty:
                processou_compras = True
    else:
        st.sidebar.error(
            f"Diretório de Compras não encontrado: {DIR_COMPRAS_FIXO}")

    # --- Análise e Exibição dos Resultados ---
    if not processou_vendas and not processou_compras:
        st.error(
            "Nenhum dado válido foi processado. Verifique os diretórios e os arquivos XML nos caminhos configurados.")
    else:
        st.header("📈 Resumo Geral")

        # --- Função de Resumo ---
        def criar_resumo_produto(df, tipo):
            if df.empty or 'codigo_produto' not in df.columns:  # Verifica se a coluna chave existe
                return pd.DataFrame()

            # Trata NaNs no código do produto antes de agrupar, se necessário
            df_clean = df.dropna(subset=['codigo_produto'])
            if df_clean.empty:
                return pd.DataFrame()

            agg_funcs = {
                'nome_produto': 'first',  # Pega o primeiro nome encontrado para o código
                'quantidade': 'sum',
                'valor_total_item': 'sum',
                # CFOP mais comum
                'CFOP': lambda x: x.mode()[0] if not x.mode().empty else None,
                'NCM': 'first',
                'valor_unitario': 'mean'  # Valor unitário médio
            }
            try:
                # Agrupa por código do produto (mais confiável)
                df_resumo = df_clean.groupby(
                    'codigo_produto', as_index=False).agg(agg_funcs)

                # Renomeia colunas para clareza
                df_resumo = df_resumo.rename(columns={
                    'quantidade': f'qtd_total_{tipo}',
                    'valor_total_item': f'valor_total_{tipo}',
                    'CFOP': f'CFOP_predominante_{tipo}',
                    'valor_unitario': f'valor_unitario_medio_{tipo}'
                })

                # Ordena pela quantidade
                if f'qtd_total_{tipo}' in df_resumo.columns:
                    df_resumo = df_resumo.sort_values(
                        by=f'qtd_total_{tipo}', ascending=False)
                return df_resumo
            except Exception as e:
                st.error(f"Erro ao criar resumo de {tipo}: {e}")
                return pd.DataFrame()

        df_vendas_resumo = criar_resumo_produto(df_vendas_detalhe, 'vendida')
        df_compras_resumo = criar_resumo_produto(
            df_compras_detalhe, 'comprada')

        # --- Cálculos Específicos (CFOPs) ---
        cfop_vendas_5102 = len(df_vendas_detalhe[df_vendas_detalhe['CFOP'] == '5102']
                               ) if processou_vendas and 'CFOP' in df_vendas_detalhe else 0
        cfop_vendas_5405 = len(df_vendas_detalhe[df_vendas_detalhe['CFOP'] == '5405']
                               ) if processou_vendas and 'CFOP' in df_vendas_detalhe else 0
        cfop_compras_1102_ou_2102 = len(df_compras_detalhe[df_compras_detalhe['CFOP'].isin(
            ['1102', '2102'])]) if processou_compras and 'CFOP' in df_compras_detalhe else 0
        cfop_compras_1403_ou_2403 = len(df_compras_detalhe[df_compras_detalhe['CFOP'].isin(
            ['1403', '2403'])]) if processou_compras and 'CFOP' in df_compras_detalhe else 0

        # --- Exibição dos KPIs ---
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("Vendas (CFOP 5102)", cfop_vendas_5102)
            st.metric("Compras (CFOP 1102/2102)", cfop_compras_1102_ou_2102)
        with col2:
            st.metric("Vendas (CFOP 5405 ST)", cfop_vendas_5405)
            st.metric("Compras (CFOP 1403/2403 ST)", cfop_compras_1403_ou_2403)
        with col3:
            st.metric("Total Linhas Itens Vendidos", len(
                df_vendas_detalhe) if processou_vendas else 0)
            st.metric("Total Linhas Itens Comprados", len(
                df_compras_detalhe) if processou_compras else 0)
        with col4:
            total_vendido = df_vendas_resumo[f'valor_total_vendida'].sum(
            ) if processou_vendas and not df_vendas_resumo.empty and f'valor_total_vendida' in df_vendas_resumo else 0
            total_comprado = df_compras_resumo[f'valor_total_comprada'].sum(
            ) if processou_compras and not df_compras_resumo.empty and f'valor_total_comprada' in df_compras_resumo else 0
            st.metric("Valor Total Vendido", f"R$ {total_vendido:,.2f}")
            st.metric("Valor Total Comprado", f"R$ {total_comprado:,.2f}")

        # --- Exibição dos Top 20 ---
        st.header("🏆 Top 20 Produtos")
        col_v, col_c = st.columns(2)

        with col_v:
            st.subheader("Mais Vendidos (por Quantidade)")
            if processou_vendas and not df_vendas_resumo.empty:
                top_20_vendas = df_vendas_resumo.head(20)
                # Seleciona colunas que existem no dataframe de resumo
                cols_venda_display = [col for col in ['codigo_produto', 'nome_produto', 'qtd_total_vendida',
                                                      'valor_total_vendida', f'CFOP_predominante_vendida'] if col in top_20_vendas.columns]
                st.dataframe(
                    top_20_vendas[cols_venda_display], use_container_width=True)
                # Grafico
                if 'qtd_total_vendida' in top_20_vendas.columns and 'nome_produto' in top_20_vendas.columns:
                    st.bar_chart(top_20_vendas.set_index(
                        'nome_produto')[['qtd_total_vendida']])
            else:
                st.info("Sem dados de resumo de vendas para exibir.")

        with col_c:
            st.subheader("Mais Comprados (por Quantidade)")
            if processou_compras and not df_compras_resumo.empty:
                top_20_compras = df_compras_resumo.head(20)
                # Seleciona colunas que existem no dataframe de resumo
                cols_compra_display = [col for col in ['codigo_produto', 'nome_produto', 'qtd_total_comprada',
                                                       'valor_total_comprada', f'CFOP_predominante_comprada'] if col in top_20_compras.columns]
                st.dataframe(
                    top_20_compras[cols_compra_display], use_container_width=True)
                # Grafico
                if 'qtd_total_comprada' in top_20_compras.columns and 'nome_produto' in top_20_compras.columns:
                    st.bar_chart(top_20_compras.set_index(
                        'nome_produto')[['qtd_total_comprada']])
            else:
                st.info("Sem dados de resumo de compras para exibir.")

        # --- Seção para Download e Detalhes ---
        st.header("📄 Dados Detalhados e Download")

        # --- Remoção do Timezone ANTES de criar o ExcelWriter ---
        df_vendas_excel = df_vendas_detalhe.copy()
        df_compras_excel = df_compras_detalhe.copy()

        if processou_vendas and 'data_hora_emissao' in df_vendas_excel.columns:
            if pd.api.types.is_datetime64_any_dtype(df_vendas_excel['data_hora_emissao']) and df_vendas_excel['data_hora_emissao'].dt.tz is not None:
                df_vendas_excel['data_hora_emissao'] = df_vendas_excel['data_hora_emissao'].dt.tz_localize(
                    None)

        if processou_compras and 'data_hora_emissao' in df_compras_excel.columns:
            if pd.api.types.is_datetime64_any_dtype(df_compras_excel['data_hora_emissao']) and df_compras_excel['data_hora_emissao'].dt.tz is not None:
                df_compras_excel['data_hora_emissao'] = df_compras_excel['data_hora_emissao'].dt.tz_localize(
                    None)
        # ---------------------------------------------------------

        # Botão de Download
        if processou_vendas or processou_compras:
            try:
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    if processou_vendas:
                        df_vendas_excel.to_excel(
                            writer, sheet_name='Detalhes Vendas', index=False)
                        if not df_vendas_resumo.empty:
                            df_vendas_resumo.to_excel(
                                writer, sheet_name='Resumo Vendas', index=False)
                    if processou_compras:
                        df_compras_excel.to_excel(
                            writer, sheet_name='Detalhes Compras', index=False)
                        if not df_compras_resumo.empty:
                            df_compras_resumo.to_excel(
                                writer, sheet_name='Resumo Compras', index=False)
                output.seek(0)  # Volta para o início do buffer para leitura
                st.download_button(label="📥 Baixar Dados em Excel", data=output,
                                   file_name="analise_xml_vendas_compras.xlsx",
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            except Exception as e:
                st.error(f"Erro ao gerar o arquivo Excel para download: {e}")

        # Abas para visualizar os detalhes (Usa os DFs originais)
        tab_vendas, tab_compras = st.tabs(
            ["Detalhes Vendas", "Detalhes Compras"])
        with tab_vendas:
            st.subheader("Todos os Itens de Venda Extraídos")
            if processou_vendas:
                st.dataframe(df_vendas_detalhe, use_container_width=True)
            else:
                st.info("Nenhum detalhe de venda para exibir.")
        with tab_compras:
            st.subheader("Todos os Itens de Compra Extraídos")
            if processou_compras:
                st.dataframe(df_compras_detalhe, use_container_width=True)
            else:
                st.info("Nenhum detalhe de compra para exibir.")

# Mensagem exibida antes de clicar no botão
else:
    st.info("Clique em 'Analisar XMLs' na barra lateral para iniciar o processamento.")
    st.write("Os seguintes diretórios serão analisados:")
    st.write(f"- **Vendas:** `{DIR_VENDAS_FIXO}`")
    st.write(f"- **Compras:** `{DIR_COMPRAS_FIXO}`")
    # Verifica e informa se os diretórios existem
    if not os.path.isdir(DIR_VENDAS_FIXO):
        st.warning("Diretório de Vendas não encontrado!")
    if not os.path.isdir(DIR_COMPRAS_FIXO):
        st.warning("Diretório de Compras não encontrado!")
        
