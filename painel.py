import streamlit as st
import pandas as pd
import plotly.express as px
import re

# Configuração da página do Streamlit para modo widescreen
st.set_page_config(page_title="Dashboard Financeiro", page_icon="📊", layout="wide")

# Título principal do Dashboard
st.title("📊 Dashboard de Controle Financeiro Capela")
st.markdown("Análise interativa dos lançamentos consolidados do Caixa, Banco Sicredi e Caixa Interno.")

def limpar_descricao_para_ranking(texto):
    """
    Remove datas, meses por extenso e unifica termos semelhantes (como Dízimo, Coleta, etc.)
    para agrupar e acumular os valores corretamente no ano inteiro.
    """
    if not isinstance(texto, str):
        return ""
    
    t = texto.upper()
    
    # --- REGRAS DE UNIFICAÇÃO DE TERMOS CHAVE ---
    if 'DÍZIMO' in t or 'DIZIMO' in t:
        return 'DÍZIMO'
    elif 'COLETA' in t:
        return 'COLETA'
    elif 'EVENTO PASTEL' in t or 'VENDAS PASTEL' in t or 'PASTEL' in t:
        return 'EVENTO / VENDAS DE PASTEL'
    # ---------------------------------------------
    
    # Remove padrões de data como 01/2026, 31/12/2025, 09/2026, etc.
    t = re.sub(r'\b\d{2}/\d{2}/\d{4}\b', '', t)
    t = re.sub(r'\b\d{2}/\d{4}\b', '', t)
    t = re.sub(r'\b\d{1,2}/\d{2}\b', '', t)
    
    # Remove meses por extenso comuns se vierem acompanhados
    meses = ['JANEIRO', 'FEVEREIRO', 'MARÇO', 'MARCO', 'ABRIL', 'MAIO', 'JUNHO', 
             'JULHO', 'AGOSTO', 'SETEMBRO', 'OUTUBRO', 'NOVEMBRO', 'DEZEMBRO']
    for m in meses:
        t = t.replace(m, '')
        
    # Remove termos corriqueiros de transação bancária para deixar o nome mais limpo
    termos_comuns = ['RECEBIMENTO', 'PAGAMENTO', 'REC', 'PAG', 'PIX', 'DEB', 'DEBITO', 'TED', 'DOC']
    for termo in termos_comuns:
        t = re.sub(rf'\b{termo}\b', '', t)
        
    # Remove espaços extras que possam ter ficado nas pontas ou duplicados
    t = re.sub(r'\s+', ' ', t).strip()
    
    return t if t else texto.upper()

def carregar_dados():
    try:
        df = pd.read_excel('Controle_Financeiro.xlsx')
        
        # Converte a data considerando o formato brasileiro (dia primeiro: dayfirst=True)
        df['Data_Temp'] = pd.to_datetime(df['Data'], errors='coerce', dayfirst=True)
        
        # Remove linhas onde a data é nula (exclui cabeçalhos vazios e a linha de total geral)
        df = df.dropna(subset=['Data_Temp'])
        
        # Substitui a data original pela data limpa
        df['Data'] = df['Data_Temp']
        df = df.drop(columns=['Data_Temp'])
        
        # Tratamento seguro de colunas de texto
        df['Origem'] = df['Origem'].astype(str)
        df['Histórico / Descrição'] = df['Histórico / Descrição'].fillna('').astype(str)
        
        # Identifica se a linha é saldo anterior ou reserva anterior
        historico_upper = df['Histórico / Descrição'].str.upper()
        df['Is_Saldo_Inicial'] = historico_upper.str.contains(
            'SALDO ANTERIOR|RESERVA MÊS ANTERIOR|RESERVA MES ANTERIOR', na=False
        )
        
        # Cria coluna com a descrição limpa (sem datas e unificada) para os rankings anuais
        df['Histórico_Limpo'] = df['Histórico / Descrição'].apply(limpar_descricao_para_ranking)
        
        # Garante o Mês/Ano preenchido corretamente como string pura extraído da data real
        df['Mês/Ano'] = df['Data'].dt.strftime('%m/%Y').astype(str)
        
        # Formata a data para string limpa sem as horas (DD/MM/AAAA)
        df['Data_Formatada'] = df['Data'].dt.strftime('%d/%m/%Y').fillna('')
        
        # Tratamento de valores numéricos para cálculos e exibição
        df['Entrada (R$)'] = pd.to_numeric(df['Entrada (R$)'], errors='coerce').fillna(0).round(2)
        df['Saída (R$)'] = pd.to_numeric(df['Saída (R$)'], errors='coerce').fillna(0).round(2)
        
        return df
    except Exception as e:
        st.error(f"Erro ao carregar os dados: {e}")
        return None

df_original = carregar_dados()

if df_original is None or df_original.empty:
    st.warning("⚠️ O arquivo `Controle_Financeiro.xlsx` está vazio ou não pôde ser lido.")
else:
    # --- BARRA LATERAL (FILTROS) ---
    st.sidebar.header("🔍 Filtros de Análise")
    
    # Filtro de Tipo de Movimentação (Entradas / Saídas / Todos)
    tipo_movimento = st.sidebar.selectbox(
        "Filtrar por Tipo",
        options=["Todos", "Somente Entradas", "Somente Saídas"],
        index=0
    )
    
    # Filtro de Origem
    origens_disponiveis = sorted(df_original['Origem'].unique().tolist())
    origem_selecionada = st.sidebar.multiselect("Selecione a Origem", options=origens_disponiveis, default=origens_disponiveis)
    
    # Filtro de Mês/Ano
    meses_disponiveis = sorted(df_original['Mês/Ano'].astype(str).unique().tolist())
    mes_selecionado = st.sidebar.multiselect("Selecione o Mês/Ano", options=meses_disponiveis, default=meses_disponiveis)
    
    # Aplicação dos Filtros Base no DataFrame (Origem e Mês/Ano)
    df_filtrado = df_original[
        (df_original['Origem'].isin(origem_selecionada)) & 
        (df_original['Mês/Ano'].isin(mes_selecionado))
    ]
    
    # Aplicação do filtro de tipo de movimento (Entradas/Saídas) em todo o painel
    if tipo_movimento == "Somente Entradas":
        df_painel = df_filtrado[df_filtrado['Entrada (R$)'] > 0]
    elif tipo_movimento == "Somente Saídas":
        df_painel = df_filtrado[df_filtrado['Saída (R$)'] > 0]
    else:
        df_painel = df_filtrado
    
    if df_filtrado.empty:
        st.warning("⚠️ Nenhum registro encontrado para os filtros selecionados.")
    else:
        # --- SEPARAÇÃO PARA OS GRÁFICOS (EXCLUI SALDOS INICIAIS DO FLUXO DO PERÍODO) ---
        df_graficos = df_painel[~df_painel['Is_Saldo_Inicial']]
        
        # --- CARDS DE RESUMO (KPIs) ---
        # 1. Entradas e Saídas operacionais puras (exclui saldo anterior para não inflar receitas/despesas)
        total_entradas_operacionais = df_graficos['Entrada (R$)'].sum()
        total_saidas_operacionais = df_graficos['Saída (R$)'].sum()
        
        # 2. Saldo Inicial considerado para o fechamento correto (considera o que veio de dezembro/janeiro)
        saldo_inicial_acumulado = df_filtrado[df_filtrado['Is_Saldo_Inicial']]['Entrada (R$)'].sum() - df_filtrado[df_filtrado['Is_Saldo_Inicial']]['Saída (R$)'].sum()
        
        # 3. Saldo Líquido Final (Saldo Inicial + Entradas Operacionais - Saídas Operacionais)
        saldo_liquido_final = saldo_inicial_acumulado + total_entradas_operacionais - total_saidas_operacionais
        
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric(label="📥 Total de Entradas", value=f"R$ {total_entradas_operacionais:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
        with col2:
            st.metric(label="📤 Total de Saídas", value=f"R$ {total_saidas_operacionais:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
        with col3:
            st.metric(label="💰 Saldo Líquido (Acumulado)", value=f"R$ {saldo_liquido_final:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."),
                      delta=f"R$ {saldo_liquido_final:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
            
        st.markdown("---")
        
        # --- EXIBIÇÃO DINÂMICA CONFORME O FILTRO SELECIONADO ---
        
        # Se o usuário escolheu "Todos" ou "Somente Entradas", exibe a seção de Entradas
        if tipo_movimento in ["Todos", "Somente Entradas"]:
            st.subheader("📥 Análise de Entradas")
            col_ent1, col_ent2 = st.columns(2)
            
            with col_ent1:
                st.markdown("##### 📈 Evolução de Entradas por Data")
                if not df_graficos.empty and df_graficos['Entrada (R$)'].sum() > 0:
                    df_entradas_tempo = df_graficos.groupby('Data', as_index=False)['Entrada (R$)'].sum()
                    fig_tempo_ent = px.line(df_entradas_tempo, x='Data', y='Entrada (R$)', markers=True, 
                                            title="Receitas ao longo do tempo", template="plotly_white")
                    st.plotly_chart(fig_tempo_ent, use_container_width=True)
                else:
                    st.info("Sem dados de entrada para exibir no gráfico temporal.")
                    
            with col_ent2:
                st.markdown("##### ⚖️ Distribuição de Entradas por Origem")
                if not df_graficos.empty and df_graficos['Entrada (R$)'].sum() > 0:
                    df_origem_ent = df_graficos.groupby('Origem', as_index=False)['Entrada (R$)'].sum()
                    fig_origem_ent = px.pie(df_origem_ent, names='Origem', values='Entrada (R$)', hole=0.4, 
                                            title="Participação das receitas por conta", template="plotly_white")
                    st.plotly_chart(fig_origem_ent, use_container_width=True)
                else:
                    st.info("Sem dados de entrada para exibir por origem.")
                    
            # Ranking de Maiores Entradas
            st.markdown("##### 🏆 Top Maiores Receitas Acumuladas (Ano Inteiro)")
            if not df_graficos.empty and df_graficos['Entrada (R$)'].sum() > 0:
                df_receitas_top = df_graficos[df_graficos['Entrada (R$)'] > 0].groupby('Histórico_Limpo', as_index=False)['Entrada (R$)'].sum()
                df_receitas_top = df_receitas_top.sort_values(by='Entrada (R$)', ascending=False).head(10)
                
                if not df_receitas_top.empty:
                    fig_ranking_ent = px.bar(df_receitas_top, x='Entrada (R$)', y='Histórico_Limpo', orientation='h',
                                             title="Top 10 Receitas Acumuladas", template="plotly_white",
                                             text_auto='.2s')
                    fig_ranking_ent.update_layout(yaxis={'categoryorder':'total ascending'}, yaxis_title="Histórico / Descrição")
                    st.plotly_chart(fig_ranking_ent, use_container_width=True)
            else:
                st.info("Não há registros de entrada para montar o ranking.")

            if tipo_movimento == "Todos":
                st.markdown("---")
        
        # Se o usuário escolheu "Todos" ou "Somente Saídas", exibe a seção de Saídas
        if tipo_movimento in ["Todos", "Somente Saídas"]:
            st.subheader("📤 Análise de Saídas (Despesas)")
            col_graf1, col_graf2 = st.columns(2)
            
            with col_graf1:
                st.markdown("##### 📈 Evolução de Saídas por Data")
                if not df_graficos.empty and df_graficos['Saída (R$)'].sum() > 0:
                    df_saidas_tempo = df_graficos.groupby('Data', as_index=False)['Saída (R$)'].sum()
                    fig_tempo = px.line(df_saidas_tempo, x='Data', y='Saída (R$)', markers=True, 
                                        title="Despesas ao longo do tempo", template="plotly_white")
                    st.plotly_chart(fig_tempo, use_container_width=True)
                else:
                    st.info("Sem dados de saída para exibir no gráfico temporal.")
                    
            with col_graf2:
                st.markdown("##### ⚖️ Distribuição de Saídas por Origem")
                if not df_graficos.empty and df_graficos['Saída (R$)'].sum() > 0:
                    df_origem = df_graficos.groupby('Origem', as_index=False)['Saída (R$)'].sum()
                    fig_origem = px.pie(df_origem, names='Origem', values='Saída (R$)', hole=0.4, 
                                        title="Participação das despesas por conta", template="plotly_white")
                    st.plotly_chart(fig_origem, use_container_width=True)
                else:
                    st.info("Sem dados de saída para exibir por origem.")
                
            # Ranking de Maiores Gastos
            st.markdown("##### 🏆 Top Maiores Gastos Acumulados (Ano Inteiro)")
            if not df_graficos.empty and df_graficos['Saída (R$)'].sum() > 0:
                df_gastos = df_graficos[df_graficos['Saída (R$)'] > 0].groupby('Histórico_Limpo', as_index=False)['Saída (R$)'].sum()
                df_gastos = df_gastos.sort_values(by='Saída (R$)', ascending=False).head(10)
                
                if not df_gastos.empty:
                    fig_ranking = px.bar(df_gastos, x='Saída (R$)', y='Histórico_Limpo', orientation='h',
                                         title="Top 10 Gastos Acumulados", template="plotly_white",
                                         text_auto='.2s')
                    fig_ranking.update_layout(yaxis={'categoryorder':'total ascending'}, yaxis_title="Histórico / Descrição")
                    st.plotly_chart(fig_ranking, use_container_width=True)
            else:
                st.info("Não há registros de saída para montar o ranking.")
            
        st.markdown("---")
        
        # --- TABELA DE DETALHES ---
        with st.expander("📋 Visualizar Tabela de Lançamentos Filtrados"):
            df_tabela = df_painel.copy()
            df_tabela['Data'] = df_tabela['Data_Formatada']
            df_tabela = df_tabela.drop(columns=['Data_Formatada', 'Histórico_Limpo', 'Is_Saldo_Inicial'])
            
            st.dataframe(
                df_tabela,
                use_container_width=True,
                column_config={
                    "Entrada (R$)": st.column_config.NumberColumn(
                        "Entrada (R$)",
                        format="R$ %.2f"
                    ),
                    "Saída (R$)": st.column_config.NumberColumn(
                        "Saída (R$)",
                        format="R$ %.2f"
                    )
                }
            )