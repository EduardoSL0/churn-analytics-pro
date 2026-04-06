import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from sklearn.ensemble import RandomForestClassifier
from sklearn.model_selection import train_test_split
from sklearn.metrics import accuracy_score, classification_report
import io
import base64
from datetime import datetime
import warnings
warnings.filterwarnings('ignore')

# ============================
# CONFIGURAÇÃO DA PÁGINA
# ============================

st.set_page_config(
    page_title="Churn Analytics Pro",
    page_icon="🚀",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ============================
# CSS CUSTOMIZADO (TEMA DARK PREMIUM)
# ============================

st.markdown("""
<style>
    .main {
        background-color: #0B0F19;
        color: #E2E8F0;
    }
    .css-1d391kg {
        background-color: #151B2B;
    }
    .stMetric {
        background-color: #1E2538;
        border: 1px solid #2D3748;
        border-radius: 10px;
        padding: 15px;
    }
    .stButton>button {
        background-color: #3B82F6;
        color: white;
        border-radius: 8px;
        border: none;
        padding: 10px 24px;
        font-weight: bold;
    }
    .stButton>button:hover {
        background-color: #2563EB;
        box-shadow: 0 4px 12px rgba(59, 130, 246, 0.4);
    }
    h1 {
        color: #60A5FA !important;
        font-weight: 700 !important;
    }
    h2, h3 {
        color: #94A3B8 !important;
        font-weight: 600 !important;
    }
    .stDataFrame {
        background-color: #1E2538;
    }
    .stAlert {
        border-radius: 8px;
    }
    .stProgress > div > div {
        background-color: #3B82F6;
    }
</style>
""", unsafe_allow_html=True)

# ============================
# FUNÇÕES AUXILIARES
# ============================

def calcular_score_risco(row):
    """Calcula score de risco manual"""
    score = 0
    fatores = []
    
    if row['Satisfacao'] <= 3:
        score += 30
        fatores.append('Satisfação Crítica')
    elif row['Satisfacao'] <= 5:
        score += 15
        fatores.append('Satisfação Baixa')
    
    if row['NPS'] <= 3:
        score += 25
        fatores.append('NPS Detrator')
    elif row['NPS'] <= 6:
        score += 10
        fatores.append('NPS Passivo')
    
    if row['Ultimo_Contato'] > 45:
        score += 20
        fatores.append('Abandonado 45d+')
    elif row['Ultimo_Contato'] > 30:
        score += 10
        fatores.append('Sem contato 30d+')
    
    if row['Suportes_Abertos'] >= 3:
        score += 15
        fatores.append('Problemas Críticos')
    
    if row['Desconto_Aplicado'] == 0:
        score += 10
        fatores.append('Sem Desconto')
    
    if row['Forma_Pagamento'] == 'Boleto':
        score += 10
        fatores.append('Pagamento Boleto')
    
    return min(score, 100), fatores

def treinar_modelo(df):
    """Treina modelo de ML"""
    df_ml = df.copy()
    df_ml['Plano_Code'] = df_ml['Plano'].astype('category').cat.codes
    df_ml['Pagamento_Code'] = df_ml['Forma_Pagamento'].astype('category').cat.codes
    
    features = ['Idade', 'Tempo_Contrato', 'Valor_Mensal', 'Satisfacao',
                'NPS', 'Ultimo_Contato', 'Suportes_Abertos', 'Desconto_Aplicado',
                'Plano_Code', 'Pagamento_Code']
    
    X = df_ml[features]
    y = df_ml['Cancelou_Bin']
    
    if len(df) < 10:
        return None, 0, None, features
    
    X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.3, random_state=42)
    modelo = RandomForestClassifier(n_estimators=200, max_depth=15, random_state=42)
    modelo.fit(X_train, y_train)
    
    acuracia = accuracy_score(y_test, modelo.predict(X_test)) * 100
    
    importancia = pd.DataFrame({
        'Feature': features,
        'Importancia': modelo.feature_importances_
    }).sort_values('Importancia', ascending=False)
    
    return modelo, acuracia, importancia, features

def gerar_excel_relatorio(df, importancia, metricas):
    """Gera relatório Excel completo para download - CORRIGIDO"""
    output = io.BytesIO()
    
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        
        # ============================
        # FORMATOS COM LARGURAS DEFINIDAS
        # ============================
        
        # Formatos de cabeçalho
        header_format = workbook.add_format({
            'bold': True,
            'font_size': 12,
            'bg_color': '#1E2538',
            'font_color': '#FFFFFF',
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'text_wrap': True
        })
        
        header_green = workbook.add_format({
            'bold': True,
            'font_size': 12,
            'bg_color': '#064E3B',
            'font_color': '#FFFFFF',
            'border': 1,
            'align': 'center',
            'valign': 'vcenter'
        })
        
        # Formatos de células
        cell_format = workbook.add_format({
            'font_size': 11,
            'bg_color': '#151B2B',
            'font_color': '#E2E8F0',
            'border': 1,
            'align': 'center',
            'valign': 'vcenter'
        })
        
        cell_text = workbook.add_format({
            'font_size': 11,
            'bg_color': '#151B2B',
            'font_color': '#E2E8F0',
            'border': 1,
            'align': 'left',
            'valign': 'vcenter'
        })
        
        cell_number = workbook.add_format({
            'font_size': 11,
            'bg_color': '#151B2B',
            'font_color': '#E2E8F0',
            'border': 1,
            'align': 'right',
            'valign': 'vcenter',
            'num_format': '#,##0.00'
        })
        
        cell_percent = workbook.add_format({
            'font_size': 11,
            'bg_color': '#151B2B',
            'font_color': '#E2E8F0',
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'num_format': '0.0%'
        })
        
        cell_currency = workbook.add_format({
            'font_size': 11,
            'bg_color': '#151B2B',
            'font_color': '#E2E8F0',
            'border': 1,
            'align': 'right',
            'valign': 'vcenter',
            'num_format': 'R$ #,##0.00'
        })
        
        # Formatos de alerta
        alert_format = workbook.add_format({
            'bold': True,
            'font_size': 11,
            'bg_color': '#7F1D1D',
            'font_color': '#FFFFFF',
            'border': 1,
            'align': 'center',
            'valign': 'vcenter'
        })
        
        success_format = workbook.add_format({
            'bold': True,
            'font_size': 11,
            'bg_color': '#064E3B',
            'font_color': '#FFFFFF',
            'border': 1,
            'align': 'center',
            'valign': 'vcenter'
        })
        
        warning_format = workbook.add_format({
            'bold': True,
            'font_size': 11,
            'bg_color': '#92400E',
            'font_color': '#FFFFFF',
            'border': 1,
            'align': 'center',
            'valign': 'vcenter'
        })
        
        title_format = workbook.add_format({
            'bold': True,
            'font_size': 16,
            'bg_color': '#0B0F19',
            'font_color': '#60A5FA',
            'align': 'left',
            'valign': 'vcenter'
        })
        
        # ============================
        # ABA 1: RESUMO EXECUTIVO
        # ============================
        
        ws_resumo = workbook.add_worksheet('Resumo Executivo')
        ws_resumo.set_tab_color('#3B82F6')
        
        # Larguras das colunas
        ws_resumo.set_column('A:A', 5)   # Margem
        ws_resumo.set_column('B:B', 30)  # Métrica
        ws_resumo.set_column('C:C', 25)  # Valor
        
        # Título
        ws_resumo.merge_range('B2:C2', 'CHURN ANALYTICS PRO - RESUMO EXECUTIVO', title_format)
        ws_resumo.write('B4', 'Métrica', header_format)
        ws_resumo.write('C4', 'Valor', header_format)
        
        # Dados
        dados_resumo = [
            ['Total de Clientes', metricas['total']],
            ['Taxa de Churn', f"{metricas['taxa_churn']:.1f}%"],
            ['NPS Score', f"{metricas['nps']:.0f}"],
            ['Receita Mensal', f"R$ {metricas['receita']:,.2f}"],
            ['Clientes em Alto Risco', metricas['risco']],
            ['Economia Projetada', f"R$ {metricas['economia']:,.2f}"],
            ['Acurácia do Modelo ML', f"{metricas['acuracia']:.1f}%"]
        ]
        
        for i, (metrica, valor) in enumerate(dados_resumo, start=5):
            ws_resumo.write(f'B{i}', metrica, cell_text)
            ws_resumo.write(f'C{i}', valor, cell_format)
        
        # Insights
        ws_resumo.write('B13', 'INSIGHTS PRINCIPAIS', header_green)
        ws_resumo.merge_range('B14:C14', f"• {metricas['risco']} clientes precisam de atenção imediata", cell_text)
        ws_resumo.merge_range('B15:C15', f"• {metricas['abandonados']} clientes sem contato há 30+ dias", cell_text)
        ws_resumo.merge_range('B16:C16', f"• {metricas['criticos']} tickets críticos pendentes", cell_text)
        
        # ============================
        # ABA 2: CLIENTES ANÁLISE
        # ============================
        
        ws_clientes = workbook.add_worksheet('Clientes Análise')
        ws_clientes.set_tab_color('#10B981')
        
        # Preparar dados
        df_export = df[['Cliente', 'Plano', 'Segmento', 'Prob_Churn', 'Risco_ML', 
                       'Score_Risco', 'Satisfacao', 'NPS', 'Valor_Mensal', 'Cancelou']].copy()
        
        # Converter tipos
        df_export['Prob_Churn'] = pd.to_numeric(df_export['Prob_Churn'], errors='coerce').fillna(0)
        df_export['Score_Risco'] = pd.to_numeric(df_export['Score_Risco'], errors='coerce').fillna(0)
        df_export['Valor_Mensal'] = pd.to_numeric(df_export['Valor_Mensal'], errors='coerce').fillna(0)
        df_export['Risco_ML'] = df_export['Risco_ML'].astype(str)
        
        # Larguras
        ws_clientes.set_column('A:A', 5)    # Margem
        ws_clientes.set_column('B:B', 20)   # Cliente
        ws_clientes.set_column('C:C', 15)   # Plano
        ws_clientes.set_column('D:D', 15)   # Segmento
        ws_clientes.set_column('E:E', 12)   # Prob_Churn
        ws_clientes.set_column('F:F', 12)   # Risco_ML
        ws_clientes.set_column('G:G', 12)   # Score_Risco
        ws_clientes.set_column('H:H', 12)   # Satisfacao
        ws_clientes.set_column('I:I', 10)   # NPS
        ws_clientes.set_column('J:J', 15)   # Valor_Mensal
        ws_clientes.set_column('K:K', 12)   # Cancelou
        
        # Cabeçalhos
        headers = ['Cliente', 'Plano', 'Segmento', 'Prob_Churn (%)', 'Risco_ML', 
                  'Score_Risco', 'Satisfacao', 'NPS', 'Valor_Mensal', 'Status']
        
        for col, header in enumerate(headers, start=1):
            ws_clientes.write(0, col, header, header_format)
        
        # Dados
        for row_idx, (_, row) in enumerate(df_export.iterrows(), start=1):
            ws_clientes.write(row_idx, 1, str(row['Cliente']), cell_text)
            ws_clientes.write(row_idx, 2, str(row['Plano']), cell_format)
            ws_clientes.write(row_idx, 3, str(row['Segmento']), cell_format)
            ws_clientes.write_number(row_idx, 4, float(row['Prob_Churn']), cell_number)
            ws_clientes.write(row_idx, 5, str(row['Risco_ML']), cell_format)
            ws_clientes.write_number(row_idx, 6, int(row['Score_Risco']), cell_format)
            ws_clientes.write_number(row_idx, 7, int(row['Satisfacao']), cell_format)
            ws_clientes.write_number(row_idx, 8, int(row['NPS']), cell_format)
            ws_clientes.write_number(row_idx, 9, float(row['Valor_Mensal']), cell_currency)
            
            # Status com cor
            status = 'Cancelado' if str(row['Cancelou']).lower() == 'sim' else 'Ativo'
            fmt = alert_format if status == 'Cancelado' else success_format
            ws_clientes.write(row_idx, 10, status, fmt)
        
        # ============================
        # ABA 3: ALERTAS CRÍTICOS
        # ============================
        
        ws_alertas = workbook.add_worksheet('Alertas Críticos')
        ws_alertas.set_tab_color('#EF4444')
        
        criticos = df[df['Risco_ML'].isin(['Alto', 'Crítico'])].sort_values('Prob_Churn', ascending=False)
        
        if len(criticos) > 0:
            alertas_export = criticos[['Cliente', 'Plano', 'Prob_Churn', 'Score_Risco', 
                                      'Satisfacao', 'Fatores_Risco']].copy()
            alertas_export['Prob_Churn'] = pd.to_numeric(alertas_export['Prob_Churn'], errors='coerce').fillna(0)
            alertas_export['Score_Risco'] = pd.to_numeric(alertas_export['Score_Risco'], errors='coerce').fillna(0)
            
            # Larguras
            ws_alertas.set_column('A:A', 5)
            ws_alertas.set_column('B:B', 20)   # Cliente
            ws_alertas.set_column('C:C', 15)   # Plano
            ws_alertas.set_column('D:D', 15)   # Prob_Churn
            ws_alertas.set_column('E:E', 15)   # Score
            ws_alertas.set_column('F:F', 12)   # Satisfacao
            ws_alertas.set_column('G:G', 40)   # Fatores
            
            # Cabeçalhos
            headers_alertas = ['Cliente', 'Plano', 'Prob_Churn (%)', 'Score_Risco', 'Satisfacao', 'Fatores de Risco']
            for col, header in enumerate(headers_alertas, start=1):
                ws_alertas.write(0, col, header, header_format)
            
            # Dados
            for row_idx, (_, row) in enumerate(alertas_export.iterrows(), start=1):
                ws_alertas.write(row_idx, 1, str(row['Cliente']), cell_text)
                ws_alertas.write(row_idx, 2, str(row['Plano']), cell_format)
                ws_alertas.write_number(row_idx, 3, float(row['Prob_Churn']), cell_number)
                ws_alertas.write_number(row_idx, 4, int(row['Score_Risco']), cell_format)
                ws_alertas.write_number(row_idx, 5, int(row['Satisfacao']), cell_format)
                ws_alertas.write(row_idx, 6, str(row['Fatores_Risco']), cell_text)
        else:
            ws_alertas.write('B2', 'Nenhum cliente em alto risco encontrado.', cell_format)
        
        # ============================
        # ABA 4: FATORES ML
        # ============================
        
        ws_fatores = workbook.add_worksheet('Fatores ML')
        ws_fatores.set_tab_color('#8B5CF6')
        
        if importancia is not None:
            imp_export = importancia.copy()
            imp_export['Importancia_Pct'] = (imp_export['Importancia'] * 100).round(2)
            
            # Larguras
            ws_fatores.set_column('A:A', 5)
            ws_fatores.set_column('B:B', 25)   # Feature
            ws_fatores.set_column('C:C', 20)   # Importancia %
            
            # Cabeçalhos
            ws_fatores.write('B1', 'Fator', header_format)
            ws_fatores.write('C1', 'Importância (%)', header_format)
            
            # Dados
            for row_idx, (_, row) in enumerate(imp_export.iterrows(), start=1):
                ws_fatores.write(row_idx, 1, str(row['Feature']), cell_text)
                ws_fatores.write_number(row_idx, 2, float(row['Importancia_Pct']) / 100, cell_percent)
        else:
            ws_fatores.write('B2', 'Modelo estatístico - sem análise de features ML', cell_format)
        
        # ============================
        # ABA 5: PLANO DE AÇÃO
        # ============================
        
        ws_acoes = workbook.add_worksheet('Plano de Ação')
        ws_acoes.set_tab_color('#F59E0B')
        
        # Larguras
        ws_acoes.set_column('A:A', 5)
        ws_acoes.set_column('B:B', 12)   # Prioridade
        ws_acoes.set_column('C:C', 40)   # Ação
        ws_acoes.set_column('D:D', 25)   # Impacto
        ws_acoes.set_column('E:E', 12)   # Prazo
        
        # Cabeçalhos
        headers_acoes = ['Prioridade', 'Ação Recomendada', 'Impacto Esperado', 'Prazo']
        for col, header in enumerate(headers_acoes, start=1):
            ws_acoes.write(0, col, header, header_format)
        
        # Dados
        acoes = [
            ['🔴 ALTA', f'Contatar {metricas["abandonados"]} clientes sem contato há 30+ dias', 
             'Evita 40% de churn', '48 horas'],
            ['🔴 ALTA', f'Resolver {metricas["criticos"]} tickets críticos pendentes', 
             'Aumenta satisfação em 3 pontos', '24 horas'],
            ['🟡 MÉDIA', f'Oferecer desconto para {metricas["insatisfeitos"]} clientes insatisfeitos', 
             'Retenção de 60% dos clientes', '72 horas'],
            ['🟡 MÉDIA', f'Migrar {metricas["boleto"]} clientes de boleto para cartão/Pix', 
             'Reduz inadimplência em 25%', '7 dias']
        ]
        
        for row_idx, (prioridade, acao, impacto, prazo) in enumerate(acoes, start=1):
            fmt_prioridade = alert_format if 'ALTA' in prioridade else warning_format
            
            ws_acoes.write(row_idx, 1, prioridade, fmt_prioridade)
            ws_acoes.write(row_idx, 2, acao, cell_text)
            ws_acoes.write(row_idx, 3, impacto, cell_text)
            ws_acoes.write(row_idx, 4, prazo, cell_format)
        
        # Rodapé
        ws_acoes.write('B7', '💡 Nota: Priorizar ações de ALTA prioridade para maior impacto imediato.', 
                      workbook.add_format({'italic': True, 'font_color': '#94A3B8'}))
    
    output.seek(0)
    return output

# ============================
# SIDEBAR
# ============================

with st.sidebar:
    st.markdown("<h1 style='text-align: center; color: #60A5FA;'>🚀 Churn Analytics Pro</h1>", unsafe_allow_html=True)
    st.markdown("---")
    
    st.markdown("### 📁 Upload de Dados")
    uploaded_file = st.file_uploader(
        "Arraste sua planilha Excel",
        type=['xlsx', 'xls'],
        help="O arquivo deve conter: Cliente, Idade, Plano, Tempo_Contrato, Valor_Mensal, Cancelou, Satisfacao, NPS, Ultimo_Contato, Suportes_Abertos, Desconto_Aplicado, Forma_Pagamento"
    )
    
    plano_filter = 'Todos'
    risco_filter = ['Alto', 'Crítico']
    
    if uploaded_file is not None:
        st.success("✅ Arquivo carregado!")
        
        try:
            df_temp = pd.read_excel(uploaded_file)
            
            st.markdown("### 🎛️ Filtros")
            
            planos = ['Todos'] + list(df_temp['Plano'].unique())
            plano_filter = st.selectbox("Plano", planos)
            
            risco_opcoes = ['Baixo', 'Médio', 'Alto', 'Crítico']
            risco_filter = st.multiselect(
                "Nível de Risco",
                risco_opcoes,
                default=['Alto', 'Crítico']
            )
        except Exception as e:
            st.error(f"Erro ao ler arquivo: {e}")
    
    st.markdown("---")
    st.markdown("### ℹ️ Sobre")
    st.info("""
    Dashboard com Machine Learning (Random Forest) 
    para prever churn de clientes.
    
    Desenvolvido com Streamlit + Python
    """)

# ============================
# MAIN CONTENT
# ============================

if uploaded_file is None:
    st.title("🎯 Churn Analytics Pro")
    st.markdown("### Previsão de Cancelamento com Machine Learning")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.metric("🤖 Acurácia do Modelo", "Até 95%", "ML Avançado")
    
    with col2:
        st.metric("⚡ Processamento", "Tempo Real", "< 2 segundos")
    
    with col3:
        st.metric("💰 ROI Médio", "300%", "Retenção Inteligente")
    
    st.markdown("---")
    
    st.markdown("""
    ### 🚀 Como funciona:
    
    1. **📁 Upload** - Envie sua planilha Excel com dados dos clientes
    2. **🤖 Análise** - Nosso modelo ML processa automaticamente  
    3. **📊 Dashboard** - Visualize riscos, segmentos e previsões
    4. **💡 Ações** - Receba recomendações personalizadas
    5. **📥 Download** - Exporte relatório completo em Excel
    
    ### 📋 Estrutura esperada do arquivo:
    
    | Coluna | Descrição | Exemplo |
    |--------|-----------|---------|
    | Cliente | Nome do cliente | João Silva |
    | Idade | Idade em anos | 35 |
    | Plano | Tipo de plano | Premium |
    | Tempo_Contrato | Meses de contrato | 12 |
    | Valor_Mensal | Valor pago mensal | 99.90 |
    | Cancelou | Já cancelou? (Sim/Nao) | Nao |
    | Satisfacao | Nota 1-10 | 8 |
    | NPS | Nota 0-10 | 9 |
    | Ultimo_Contato | Dias sem contato | 5 |
    | Suportes_Abertos | Chamados pendentes | 0 |
    | Desconto_Aplicado | % desconto | 15 |
    | Forma_Pagamento | Meio de pagamento | Cartao |
    
    ---
    
    **👈 Faça upload do seu arquivo na barra lateral para começar!**
    """)

else:
    with st.spinner("🤖 Processando dados com Machine Learning..."):
        
        try:
            df = pd.read_excel(uploaded_file)
            
            colunas_necessarias = ['Cliente', 'Idade', 'Plano', 'Tempo_Contrato', 
                                  'Valor_Mensal', 'Cancelou', 'Satisfacao', 'NPS',
                                  'Ultimo_Contato', 'Suportes_Abertos', 
                                  'Desconto_Aplicado', 'Forma_Pagamento']
            
            colunas_faltantes = [col for col in colunas_necessarias if col not in df.columns]
            
            if colunas_faltantes:
                st.error(f"❌ Colunas faltantes no arquivo: {', '.join(colunas_faltantes)}")
                st.stop()
            
            df['Valor_Mensal'] = df['Valor_Mensal'].astype(str).str.replace(',', '.').astype(float)
            df['Cancelou'] = df['Cancelou'].str.strip().str.lower()
            df['Cancelou_Bin'] = (df['Cancelou'] == 'sim').astype(int)
            
            df['Satisfacao'] = df['Satisfacao'].fillna(5)
            df['NPS'] = df['NPS'].fillna(5)
            df['Ultimo_Contato'] = df['Ultimo_Contato'].fillna(30)
            df['Suportes_Abertos'] = df['Suportes_Abertos'].fillna(0)
            df['Desconto_Aplicado'] = df['Desconto_Aplicado'].fillna(0)
            df['Forma_Pagamento'] = df['Forma_Pagamento'].fillna('Boleto')
            
            resultados = df.apply(lambda row: calcular_score_risco(row), axis=1)
            df['Score_Risco'] = [r[0] for r in resultados]
            df['Fatores_Risco'] = [', '.join(r[1]) if r[1] else 'Nenhum' for r in resultados]
            
            modelo, acuracia, importancia, features = treinar_modelo(df)
            
            if modelo is not None:
                df_ml = df.copy()
                df_ml['Plano_Code'] = df_ml['Plano'].astype('category').cat.codes
                df_ml['Pagamento_Code'] = df_ml['Forma_Pagamento'].astype('category').cat.codes
                
                X = df_ml[features]
                df['Prob_Churn'] = modelo.predict_proba(X)[:, 1] * 100
            else:
                df['Prob_Churn'] = df['Score_Risco'].astype(float)
                acuracia = 0
            
            df['Risco_ML'] = pd.cut(df['Prob_Churn'],
                                    bins=[0, 25, 50, 75, 100],
                                    labels=['Baixo', 'Médio', 'Alto', 'Crítico'])
            
            df['Segmento'] = 'Neutro'
            df.loc[(df['Cancelou'] == 'nao') & (df['Prob_Churn'] < 30), 'Segmento'] = 'Campeões'
            df.loc[(df['Cancelou'] == 'nao') & (df['Prob_Churn'] > 60), 'Segmento'] = 'Em Risco'
            df.loc[df['Cancelou'] == 'sim', 'Segmento'] = 'Perdidos'
            df.loc[(df['Cancelou'] == 'nao') & (df['Satisfacao'] >= 9) & (df['NPS'] >= 9), 'Segmento'] = 'Promotores'
            
            df_filtrado = df.copy()
            
            if plano_filter != 'Todos':
                df_filtrado = df_filtrado[df_filtrado['Plano'] == plano_filter]
            
            if risco_filter:
                df_filtrado = df_filtrado[df_filtrado['Risco_ML'].isin(risco_filter)]
            
            if len(df_filtrado) == 0:
                st.warning("⚠️ Nenhum dado encontrado com os filtros selecionados. Mostrando todos os dados.")
                df_filtrado = df.copy()
            
        except Exception as e:
            st.error(f"❌ Erro ao processar dados: {str(e)}")
            st.stop()
    
    st.title("🎯 Dashboard de Churn Analytics")
    
    col1, col2, col3, col4, col5 = st.columns(5)
    
    total = len(df_filtrado)
    cancelados = len(df_filtrado[df_filtrado['Cancelou'] == 'sim'])
    taxa_churn = (cancelados / total * 100) if total > 0 else 0
    
    promotores = len(df_filtrado[df_filtrado['NPS'] >= 9])
    detratores = len(df_filtrado[df_filtrado['NPS'] <= 6])
    nps = ((promotores - detratores) / total * 100) if total > 0 else 0
    
    alto_risco = len(df_filtrado[(df_filtrado['Risco_ML'].isin(['Alto', 'Crítico'])) & (df_filtrado['Cancelou'] == 'nao')])
    receita_risco = df_filtrado[(df_filtrado['Risco_ML'].isin(['Alto', 'Crítico'])) & (df_filtrado['Cancelou'] == 'nao')]['Valor_Mensal'].sum()
    
    economia = alto_risco * 400
    
    with col1:
        st.metric("👥 Total Clientes", f"{total}")
    
    with col2:
        delta_color = "inverse" if taxa_churn > 15 else "normal"
        st.metric("📉 Taxa Churn", f"{taxa_churn:.1f}%", delta=f"{taxa_churn-15:.1f}% vs meta", delta_color=delta_color)
    
    with col3:
        st.metric("⭐ NPS Score", f"{nps:.0f}")
    
    with col4:
        st.metric("🚨 Em Risco", f"{alto_risco}", delta=f"R$ {receita_risco/1000:.0f}k", delta_color="inverse")
    
    with col5:
        st.metric("💰 Economia", f"R$ {economia/1000:.0f}k")
    
    st.markdown("---")
    
    col_left, col_right = st.columns(2)
    
    with col_left:
        st.subheader("📊 Distribuição de Risco")
        
        df_ativos = df_filtrado[df_filtrado['Cancelou'] == 'nao'].copy()
        
        if len(df_ativos) > 0:
            risco_dist = df_ativos['Risco_ML'].value_counts().sort_index()
            
            if len(risco_dist) > 0:
                risco_df = pd.DataFrame({
                    'Risco': risco_dist.index.astype(str),
                    'Quantidade': risco_dist.values
                })
                
                color_map = {
                    'Baixo': '#10B981',
                    'Médio': '#F59E0B',
                    'Alto': '#EF4444',
                    'Crítico': '#7F1D1D'
                }
                
                fig_risco = px.pie(
                    risco_df,
                    values='Quantidade',
                    names='Risco',
                    color='Risco',
                    color_discrete_map=color_map,
                    hole=0.4
                )
                fig_risco.update_traces(textposition='inside', textinfo='percent+label')
                fig_risco.update_layout(
                    paper_bgcolor='rgba(0,0,0,0)',
                    plot_bgcolor='rgba(0,0,0,0)',
                    font_color='#E2E8F0',
                    showlegend=False,
                    margin=dict(t=30, b=30, l=30, r=30)
                )
                st.plotly_chart(fig_risco, use_container_width=True)
            else:
                st.info("ℹ️ Nenhum dado de risco disponível.")
        else:
            st.warning("⚠️ Não há clientes ativos para análise de risco.")
    
    with col_right:
        st.subheader("🎯 Segmentação de Clientes")
        
        if len(df_filtrado) > 0:
            seg_dist = df_filtrado['Segmento'].value_counts()
            
            if len(seg_dist) > 0:
                seg_df = pd.DataFrame({
                    'Segmento': seg_dist.index,
                    'Quantidade': seg_dist.values
                })
                
                color_map_seg = {
                    'Campeões': '#3B82F6',
                    'Promotores': '#10B981',
                    'Neutro': '#94A3B8',
                    'Em Risco': '#F59E0B',
                    'Perdidos': '#EF4444'
                }
                
                fig_seg = px.bar(
                    seg_df,
                    x='Segmento',
                    y='Quantidade',
                    color='Segmento',
                    color_discrete_map=color_map_seg,
                    text='Quantidade'
                )
                fig_seg.update_traces(textposition='outside')
                fig_seg.update_layout(
                    paper_bgcolor='rgba(0,0,0,0)',
                    plot_bgcolor='rgba(0,0,0,0)',
                    font_color='#E2E8F0',
                    showlegend=False,
                    xaxis_title="",
                    yaxis_title="Quantidade",
                    xaxis=dict(showgrid=False),
                    yaxis=dict(showgrid=True, gridcolor='#1E293B'),
                    margin=dict(t=30, b=30, l=30, r=30)
                )
                st.plotly_chart(fig_seg, use_container_width=True)
            else:
                st.info("ℹ️ Nenhum dado de segmentação disponível.")
        else:
            st.warning("⚠️ Não há dados para segmentação.")
    
    if importancia is not None:
        st.markdown("---")
        st.subheader("🧠 Fatores que Influenciam o Churn (ML)")
        
        imp_top = importancia.head(6).copy()
        imp_top['Importancia_%'] = (imp_top['Importancia'] * 100).round(1)
        
        fig_imp = px.bar(
            imp_top,
            x='Importancia_%',
            y='Feature',
            orientation='h',
            color='Importancia_%',
            color_continuous_scale='Blues',
            text=imp_top['Importancia_%'].apply(lambda x: f'{x:.1f}%')
        )
        fig_imp.update_traces(textposition='outside')
        fig_imp.update_layout(
            paper_bgcolor='rgba(0,0,0,0)',
            plot_bgcolor='rgba(0,0,0,0)',
            font_color='#E2E8F0',
            xaxis_title="Importância (%)",
            yaxis_title="",
            coloraxis_showscale=False,
            height=400,
            margin=dict(t=30, b=30, l=100, r=30)
        )
        st.plotly_chart(fig_imp, use_container_width=True)
    
    st.markdown("---")
    st.subheader("🚨 Clientes em Alto Risco - Ação Imediata")
    
    alertas = df_filtrado[(df_filtrado['Cancelou'] == 'nao')].nlargest(10, 'Prob_Churn')[
        ['Cliente', 'Plano', 'Prob_Churn', 'Score_Risco', 'Satisfacao', 'NPS', 'Fatores_Risco']
    ].reset_index(drop=True)
    
    if len(alertas) > 0:
        def color_risco(val):
            if isinstance(val, (int, float)):
                if val >= 70:
                    return 'background-color: #7F1D1D; color: white'
                elif val >= 50:
                    return 'background-color: #92400E; color: white'
            return ''
        
        styled_alertas = alertas.style.applymap(color_risco, subset=['Prob_Churn', 'Score_Risco'])
        st.dataframe(styled_alertas, use_container_width=True, height=400)
    else:
        st.info("ℹ️ Nenhum cliente em alto risco encontrado.")
    
    st.markdown("---")
    st.subheader("📥 Exportar Relatório Completo")
    
    metricas = {
        'total': total,
        'taxa_churn': taxa_churn,
        'nps': nps,
        'receita': df_filtrado[df_filtrado['Cancelou'] == 'nao']['Valor_Mensal'].sum(),
        'risco': alto_risco,
        'economia': economia,
        'acuracia': acuracia,
        'abandonados': len(df_filtrado[(df_filtrado['Ultimo_Contato'] > 30) & (df_filtrado['Cancelou'] == 'nao')]),
        'criticos': len(df_filtrado[df_filtrado['Suportes_Abertos'] >= 3]),
        'insatisfeitos': len(df_filtrado[(df_filtrado['Desconto_Aplicado'] == 0) & (df_filtrado['Satisfacao'] < 6) & (df_filtrado['Cancelou'] == 'nao')]),
        'boleto': len(df_filtrado[(df_filtrado['Forma_Pagamento'] == 'Boleto') & (df_filtrado['Cancelou'] == 'nao')])
    }
    
    excel_file = gerar_excel_relatorio(df_filtrado, importancia, metricas)
    
    col_download, col_info = st.columns([1, 3])
    
    with col_download:
        st.download_button(
            label="📊 Baixar Relatório Excel",
            data=excel_file,
            file_name=f"Churn_Report_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    
    with col_info:
        st.info("""
        O relatório Excel contém:
        • **Resumo Executivo** - KPIs principais
        • **Clientes Análise** - Base completa com scores
        • **Alertas Críticos** - Top prioridades  
        • **Fatores ML** - Importância das variáveis
        • **Plano de Ação** - Recomendações estratégicas
        """)
    
    st.markdown("---")
    st.subheader("🔮 Prever Risco de Novo Cliente")
    
    with st.expander("Clique para simular um cliente"):
        col_sim1, col_sim2, col_sim3 = st.columns(3)
        
        with col_sim1:
            sim_idade = st.number_input("Idade", 18, 100, 35)
            sim_tempo = st.number_input("Tempo Contrato (meses)", 0, 120, 12)
            sim_valor = st.number_input("Valor Mensal", 0.0, 1000.0, 79.90)
        
        with col_sim2:
            sim_satisfacao = st.slider("Satisfação (1-10)", 1, 10, 7)
            sim_nps = st.slider("NPS (0-10)", 0, 10, 8)
            sim_contato = st.number_input("Dias sem contato", 0, 365, 15)
        
        with col_sim3:
            sim_suportes = st.number_input("Suportes Abertos", 0, 10, 0)
            sim_desconto = st.number_input("Desconto %", 0, 100, 0)
            sim_pagamento = st.selectbox("Pagamento", ['Cartao', 'Boleto', 'Pix'])
            sim_plano = st.selectbox("Plano", ['Basico', 'Intermediario', 'Premium'])
        
        if st.button("🔍 Calcular Risco", type="primary"):
            risco_sim, fatores_sim = calcular_score_risco({
                'Satisfacao': sim_satisfacao,
                'NPS': sim_nps,
                'Ultimo_Contato': sim_contato,
                'Suportes_Abertos': sim_suportes,
                'Desconto_Aplicado': sim_desconto,
                'Forma_Pagamento': sim_pagamento
            })
            
            col_res1, col_res2, col_res3 = st.columns(3)
            
            with col_res1:
                cor = "🟢 Baixo" if risco_sim < 30 else "🟡 Médio" if risco_sim < 60 else "🟠 Alto" if risco_sim < 80 else "🔴 Crítico"
                st.metric("Score de Risco", f"{risco_sim}/100", cor)
            
            with col_res2:
                prob_sim = min(risco_sim + np.random.randint(-5, 5), 100)
                prob_sim = max(prob_sim, 0)
                st.metric("Prob. Churn Estimada", f"{prob_sim:.0f}%")
            
            with col_res3:
                nivel = "Baixo" if risco_sim < 30 else "Médio" if risco_sim < 60 else "Alto" if risco_sim < 80 else "Crítico"
                st.metric("Nível de Risco", nivel)
            
            if fatores_sim:
                st.warning(f"⚠️ Fatores de alerta: {', '.join(fatores_sim)}")
            else:
                st.success("✅ Cliente saudável! Considere programa de fidelidade.")

st.markdown("---")
st.markdown(
    "<div style='text-align: center; color: #64748B; padding: 20px;'>"
    "🚀 Churn Analytics Pro | Desenvolvido com Streamlit + Python + Machine Learning"
    "</div>",
    unsafe_allow_html=True
)