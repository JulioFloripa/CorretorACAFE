import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from fpdf import FPDF
import tempfile
import zipfile
import os
import traceback

# Configurar matplotlib para usar backend não-interativo
import matplotlib
matplotlib.use('Agg')

# --------------------------
# CONFIGURAÇÕES DE ESTILO
# --------------------------

def load_css():
    """Carrega CSS customizado para tema verde ACAFE"""
    st.markdown("""
    <style>
    /* Tema principal verde ACAFE */
    .main {
        background: linear-gradient(135deg, #f8fffe 0%, #e8f5f3 100%);
    }
    
    /* Header customizado */
    .header-container {
        background: linear-gradient(90deg, #2d5a3d 0%, #4a8c6a 100%);
        padding: 2rem 1rem;
        border-radius: 15px;
        margin-bottom: 2rem;
        box-shadow: 0 4px 15px rgba(45, 90, 61, 0.2);
    }
    
    .header-title {
        color: white;
        font-size: 2.5rem;
        font-weight: bold;
        text-align: center;
        margin: 0;
        text-shadow: 2px 2px 4px rgba(0,0,0,0.3);
    }
    
    .header-subtitle {
        color: #e8f5f3;
        font-size: 1.2rem;
        text-align: center;
        margin-top: 0.5rem;
        font-style: italic;
    }
    
    /* Logo ACAFE */
    .logo-container {
        display: flex;
        justify-content: center;
        align-items: center;
        margin-bottom: 1rem;
    }
    
    .acafe-logo {
        width: 80px;
        height: 80px;
        background: white;
        border-radius: 50%;
        display: flex;
        align-items: center;
        justify-content: center;
        box-shadow: 0 4px 10px rgba(0,0,0,0.2);
        margin-right: 1rem;
    }
    
    /* Métricas customizadas */
    .metric-container {
        background: white;
        padding: 1rem;
        border-radius: 10px;
        border-left: 4px solid #4a8c6a;
        box-shadow: 0 2px 5px rgba(0,0,0,0.1);
        margin: 0.5rem 0;
    }
    
    /* Progress bar customizada */
    .stProgress > div > div > div > div {
        background: linear-gradient(90deg, #4a8c6a, #2d5a3d);
    }
    
    /* Footer */
    .footer {
        text-align: center;
        padding: 2rem;
        color: #2d5a3d;
        font-style: italic;
        border-top: 2px solid #e8f5f3;
        margin-top: 3rem;
    }
    </style>
    """, unsafe_allow_html=True)

def create_acafe_logo():
    """Cria logo ACAFE em SVG"""
    logo_svg = """
    <svg width="60" height="60" viewBox="0 0 100 100" xmlns="http://www.w3.org/2000/svg">
        <circle cx="50" cy="50" r="45" fill="#2d5a3d" stroke="white" stroke-width="3"/>
        <path d="M25 35 L50 25 L75 35 L75 65 L50 75 L25 65 Z" fill="white" opacity="0.9"/>
        <text x="50" y="45" text-anchor="middle" fill="#2d5a3d" font-family="Arial, sans-serif" font-size="12" font-weight="bold">ACAFE</text>
        <text x="50" y="60" text-anchor="middle" fill="#2d5a3d" font-family="Arial, sans-serif" font-size="8">FLEMING</text>
    </svg>
    """
    return logo_svg

def show_header():
    """Mostra header customizado"""
    st.markdown(f"""
    <div class="header-container">
        <div class="logo-container">
            <div class="acafe-logo">
                {create_acafe_logo()}
            </div>
            <div>
                <h1 class="header-title">Corretor ACAFE Fleming - DEBUG</h1>
                <p class="header-subtitle">Versão de Diagnóstico</p>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)

# --------------------------
# CONFIGURAÇÕES INICIAIS
# --------------------------
st.set_page_config(
    page_title="Corretor ACAFE Fleming - DEBUG", 
    layout="wide",
    page_icon="🔧",
    initial_sidebar_state="expanded"
)

# Aplicar CSS customizado
load_css()

# Mostrar header
show_header()

st.markdown("### 🔧 **Versão DEBUG - Diagnóstico de Problemas**")
st.markdown("### 📚 Faça upload da planilha com as abas **RESPOSTAS** e **GABARITO**")

# --------------------------
# FUNÇÕES DE VALIDAÇÃO
# --------------------------

def validar_arquivo_excel(dados):
    """Valida se o arquivo Excel tem a estrutura esperada"""
    erros = []
    
    # Verificar se as abas existem
    if "RESPOSTAS" not in dados:
        erros.append("❌ Aba 'RESPOSTAS' não encontrada no arquivo")
    if "GABARITO" not in dados:
        erros.append("❌ Aba 'GABARITO' não encontrada no arquivo")
    
    if erros:
        return False, erros
    
    respostas = dados["RESPOSTAS"]
    gabarito = dados["GABARITO"]
    
    # Verificar colunas obrigatórias na aba RESPOSTAS
    colunas_obrigatorias_respostas = ["ID", "Nome"]
    for col in colunas_obrigatorias_respostas:
        if col not in respostas.columns:
            erros.append(f"❌ Coluna '{col}' não encontrada na aba RESPOSTAS")
    
    # Verificar colunas obrigatórias na aba GABARITO
    colunas_obrigatorias_gabarito = ["Questão", "Resposta", "Disciplina"]
    for col in colunas_obrigatorias_gabarito:
        if col not in gabarito.columns:
            erros.append(f"❌ Coluna '{col}' não encontrada na aba GABARITO")
    
    # Verificar se há dados
    if len(respostas) == 0:
        erros.append("❌ Aba RESPOSTAS está vazia")
    if len(gabarito) == 0:
        erros.append("❌ Aba GABARITO está vazia")
    
    return len(erros) == 0, erros

def validar_dados_gabarito(gabarito):
    """Valida os dados do gabarito"""
    erros = []
    
    # Verificar questões duplicadas APENAS dentro da mesma disciplina
    for disciplina in gabarito['Disciplina'].unique():
        if pd.isna(disciplina):
            continue
        
        gabarito_disciplina = gabarito[gabarito['Disciplina'] == disciplina]
        questoes_duplicadas = gabarito_disciplina[gabarito_disciplina.duplicated(subset=['Questão'], keep=False)]
        
        if len(questoes_duplicadas) > 0:
            questoes_dup = questoes_duplicadas['Questão'].unique().tolist()
            erros.append(f"❌ Questões duplicadas em {disciplina}: {questoes_dup}")
    
    # Verificar se há valores nulos
    if gabarito['Questão'].isnull().any():
        erros.append("❌ Há questões com número vazio no gabarito")
    if gabarito['Resposta'].isnull().any():
        erros.append("❌ Há questões sem resposta no gabarito")
    if gabarito['Disciplina'].isnull().any():
        erros.append("❌ Há questões sem disciplina no gabarito")
    
    # Verificar questões de línguas estrangeiras (informativo)
    linguas = ['Inglês', 'Espanhol', 'Ingles', 'Espanol']
    questoes_linguas = gabarito[gabarito['Disciplina'].isin(linguas)]
    
    if len(questoes_linguas) > 0:
        st.info(f"ℹ️ Detectadas {len(questoes_linguas)} questões de línguas estrangeiras.")
    
    return len(erros) == 0, erros

# --------------------------
# FUNÇÕES AUXILIARES COM DEBUG
# --------------------------

def corrigir_respostas_debug(df_respostas, gabarito):
    """Corrige as respostas dos alunos baseado no gabarito - COM DEBUG"""
    respostas = df_respostas.copy()
    
    st.markdown("### 🔍 **DEBUG - Processo de Correção**")
    
    # Debug: Mostrar estrutura dos dados
    st.markdown("#### 📊 **Estrutura dos Dados:**")
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("**Colunas RESPOSTAS:**")
        st.write(list(respostas.columns))
        st.markdown("**Exemplo de linha:**")
        if len(respostas) > 0:
            st.write(respostas.iloc[0].to_dict())
    
    with col2:
        st.markdown("**Estrutura GABARITO:**")
        st.write(gabarito[['Questão', 'Resposta', 'Disciplina']].head(10))
    
    # Processo de correção com debug
    acertos_debug = []
    
    for _, row_gabarito in gabarito.iterrows():
        questao = row_gabarito["Questão"]
        resposta_correta = row_gabarito["Resposta"]
        disciplina = row_gabarito["Disciplina"]
        col = f"Q{int(questao)}"
        
        if col in respostas.columns:
            # Verificar quantos alunos acertaram esta questão
            acertos = (respostas[col] == resposta_correta).sum()
            total_alunos = len(respostas)
            
            acertos_debug.append({
                'Questão': questao,
                'Disciplina': disciplina,
                'Resposta_Correta': resposta_correta,
                'Coluna': col,
                'Acertos': acertos,
                'Total': total_alunos,
                'Percentual': round(acertos/total_alunos*100, 1) if total_alunos > 0 else 0
            })
            
            # Criar coluna de correção
            respostas[f"{col}_OK"] = respostas[col] == resposta_correta
        else:
            acertos_debug.append({
                'Questão': questao,
                'Disciplina': disciplina,
                'Resposta_Correta': resposta_correta,
                'Coluna': col,
                'Acertos': 0,
                'Total': len(respostas),
                'Percentual': 0,
                'Erro': 'Coluna não encontrada'
            })
            respostas[f"{col}_OK"] = False
    
    # Mostrar debug dos acertos
    st.markdown("#### 📈 **Debug - Acertos por Questão:**")
    debug_df = pd.DataFrame(acertos_debug)
    st.dataframe(debug_df, use_container_width=True)
    
    # Estatísticas gerais
    total_acertos = debug_df['Acertos'].sum()
    total_questoes = len(debug_df)
    total_respostas = total_questoes * len(respostas)
    
    st.markdown(f"""
    **📊 Estatísticas Gerais:**
    - Total de acertos: {total_acertos}
    - Total de questões: {total_questoes}
    - Total de respostas possíveis: {total_respostas}
    - Percentual geral de acertos: {total_acertos/total_respostas*100:.1f}%
    """)
    
    return respostas, debug_df

# --------------------------
# INTERFACE PRINCIPAL
# --------------------------

# Upload do arquivo
arquivo = st.file_uploader(
    "📎 **Selecione o arquivo Excel**", 
    type=["xlsx"], 
    help="Arquivo deve conter as abas 'RESPOSTAS' e 'GABARITO'",
    key="file_uploader"
)

if arquivo:
    try:
        # Mostrar progresso
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        status_text.success("📖 Lendo arquivo Excel...")
        progress_bar.progress(20)
        
        # Ler arquivo Excel
        dados = pd.read_excel(arquivo, sheet_name=None)
        
        # Validar arquivo
        valido, erros = validar_arquivo_excel(dados)
        if not valido:
            st.error("**🚨 Problemas encontrados no arquivo:**")
            for erro in erros:
                st.error(erro)
            st.stop()
        
        respostas = dados["RESPOSTAS"]
        gabarito = dados["GABARITO"]
        
        # Validar gabarito
        gabarito_valido, erros_gabarito = validar_dados_gabarito(gabarito)
        if not gabarito_valido:
            st.error("**🚨 Problemas encontrados no gabarito:**")
            for erro in erros_gabarito:
                st.error(erro)
            st.stop()
        
        progress_bar.progress(50)
        
        # Mostrar preview dos dados
        st.markdown("### 📋 **Preview dos Dados**")
        col1, col2 = st.columns(2)
        with col1:
            st.markdown("#### **Respostas (primeiras 3 linhas):**")
            st.dataframe(respostas.head(3), use_container_width=True)
        
        with col2:
            st.markdown("#### **Gabarito (primeiras 10 questões):**")
            st.dataframe(gabarito.head(10), use_container_width=True)
        
        progress_bar.progress(70)
        
        # Processo de correção com debug
        respostas_corr, debug_df = corrigir_respostas_debug(respostas, gabarito)
        
        progress_bar.progress(90)
        
        # Análise de um aluno específico
        st.markdown("### 👤 **Debug - Análise de Aluno Específico**")
        
        if len(respostas_corr) > 0:
            aluno_exemplo = respostas_corr.iloc[0]
            st.markdown(f"**Analisando: {aluno_exemplo['Nome']}**")
            
            # Mostrar respostas do aluno vs gabarito
            analise_aluno = []
            for _, row_gab in gabarito.head(10).iterrows():  # Primeiras 10 questões
                questao = row_gab["Questão"]
                col = f"Q{int(questao)}"
                
                if col in aluno_exemplo:
                    resposta_aluno = aluno_exemplo[col]
                    resposta_correta = row_gab["Resposta"]
                    acertou = resposta_aluno == resposta_correta
                    
                    analise_aluno.append({
                        'Questão': questao,
                        'Disciplina': row_gab["Disciplina"],
                        'Resposta_Aluno': resposta_aluno,
                        'Resposta_Correta': resposta_correta,
                        'Acertou': '✅' if acertou else '❌'
                    })
            
            analise_df = pd.DataFrame(analise_aluno)
            st.dataframe(analise_df, use_container_width=True)
            
            # Calcular acertos do aluno
            acertos_aluno = sum([aluno_exemplo.get(f"Q{int(q)}_OK", False) for q in gabarito["Questão"]])
            total_questoes = len(gabarito["Questão"].unique())
            percentual_aluno = acertos_aluno / total_questoes * 100 if total_questoes > 0 else 0
            
            st.markdown(f"""
            **📊 Resultado do Aluno:**
            - Acertos: {acertos_aluno}/{total_questoes}
            - Percentual: {percentual_aluno:.1f}%
            """)
        
        progress_bar.progress(100)
        status_text.success("✅ Análise de debug concluída!")
        
    except Exception as e:
        st.error(f"❌ **Erro durante o processamento:** {str(e)}")
        st.code(traceback.format_exc())

# Footer
st.markdown("""
<div class="footer">
    <p><strong>Corretor ACAFE - Versão DEBUG</strong></p>
    <p>Esta versão mostra detalhes do processo de correção para identificar problemas</p>
</div>
""", unsafe_allow_html=True)

