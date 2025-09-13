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
# CONFIGURAÇÕES INICIAIS
# --------------------------
st.set_page_config(page_title="Gerador de Boletins - Fleming", layout="wide")
st.title("📊 Gerador de Boletins - Colégio Fleming")
st.markdown("Faça upload da planilha com as abas **RESPOSTAS** e **GABARITO**.")

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
    
    # Verificar se há questões duplicadas
    questoes_duplicadas = gabarito[gabarito.duplicated(subset=['Questão'], keep=False)]
    if len(questoes_duplicadas) > 0:
        erros.append(f"❌ Questões duplicadas encontradas: {questoes_duplicadas['Questão'].unique().tolist()}")
    
    # Verificar se há valores nulos
    if gabarito['Questão'].isnull().any():
        erros.append("❌ Há questões com número vazio no gabarito")
    if gabarito['Resposta'].isnull().any():
        erros.append("❌ Há questões sem resposta no gabarito")
    if gabarito['Disciplina'].isnull().any():
        erros.append("❌ Há questões sem disciplina no gabarito")
    
    return len(erros) == 0, erros

# --------------------------
# FUNÇÕES AUXILIARES
# --------------------------

def corrigir_respostas(df_respostas, gabarito, mapa_disciplinas):
    """Corrige as respostas dos alunos baseado no gabarito"""
    respostas = df_respostas.copy()
    
    for disc, questoes in mapa_disciplinas.items():
        for q in questoes:
            col = f"Q{int(q)}"
            if col in respostas.columns:
                try:
                    resp_correta = gabarito.loc[gabarito["Questão"] == q, "Resposta"].values[0]
                    respostas[f"{col}_OK"] = respostas[col] == resp_correta
                except IndexError:
                    st.warning(f"⚠️ Questão {q} não encontrada no gabarito")
                    respostas[f"{col}_OK"] = False
            else:
                respostas[f"{col}_OK"] = False
    
    return respostas

def resultados_disciplina(linha, mapa_disciplinas):
    """Calcula os resultados por disciplina para um aluno"""
    resultados = []
    for disc, questoes in mapa_disciplinas.items():
        acertos = sum([linha.get(f"Q{int(q)}_OK", False) for q in questoes])
        total = len(questoes)
        perc = round(100 * acertos / total, 1) if total > 0 else 0
        resultados.append((disc, acertos, total, perc))
    return resultados

def gerar_graficos(nome, posicao, percentual, df_boletim, media_df, ranking_df, pasta):
    """Gera os gráficos para o boletim individual"""
    try:
        labels = df_boletim["Disciplina"].tolist()
        aluno_vals = df_boletim["%"].values
        media_vals = media_df["%"].values

        # Configurar estilo dos gráficos
        plt.style.use('default')
        
        # Radar
        if len(labels) > 0:
            angles = np.linspace(0, 2 * np.pi, len(labels), endpoint=False).tolist()
            aluno_circ = np.concatenate((aluno_vals, [aluno_vals[0]]))
            media_circ = np.concatenate((media_vals, [media_vals[0]]))
            angles += [angles[0]]

            fig = plt.figure(figsize=(6, 6))
            ax = plt.subplot(111, polar=True)
            ax.plot(angles, aluno_circ, "o-", label=nome, linewidth=2)
            ax.fill(angles, aluno_circ, alpha=0.25)
            ax.plot(angles, media_circ, "o--", label="Média da Turma", color="gray", linewidth=2)
            ax.fill(angles, media_circ, alpha=0.1, color="gray")
            ax.set_thetagrids(np.degrees(angles[:-1]), labels)
            ax.legend(loc="upper right", bbox_to_anchor=(1.2, 1.1))
            ax.set_ylim(0, 100)
            radar_path = os.path.join(pasta, f"{nome}_radar.png")
            plt.savefig(radar_path, bbox_inches="tight", dpi=150)
            plt.close()
        else:
            radar_path = None

        # Barras
        x = np.arange(len(labels))
        bar_width = 0.35
        plt.figure(figsize=(12, 6))
        bars1 = plt.bar(x - bar_width/2, aluno_vals, bar_width, label=nome, color="teal", alpha=0.8)
        bars2 = plt.bar(x + bar_width/2, media_vals, bar_width, label="Média Turma", color="lightgray", alpha=0.8)
        
        # Adicionar valores nas barras
        for i, v in enumerate(aluno_vals):
            plt.text(i - bar_width/2, v + 1, f"{v:.1f}%", ha="center", fontsize=9, fontweight='bold')
        for i, v in enumerate(media_vals):
            plt.text(i + bar_width/2, v + 1, f"{v:.1f}%", ha="center", fontsize=9)
            
        plt.xticks(ticks=x, labels=labels, rotation=45, ha='right')
        plt.ylabel("Percentual de Acertos (%)")
        plt.title(f"Desempenho por Disciplina - {nome}", fontsize=14, fontweight='bold')
        plt.legend()
        plt.grid(axis='y', alpha=0.3)
        plt.ylim(0, 105)
        barras_path = os.path.join(pasta, f"{nome}_barras.png")
        plt.savefig(barras_path, bbox_inches="tight", dpi=150)
        plt.close()

        # Distribuição
        plt.figure(figsize=(10, 6))
        sns.histplot(ranking_df["Percentual"]*100, bins=min(10, len(ranking_df)), 
                    color="lightblue", edgecolor="black", alpha=0.7)
        plt.axvline(percentual, color="red", linewidth=3, 
                   label=f"{nome} ({percentual:.1f}%)", linestyle='--')
        plt.xlabel("Percentual de Acertos (%)")
        plt.ylabel("Número de Estudantes")
        plt.title("Distribuição das Notas da Turma", fontsize=14, fontweight='bold')
        plt.legend()
        plt.grid(alpha=0.3)
        dist_path = os.path.join(pasta, f"{nome}_dist.png")
        plt.savefig(dist_path, bbox_inches="tight", dpi=150)
        plt.close()

        # Ranking
        plt.figure(figsize=(10, 6))
        plt.plot(ranking_df["Posição"], ranking_df["Percentual"]*100, "o-", 
                color="lightgray", markersize=6, linewidth=2, alpha=0.7)
        plt.scatter(posicao, percentual, color="red", s=150, 
                   label=f"{nome} - {posicao}º lugar", zorder=5, edgecolor='black')
        plt.xlabel("Posição no Ranking")
        plt.ylabel("Percentual de Acertos (%)")
        plt.title("Ranking da Turma", fontsize=14, fontweight='bold')
        plt.legend()
        plt.grid(alpha=0.3)
        rank_path = os.path.join(pasta, f"{nome}_rank.png")
        plt.savefig(rank_path, bbox_inches="tight", dpi=150)
        plt.close()

        return barras_path, radar_path, dist_path, rank_path
    
    except Exception as e:
        st.error(f"Erro ao gerar gráficos para {nome}: {str(e)}")
        return None, None, None, None

class BoletimPDF(FPDF):
    def header(self):
        self.set_font("Arial", "B", 14)
        self.cell(0, 15, "SIMULADO ACAFE - RELATÓRIO INDIVIDUAL", ln=True, align="C")
        self.ln(5)

    def add_aluno_info(self, nome, posicao, percentual, media_turma):
        self.set_font("Arial", "B", 12)
        self.cell(0, 10, f"Aluno: {nome}", ln=True)
        self.set_font("Arial", "", 11)
        self.cell(0, 8, f"Posição no Ranking: {posicao}º lugar", ln=True)
        self.cell(0, 8, f"Nota Individual: {percentual:.1f}%", ln=True)
        self.cell(0, 8, f"Média da Turma: {media_turma:.1f}%", ln=True)
        diferenca = percentual - media_turma
        if diferenca > 0:
            self.cell(0, 8, f"Diferença: +{diferenca:.1f}% (acima da média)", ln=True)
        else:
            self.cell(0, 8, f"Diferença: {diferenca:.1f}% (abaixo da média)", ln=True)
        self.ln(5)

    def add_table(self, df):
        self.set_font("Arial", "B", 9)
        # Cabeçalho da tabela
        self.cell(35, 8, "Disciplina", 1, 0, 'C')
        self.cell(20, 8, "Acertos", 1, 0, 'C')
        self.cell(20, 8, "Total", 1, 0, 'C')
        self.cell(20, 8, "Nota (%)", 1, 0, 'C')
        self.cell(25, 8, "Média (%)", 1, 0, 'C')
        self.cell(25, 8, "Diferença", 1, 0, 'C')
        self.ln()
        
        self.set_font("Arial", "", 9)
        for _, row in df.iterrows():
            self.cell(35, 7, str(row["Disciplina"])[:15], 1, 0, 'L')
            self.cell(20, 7, str(row["Acertos"]), 1, 0, 'C')
            self.cell(20, 7, str(row["Total"]), 1, 0, 'C')
            self.cell(20, 7, f"{row['%']:.1f}%", 1, 0, 'C')
            self.cell(25, 7, f"{row['Média Turma']:.1f}%", 1, 0, 'C')
            diferenca = row['Diferença']
            cor_diferenca = "+" if diferenca > 0 else ""
            self.cell(25, 7, f"{cor_diferenca}{diferenca:.1f}%", 1, 0, 'C')
            self.ln()
        self.ln(5)

    def add_image(self, path, largura=160):
        if path and os.path.exists(path):
            try:
                self.image(path, x=(210-largura)/2, w=largura)
                self.ln(5)
            except Exception as e:
                self.set_font("Arial", "", 10)
                self.cell(0, 10, f"Erro ao carregar imagem: {str(e)}", ln=True)

# --------------------------
# INTERFACE STREAMLIT
# --------------------------

# Sidebar com informações
with st.sidebar:
    st.header("ℹ️ Instruções")
    st.markdown("""
    **Formato do arquivo Excel:**
    
    **Aba RESPOSTAS:**
    - Coluna 'ID': Identificador único do aluno
    - Coluna 'Nome': Nome completo do aluno
    - Colunas 'Q1', 'Q2', etc.: Respostas (A, B, C, D, E)
    
    **Aba GABARITO:**
    - Coluna 'Questão': Número da questão
    - Coluna 'Resposta': Resposta correta (A, B, C, D, E)
    - Coluna 'Disciplina': Nome da disciplina
    """)
    
    st.header("📊 Estatísticas")
    if 'stats' in st.session_state:
        stats = st.session_state.stats
        st.metric("Total de Alunos", stats.get('total_alunos', 0))
        st.metric("Total de Questões", stats.get('total_questoes', 0))
        st.metric("Disciplinas", stats.get('total_disciplinas', 0))

# Upload do arquivo
arquivo = st.file_uploader("📎 Upload do arquivo Excel", type=["xlsx"], 
                          help="Arquivo deve conter as abas 'RESPOSTAS' e 'GABARITO'")

if arquivo:
    try:
        # Mostrar progresso
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        status_text.text("📖 Lendo arquivo Excel...")
        progress_bar.progress(10)
        
        # Ler arquivo Excel
        dados = pd.read_excel(arquivo, sheet_name=None)
        
        status_text.text("✅ Validando estrutura do arquivo...")
        progress_bar.progress(20)
        
        # Validar arquivo
        valido, erros = validar_arquivo_excel(dados)
        if not valido:
            st.error("**Problemas encontrados no arquivo:**")
            for erro in erros:
                st.error(erro)
            st.stop()
        
        respostas = dados["RESPOSTAS"]
        gabarito = dados["GABARITO"]
        
        # Validar gabarito
        gabarito_valido, erros_gabarito = validar_dados_gabarito(gabarito)
        if not gabarito_valido:
            st.error("**Problemas encontrados no gabarito:**")
            for erro in erros_gabarito:
                st.error(erro)
            st.stop()
        
        status_text.text("📊 Processando dados...")
        progress_bar.progress(30)
        
        # Mostrar preview dos dados
        col1, col2 = st.columns(2)
        with col1:
            st.subheader("📋 Preview - Respostas")
            st.dataframe(respostas.head(), use_container_width=True)
        
        with col2:
            st.subheader("📝 Preview - Gabarito")
            st.dataframe(gabarito.head(), use_container_width=True)
        
        # Estatísticas
        total_alunos = len(respostas)
        total_questoes = len(gabarito)
        disciplinas = gabarito['Disciplina'].unique()
        total_disciplinas = len(disciplinas)
        
        st.session_state.stats = {
            'total_alunos': total_alunos,
            'total_questoes': total_questoes,
            'total_disciplinas': total_disciplinas
        }
        
        # Mostrar estatísticas
        col1, col2, col3 = st.columns(3)
        col1.metric("👥 Total de Alunos", total_alunos)
        col2.metric("❓ Total de Questões", total_questoes)
        col3.metric("📚 Disciplinas", total_disciplinas)
        
        st.subheader("📚 Disciplinas encontradas:")
        st.write(", ".join(disciplinas))
        
        # Processar dados
        status_text.text("🔄 Corrigindo respostas...")
        progress_bar.progress(40)
        
        # Mapeamento disciplinas
        mapa_disciplinas = (
            gabarito[["Questão", "Disciplina"]]
            .dropna()
            .groupby("Disciplina")["Questão"]
            .apply(list)
            .to_dict()
        )

        respostas_corr = corrigir_respostas(respostas, gabarito, mapa_disciplinas)
        
        status_text.text("📈 Calculando ranking...")
        progress_bar.progress(50)

        # Ranking
        percentuais = []
        for i, row in respostas_corr.iterrows():
            acertos_tot = sum([row.get(f"Q{int(q)}_OK", False) for q in gabarito["Questão"]])
            percentuais.append(acertos_tot / len(gabarito))
        respostas_corr["Percentual"] = percentuais

        ranking_df = respostas_corr[["ID", "Nome", "Percentual"]].sort_values("Percentual", ascending=False).reset_index(drop=True)
        ranking_df["Posição"] = ranking_df.index + 1
        ranking_df["Nota (%)"] = (ranking_df["Percentual"] * 100).round(1)
        media_turma = ranking_df["Percentual"].mean() * 100

        # Mostrar ranking
        st.subheader("🏆 Ranking da Turma")
        st.dataframe(ranking_df[["Posição", "Nome", "Nota (%)"]].head(10), use_container_width=True)
        
        status_text.text("📊 Calculando médias por disciplina...")
        progress_bar.progress(60)

        # Médias por disciplina
        media_disciplinas = []
        for disc, questoes in mapa_disciplinas.items():
            acertos = []
            for _, row in respostas_corr.iterrows():
                acertos.append(sum([row.get(f"Q{int(q)}_OK", False) for q in questoes]) / len(questoes))
            media_disciplinas.append((disc, round(np.mean(acertos)*100, 1)))
        media_df = pd.DataFrame(media_disciplinas, columns=["Disciplina", "%"])
        
        # Mostrar médias por disciplina
        st.subheader("📊 Médias por Disciplina")
        st.dataframe(media_df, use_container_width=True)
        
        status_text.text("📄 Gerando boletins individuais...")
        progress_bar.progress(70)

        # Gerar boletins
        with tempfile.TemporaryDirectory() as tmpdir:
            zip_path = os.path.join(tmpdir, "boletins.zip")
            
            with zipfile.ZipFile(zip_path, "w") as zipf:
                total_alunos = len(respostas_corr)
                
                for i, aluno in respostas_corr.iterrows():
                    # Atualizar progresso
                    progresso = 70 + (i / total_alunos) * 25
                    progress_bar.progress(int(progresso))
                    status_text.text(f"📄 Gerando boletim: {aluno['Nome']} ({i+1}/{total_alunos})")
                    
                    nome = aluno["Nome"].replace(" ", "_").replace("/", "_")
                    posicao = int(ranking_df.loc[ranking_df["ID"] == aluno["ID"], "Posição"].iloc[0])
                    percentual = aluno["Percentual"] * 100

                    resultados = resultados_disciplina(aluno, mapa_disciplinas)
                    df_boletim = pd.DataFrame(resultados, columns=["Disciplina", "Acertos", "Total", "%"])
                    df_boletim["Média Turma"] = media_df["%"]
                    df_boletim["Diferença"] = (df_boletim["%"] - media_df["%"]).round(1)

                    # Gráficos
                    barras, radar, dist, rank = gerar_graficos(nome, posicao, percentual, df_boletim, media_df, ranking_df, tmpdir)

                    # PDF
                    try:
                        pdf = BoletimPDF()
                        pdf.add_page()
                        pdf.add_aluno_info(aluno["Nome"], posicao, percentual, media_turma)
                        pdf.add_table(df_boletim)
                        
                        if barras:
                            pdf.add_image(barras)
                        if radar:
                            pdf.add_image(radar)
                        if dist:
                            pdf.add_image(dist)
                        if rank:
                            pdf.add_image(rank)

                        pdf_path = os.path.join(tmpdir, f"Boletim_{nome}.pdf")
                        pdf.output(pdf_path)
                        zipf.write(pdf_path, f"Boletim_{nome}.pdf")
                    
                    except Exception as e:
                        st.warning(f"⚠️ Erro ao gerar PDF para {aluno['Nome']}: {str(e)}")
                        continue

            status_text.text("✅ Processamento concluído!")
            progress_bar.progress(100)
            
            # Botão de download
            with open(zip_path, "rb") as f:
                st.download_button(
                    "📥 Baixar todos os boletins (ZIP)", 
                    f.read(), 
                    "boletins_acafe.zip", 
                    "application/zip",
                    help=f"Arquivo contém {total_alunos} boletins individuais em PDF"
                )
            
            st.success(f"✅ {total_alunos} boletins gerados com sucesso!")
            
    except Exception as e:
        st.error(f"❌ Erro durante o processamento: {str(e)}")
        st.error("**Detalhes do erro:**")
        st.code(traceback.format_exc())
        st.info("💡 Verifique se o arquivo está no formato correto e tente novamente.")

