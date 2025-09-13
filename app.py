import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from fpdf import FPDF
import tempfile
import zipfile
import os

# --------------------------
# CONFIGURAÇÕES INICIAIS
# --------------------------
st.set_page_config(page_title="Gerador de Boletins - Fleming", layout="wide")
st.title("📊 Gerador de Boletins - Colégio Fleming")
st.markdown("Faça upload da planilha com as abas **RESPOSTAS** e **GABARITO**.")

# --------------------------
# FUNÇÕES AUXILIARES
# --------------------------

def corrigir_respostas(df_respostas, gabarito, mapa_disciplinas):
    respostas = df_respostas.copy()
    for disc, questoes in mapa_disciplinas.items():
        for q in questoes:
            col = f"Q{int(q)}"
            if col in respostas.columns:
                resp_correta = gabarito.loc[gabarito["Questão"] == q, "Resposta"].values[0]
                respostas[f"{col}_OK"] = respostas[col] == resp_correta
            else:
                respostas[f"{col}_OK"] = False
    return respostas

def resultados_disciplina(linha, mapa_disciplinas):
    resultados = []
    for disc, questoes in mapa_disciplinas.items():
        acertos = sum([linha.get(f"Q{int(q)}_OK", False) for q in questoes])
        total = len(questoes)
        perc = round(100 * acertos / total, 1) if total > 0 else 0
        resultados.append((disc, acertos, total, perc))
    return resultados

def gerar_graficos(nome, posicao, percentual, df_boletim, media_df, ranking_df, pasta):
    labels = df_boletim["Disciplina"].tolist()
    aluno_vals = df_boletim["%"].values
    media_vals = media_df["%"].values

    # Radar
    angles = np.linspace(0, 2 * np.pi, len(labels), endpoint=False).tolist()
    aluno_circ = np.concatenate((aluno_vals, [aluno_vals[0]]))
    media_circ = np.concatenate((media_vals, [media_vals[0]]))
    angles += [angles[0]]

    fig = plt.figure(figsize=(6, 6))
    ax = plt.subplot(111, polar=True)
    ax.plot(angles, aluno_circ, "o-", label=nome)
    ax.fill(angles, aluno_circ, alpha=0.25)
    ax.plot(angles, media_circ, "o--", label="Média da Turma", color="gray")
    ax.fill(angles, media_circ, alpha=0.1, color="gray")
    ax.set_thetagrids(np.degrees(angles[:-1]), labels)
    ax.legend(loc="upper right", bbox_to_anchor=(1.2, 1.1))
    radar_path = os.path.join(pasta, f"{nome}_radar.png")
    plt.savefig(radar_path, bbox_inches="tight")
    plt.close()

    # Barras
    x = np.arange(len(labels))
    bar_width = 0.35
    plt.figure(figsize=(10, 5))
    plt.bar(x - bar_width/2, aluno_vals, bar_width, label=nome, color="teal")
    plt.bar(x + bar_width/2, media_vals, bar_width, label="Média Turma", color="lightgray")
    for i, v in enumerate(aluno_vals):
        plt.text(i - bar_width/2, v + 1, f"{v:.1f}%", ha="center", fontsize=8)
    plt.xticks(ticks=x, labels=labels, rotation=45)
    plt.ylabel("Percentual de Acertos (%)")
    plt.title(f"Desempenho por Disciplina - {nome}")
    plt.legend()
    barras_path = os.path.join(pasta, f"{nome}_barras.png")
    plt.savefig(barras_path, bbox_inches="tight")
    plt.close()

    # Distribuição
    plt.figure(figsize=(8, 4))
    sns.histplot(ranking_df["Percentual"]*100, bins=10, color="lightgray", edgecolor="black")
    plt.axvline(percentual, color="teal", linewidth=2, label=f"{nome} ({percentual:.1f}%)")
    plt.xlabel("Percentual de Acertos (%)")
    plt.ylabel("Nº Estudantes")
    plt.title("Distribuição das Notas")
    plt.legend()
    dist_path = os.path.join(pasta, f"{nome}_dist.png")
    plt.savefig(dist_path, bbox_inches="tight")
    plt.close()

    # Ranking
    plt.figure(figsize=(8, 4))
    plt.plot(ranking_df["Posição"], ranking_df["Percentual"]*100, "o-", color="lightgray")
    plt.scatter(posicao, percentual, color="teal", s=100, label=f"{nome} - {posicao}º lugar")
    plt.xlabel("Posição no Ranking")
    plt.ylabel("Percentual de Acertos (%)")
    plt.title("Ranking da Turma")
    plt.legend()
    rank_path = os.path.join(pasta, f"{nome}_rank.png")
    plt.savefig(rank_path, bbox_inches="tight")
    plt.close()

    return barras_path, radar_path, dist_path, rank_path

class BoletimPDF(FPDF):
    def header(self):
        self.set_font("Arial", "B", 12)
        self.cell(0, 10, "SIMULADO ACAFE - RELATÓRIO INDIVIDUAL", ln=True, align="C")
        self.ln(5)

    def add_aluno_info(self, nome, posicao, percentual, media_turma):
        self.set_font("Arial", "", 11)
        self.cell(0, 10, f"Nome: {nome}", ln=True)
        self.cell(0, 10, f"Posição: {posicao} | Nota: {percentual:.1f}% | Média Turma: {media_turma:.1f}%", ln=True)
        self.ln(5)

    def add_table(self, df):
        self.set_font("Arial", "B", 10)
        self.cell(40, 8, "Disciplina", 1)
        self.cell(25, 8, "Acertos", 1)
        self.cell(25, 8, "Total", 1)
        self.cell(25, 8, "%", 1)
        self.cell(35, 8, "Média Turma", 1)
        self.cell(30, 8, "Diferença", 1)
        self.ln()
        self.set_font("Arial", "", 10)
        for _, row in df.iterrows():
            self.cell(40, 8, row["Disciplina"], 1)
            self.cell(25, 8, str(row["Acertos"]), 1)
            self.cell(25, 8, str(row["Total"]), 1)
            self.cell(25, 8, f"{row['%']}%", 1)
            self.cell(35, 8, f"{row['Média Turma']}%", 1)
            self.cell(30, 8, f"{row['Diferença']}%", 1)
            self.ln()
        self.ln(5)

    def add_image(self, path, largura=160):
        self.image(path, x=(210-largura)/2, w=largura)
        self.ln(5)

# --------------------------
# INTERFACE STREAMLIT
# --------------------------
arquivo = st.file_uploader("📎 Upload do arquivo Excel", type=["xlsx"])

if arquivo:
    with tempfile.TemporaryDirectory() as tmpdir:
        dados = pd.read_excel(arquivo, sheet_name=None)
        respostas = dados["RESPOSTAS"]
        gabarito = dados["GABARITO"]

        # Mapeamento disciplinas
        mapa_disciplinas = (
            gabarito[["Questão", "Disciplina"]]
            .dropna()
            .groupby("Disciplina")["Questão"]
            .apply(list)
            .to_dict()
        )

        respostas_corr = corrigir_respostas(respostas, gabarito, mapa_disciplinas)

        # Ranking
        percentuais = []
        for i, row in respostas_corr.iterrows():
            acertos_tot = sum([row.get(f"Q{int(q)}_OK", False) for q in gabarito["Questão"]])
            percentuais.append(acertos_tot / len(gabarito))
        respostas_corr["Percentual"] = percentuais

        ranking_df = respostas_corr[["ID", "Nome", "Percentual"]].sort_values("Percentual", ascending=False).reset_index(drop=True)
        ranking_df["Posição"] = ranking_df.index + 1
        media_turma = ranking_df["Percentual"].mean() * 100

        # Médias por disciplina
        media_disciplinas = []
        for disc, questoes in mapa_disciplinas.items():
            acertos = []
            for _, row in respostas_corr.iterrows():
                acertos.append(sum([row.get(f"Q{int(q)}_OK", False) for q in questoes]) / len(questoes))
            media_disciplinas.append((disc, round(np.mean(acertos)*100, 1)))
        media_df = pd.DataFrame(media_disciplinas, columns=["Disciplina", "%"])

        # Gerar boletins
        zip_path = os.path.join(tmpdir, "boletins.zip")
        with zipfile.ZipFile(zip_path, "w") as zipf:
            for i, aluno in respostas_corr.iterrows():
                nome = aluno["Nome"].replace(" ", "_")
                posicao = int(ranking_df[ranking_df["ID"] == aluno["ID"]]["Posição"])
                percentual = aluno["Percentual"] * 100

                resultados = resultados_disciplina(aluno, mapa_disciplinas)
                df_boletim = pd.DataFrame(resultados, columns=["Disciplina", "Acertos", "Total", "%"])
                df_boletim["Média Turma"] = media_df["%"]
                df_boletim["Diferença"] = (df_boletim["%"] - media_df["%"]).round(1)

                # Gráficos
                barras, radar, dist, rank = gerar_graficos(nome, posicao, percentual, df_boletim, media_df, ranking_df, tmpdir)

                # PDF
                pdf = BoletimPDF()
                pdf.add_page()
                pdf.add_aluno_info(aluno["Nome"], posicao, percentual, media_turma)
                pdf.add_table(df_boletim)
                pdf.add_image(barras)
                pdf.add_image(radar)
                pdf.add_image(dist)
                pdf.add_image(rank)

                pdf_path = os.path.join(tmpdir, f"Boletim_{nome}.pdf")
                pdf.output(pdf_path)
                zipf.write(pdf_path, f"Boletim_{nome}.pdf")

        with open(zip_path, "rb") as f:
            st.download_button("📥 Baixar todos os boletins (ZIP)", f, "boletins.zip", "application/zip")
