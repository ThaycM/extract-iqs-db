"""
Gera um relatório A4 paisagem em PDF, com texto alinhado à esquerda
e quebra automática de colunas longas, replicando o mesmo filtro do Power BI.

Requisitos extra:
    pip install pandas numpy reportlab
"""

import sys
import datetime as dt
from pathlib import Path
import numpy as np
import pandas as pd
from time import sleep
# --------- ReportLab -----------
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors
from reportlab.platypus import (
    SimpleDocTemplate,
    Table,
    TableStyle,
    Paragraph,
    PageBreak,
)
from reportlab.lib.styles import getSampleStyleSheet


# ------------- CONFIG -------------- #
today=dt.date.today()
iso_year, iso_week,_=today.isocalendar()

INPUT_PATH  = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(
    r"C:\Users\00071228\OneDrive - ENERCON\QA Team - Follow up - Qualidade - databases\actions_db.xlsx"
)
OUTPUT_PATH = Path(sys.argv[2]) if len(sys.argv) > 2 else Path(
    f"C:/Users/00071228/OneDrive - ENERCON/QA Team - Follow up - Qualidade - databases/PDFs_Reports/actions_report_W{iso_week:02d}.pdf"
    )

# ---------- TRANSFORMAÇÕES ---------- #
def excel_serial_to_date(x):
    if pd.isna(x):
        return pd.NaT
    if isinstance(x, (int, float, np.integer, np.floating)):
        return (dt.datetime(1899, 12, 30) + dt.timedelta(days=int(x))).date()
    return pd.to_datetime(x, errors="coerce").date()


def transformar_dataframe(df,type="all"):
    # Mesmas regras que você já usava  -------------------------
    cols_orig = [
        "Status", "Title", "Description", "Planned end", "Actual end", "Created on",
        "Linked with", "Read", "Editor", "Creator", "Activated at", "% complete",
        "Action result", "Actual begin", "Completion status (bar)", "Creator (position)",
        "Organization - editor", "Partial task of", "Planned begin", "Version date",
        "MId",
    ]
    df = df[cols_orig].copy()

    rename_map = {
        "Title": "Título",
        "Description": "Descrição",
        "Planned end": "Prazo Planejado",
        "Actual end": "Data de Conclusão",
        "Created on": "Data de Criação",
        "Linked with": "Vinculado a",
        "Read": "Data de Leitura",
        "Editor": "Editor",
        "Creator": "Criador",
        "Activated at": "Data de Ativação",
        "% complete": "% Concluído",
        "Action result": "Resultado da Ação",
        "Actual begin": "Início Real",
        "Completion status (bar)": "Progresso Visual (%)",
        "Creator (position)": "Cargo do Criador",
        "Organization - editor": "Organização do Editor",
        "Partial task of": "Subtarefa de",
        "Planned begin": "Início Planejado",
        "Version date": "Data da Versão",
        "MId": "ID",
    }
    df.rename(columns=rename_map, inplace=True)

    df = df[~df["Criador"].isin(["Costa, Pedro Alexandre", "Krause, Stefan"])]
    df = df[df["Status"].isin(["Completed", "In progress", "Unprocessed"])]
    df = df[df["Descrição"].notna()]
    df = df[~df["Descrição"].str.lower().str.startswith("complaint process")]
    df = df[~df["Título"].str.startswith("RKM", na=False)]

    date_cols = [
        "Prazo Planejado", "Data de Conclusão", "Data de Criação", "Data de Leitura",
        "Data de Ativação", "Início Real", "Início Planejado", "Data da Versão",
    ]
    for c in date_cols:
        df[c] = df[c].apply(excel_serial_to_date)

    today = dt.date.today()

    df["Dias em aberto"] = df.apply(
        lambda r: (
            (r["Data de Conclusão"] if pd.notna(r["Data de Conclusão"]) else today)
            - r["Data de Criação"]
        ).days
        if pd.notna(r["Data de Criação"])
        else None,
        axis=1,
    )

    df["Status de Prazo"] = df.apply(
        lambda r: "Concluída"
        if pd.notna(r["Data de Conclusão"])
        else (
            "Atrasada"
            if pd.notna(r["Prazo Planejado"]) and r["Prazo Planejado"] < today
            else "Dentro do Prazo"
        ),
        axis=1,
    )

    # Mostra só as linhas em aberto (já faz parte do seu script)
    df = df[df["Status"].isin(["In progress", "Unprocessed"])]
    # Filtrar as colunas que são da ENERCON:
    if type=="all":
        mascara_linked=(
        df['Vinculado a']
        .fillna('')
        .str.contains(
            r'ENERCON GmbH \(91200000\)| General actions',
            case=False,
            regex=True
            )
                        )
    elif type=="kaizen":
        mascara_linked=(
        df['Vinculado a']
        .fillna('')
        .str.contains(
            r'General actions',
            case=False,
            regex=True
            )
                        )
    elif type=="rkm":
        mascara_linked=(
        df['Vinculado a']
        .fillna('')
        .str.contains(
            r'ENERCON GmbH \(91200000\)',
            case=False,
            regex=True
            )
                        )
    else:
        mascara_linked=(
        df['Vinculado a']
        .fillna('')
        .str.contains(
            r'ENERCON GmbH \(91200000\)| General actions',
            case=False,
            regex=True
            )
                        )         
               
    df=df[mascara_linked]

    criadores_ok = [
    'Gaifem, Sonia',
    'Marinho, Thayc',
    'Novo, Vitor',
    'Pinto, Ricardo',
    'Pita, Liliana',
    'Sousa, Monica',   
    'Vieira, Maria'
    ]
    df= (
        df[df['Criador'].str.strip().isin(criadores_ok)]
        .copy()
    )
    
    # Colunas finais a mostrar
    df = df[
        [
            "ID",
            "Título",
            "Descrição",
            "Editor",
            "Data de Criação",
            "Prazo Planejado",
            "Vinculado a",
            "Status de Prazo",
        ]
    ]

    # Wrap manual de texto (ReportLab já faz, mas melhor limitar tamanho bruto)
    wrap = lambda s, width: "<br/>".join(
        [s[i : i + width] for i in range(0, len(s), width)]
    )
    df["Título"] = df["Título"].astype(str).apply(lambda s: wrap(s, 40))
    df["Descrição"] = df["Descrição"].astype(str).apply(lambda s: wrap(s, 60))
    # Limpeza dos caracteres estranhos dentro dos textos
    df["Título"]=(df["Título"].
                  str.replace(r'(?i)_x000d_','',regex=True).
                  str.replace(r'[\r\n]+','',regex=True).
                  str.replace(r'\s{2,}','',regex=True).
                  str.strip()
                  )
    df["Descrição"]=(df["Descrição"].
                  str.replace(r'(?i)_x000d_','',regex=True).
                  str.replace(r'[\r\n]+','',regex=True).
                  str.replace(r'\s{2,}','',regex=True).
                  str.strip()
                  )
    # Organizar dataframe pela data de criação
    df= df.sort_values(by="Data de Criação", ascending= True)
    df=df.reset_index(drop=True)
    return df


# ---------- GERADOR DE PDF (ReportLab) ---------- #
def df_to_pdf(df: pd.DataFrame, pdf_path: Path):
    doc = SimpleDocTemplate(
        str(pdf_path),
        pagesize=landscape(A4),
        leftMargin=18,
        rightMargin=18,
        topMargin=18,
        bottomMargin=18,
    )
    styles = getSampleStyleSheet()
    base = styles["Normal"]
    base.fontSize = 7
    base.leading = 9
    base.spaceAfter = 0
    base.alignment = 0  # LEFT

    # Conversão DataFrame → lista de listas, usando Paragraph para quebrar texto
    data = [list(df.columns)]
    for _, row in df.iterrows():
        data.append([Paragraph(str(row[c]), base) for c in df.columns])

    # Ajuste de larguras (em points; 1 pt ≈ 0.35 mm) — some até ~800 pt para caber A4 landscape
    col_widths = [35, 90, 230, 65, 60, 60, 80, 60]

    table = Table(data, colWidths=col_widths, repeatRows=1)

    table.setStyle(
        TableStyle(
            [
                ("ALIGN", (0, 0), (-1, -1), "LEFT"),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
                ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
                ("FONTSIZE", (0, 0), (-1, -1), 7),
            ]
        )
    )

    doc.build([table])

# ------------------------------ MAIN ------------------------------ #
def main():
    print(f"📄 Lendo: {INPUT_PATH}")
    df_raw = pd.read_excel(INPUT_PATH, sheet_name="Sheet")
    df_tratado = transformar_dataframe(df_raw)
    df_to_pdf(df_tratado, OUTPUT_PATH)
    print(f"✅ PDF gerado em: {OUTPUT_PATH.resolve()}")


if __name__ == "__main__":
    main()
