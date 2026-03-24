#!/usr/bin/env python3
"""
convert_to_html.py
Le a planilha clientes_por_estado.xlsx e atualiza o index.html com os dados mais recentes.
Execute localmente ou via GitHub Actions.
"""

import pandas as pd
import json
import re
from datetime import datetime

XLSX_FILE = "clientes_por_estado.xlsx"
HTML_FILE  = "index.html"

def xlsx_to_json(xlsx_path):
    df = pd.read_excel(xlsx_path)
    data = {}
    for _, row in df.iterrows():
        sigla   = str(row["Sigla"]).strip()
        nome    = str(row["Estado"]).strip()
        cliente = str(row["Cliente"]).strip()
        if sigla not in data:
            data[sigla] = {"nome": nome, "clientes": []}
        data[sigla]["clientes"].append(cliente)
    return data

def inject_data_into_html(html_path, data, timestamp):
    with open(html_path, "r", encoding="utf-8") as f:
        content = f.read()

    new_json  = json.dumps(data, ensure_ascii=False, indent=2)
    new_block = f"const estadosData = {new_json};"
    content   = re.sub(
        r"const estadosData\s*=\s*\{.*?\};",
        new_block,
        content,
        flags=re.DOTALL
    )

    content = re.sub(
        r"(Ultima atualizacao:|Last update:).*?<",
        f"Ultima atualizacao: {timestamp}<",
        content
    )

    with open(html_path, "w", encoding="utf-8") as f:
        f.write(content)

    total = sum(len(v["clientes"]) for v in data.values())
    print(f"[OK] {html_path} atualizado com {total} clientes em {len(data)} estados.")

if __name__ == "__main__":
    ts   = datetime.now().strftime("%d/%m/%Y %H:%M")
    data = xlsx_to_json(XLSX_FILE)
    inject_data_into_html(HTML_FILE, data, ts)
    print("Concluido!")
