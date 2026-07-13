from flask import Flask, send_file
import pandas as pd
import os
import zipfile
import uuid

from gerarrelatorios import (
    gerar_pdf_mensal,
    gerar_pdf_anual_layout,
    gerar_pdf_regional_layout,
    gerar_pdf_anual_sem_regional,
    gerar_pdf_metas_layout,
    gerar_excel,
    carregar_dados,
)

app = Flask(__name__)

@app.route("/")
def home():
    return "API de Relatórios funcionando 🚀"

@app.route("/gerar")
def gerar():
    # Usa carregar_dados() do próprio módulo, que já trata encoding,
    # BOM, colunas numéricas, datas e total_veiculos corretamente.
    df = carregar_dados()

    # 🔥 gerar todos os relatórios
    pdf0  = gerar_pdf_mensal(df)
    pdf1  = gerar_pdf_anual_layout(df)
    pdf2  = gerar_pdf_regional_layout(df)
    pdf3  = gerar_pdf_anual_sem_regional(df)
    pdf4  = gerar_pdf_metas_layout(df)
    excel = gerar_excel(df)

    # 🔹 criar zip com nome único
    zip_path = f"relatorios_{uuid.uuid4().hex}.zip"

    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zipf:
        for arquivo in [pdf0, pdf1, pdf2, pdf3, pdf4, excel]:
            if arquivo and os.path.exists(arquivo):
                zipf.write(arquivo, os.path.basename(arquivo))

    return send_file(
        zip_path,
        as_attachment=True,
        download_name="relatorios.zip",
        mimetype="application/zip"
    )

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port)