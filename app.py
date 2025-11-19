# -*- coding: utf-8 -*-
from flask import Flask, render_template, request, redirect, url_for, jsonify
from ProcFluxograma import gerar_fluxograma, processar_para_drawflow
import os

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs('static', exist_ok=True)

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        file = request.files.get("excel_file")
        if file:
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
            file.save(filepath)
            # Gera imagens estáticas (mantém funcionalidade antiga)
            gerar_fluxograma(filepath)
            return redirect(url_for("index"))
    return render_template("index.html")

@app.route("/api/fluxograma", methods=["POST"])
def api_fluxograma():
    """Novo endpoint para retornar dados JSON do fluxograma"""
    file = request.files.get("excel_file")
    if not file:
        return jsonify({"error": "Nenhum arquivo enviado"}), 400

    filepath = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
    file.save(filepath)

    try:
        dados = processar_para_drawflow(filepath)
        return jsonify(dados)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == "__main__":
    app.run(debug=True)
