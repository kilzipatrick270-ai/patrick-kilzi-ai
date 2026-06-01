#!/usr/bin/env python3
"""
Patrick Kilzi AI — Web App
Lance : python3 app.py
"""

import os
import json
import datetime
import anthropic
from flask import Flask, render_template, request, jsonify, send_file
from generate_presentation import generate_slide_structure, build_presentation
from financial_data import search_company_financials, format_financial_report

# Set API key
os.environ["ANTHROPIC_API_KEY"] = "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"

app = Flask(__name__)

EXPORT_DIR = os.path.join(os.path.dirname(__file__), "exports")
os.makedirs(EXPORT_DIR, exist_ok=True)


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/generate", methods=["POST"])
def generate():
    data = request.get_json()
    client      = data.get("client", "Client").strip() or "Client"
    description = data.get("description", "").strip()
    pres_type   = data.get("type", "Rapport BI")
    lang        = data.get("lang", "Français")

    if not description:
        return jsonify({"error": "Description vide."}), 400

    if not os.environ.get("ANTHROPIC_API_KEY"):
        return jsonify({"error": "ANTHROPIC_API_KEY non définie sur le serveur."}), 500

    project_info = f"Type de présentation : {pres_type}\nLangue : {lang}\n\n{description}"

    try:
        slide_data = generate_slide_structure(project_info, client)
    except Exception as e:
        return jsonify({"error": f"Erreur API Claude : {str(e)}"}), 500

    safe_client = client.replace(" ", "_").replace("/", "-")
    timestamp   = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    filename    = f"Patrick Kilzi_{safe_client}_{timestamp}.pptx"
    output_path = os.path.join(EXPORT_DIR, filename)

    try:
        build_presentation(slide_data, client, output_path)
    except Exception as e:
        return jsonify({"error": f"Erreur création PPTX : {str(e)}"}), 500

    size_kb = round(os.path.getsize(output_path) / 1024)
    num_slides = 4 + len(slide_data.get("sections", [])) * 2 + 4

    return jsonify({
        "title":    slide_data.get("title", "Présentation"),
        "client":   client,
        "filename": filename,
        "slides":   num_slides,
        "size_kb":  size_kb,
    })


@app.route("/search-financials", methods=["POST"])
def search_financials():
    data = request.get_json()
    company_name = data.get("company_name", "").strip()

    if not company_name:
        return jsonify({"error": "Nom d'entreprise vide."}), 400

    try:
        financial_data = search_company_financials(company_name)
    except Exception as e:
        return jsonify({"error": f"Erreur recherche : {str(e)}"}), 500

    if not financial_data.get("found"):
        return jsonify({
            "found": False,
            "error": financial_data.get("error", "Entreprise non trouvée"),
            "alternatives": financial_data.get("alternatives", [])
        }), 404

    # Format report
    report = format_financial_report(financial_data)

    return jsonify({
        "found": True,
        "company_name": financial_data.get("company_name"),
        "ticker": financial_data.get("ticker"),
        "report": report,
        "data": financial_data,
        "summary": {
            "latest_year": financial_data.get("years", [{}])[0].get("year"),
            "latest_revenue": financial_data.get("years", [{}])[0].get("revenue"),
            "latest_eps": financial_data.get("years", [{}])[0].get("eps"),
            "credit_rating": financial_data.get("credit_rating")
        }
    })


@app.route("/download/<filename>")
def download(filename):
    path = os.path.join(EXPORT_DIR, filename)
    if not os.path.exists(path):
        return "Fichier introuvable", 404
    return send_file(path, as_attachment=True,
                     download_name=filename,
                     mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation")


if __name__ == "__main__":
    print("\n" + "="*55)
    print("  INSIGHTER AI — Serveur Web")
    print("="*55)
    if not os.environ.get("ANTHROPIC_API_KEY"):
        print("  ⚠️  ANTHROPIC_API_KEY non définie !")
        print("  Exécute : export ANTHROPIC_API_KEY='ta-clé'")
    print("  Accès : http://localhost:5001")
    print("="*55 + "\n")
    app.run(debug=False, port=5002)
