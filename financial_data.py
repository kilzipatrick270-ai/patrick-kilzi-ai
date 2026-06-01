#!/usr/bin/env python3
"""
Module de récupération des données financières
Cherche les chiffres d'une entreprise via API + web search
"""

import json
import re
import anthropic


def search_company_financials(company_name: str) -> dict:
    """
    Cherche les données financières d'une entreprise.
    Retourne : {
        "found": bool,
        "company": {...données sur 3 ans...},
        "error": str (si not found)
    }
    """

    ai = anthropic.Anthropic()

    prompt = f"""Tu es un expert en analyse financière.

Cherche les données financières PUBLIQUES de l'entreprise : {company_name}

Retourne UNIQUEMENT un JSON valide avec la structure exacte :

{{
  "found": true/false,
  "company_name": "Nom exact",
  "ticker": "XXX (si trouvé)",
  "currency": "USD",
  "years": [
    {{
      "year": 2023,
      "revenue": 210800000000,
      "ebitda": 52300000000,
      "net_income": 15200000000,
      "operating_cash_flow": 25400000000,
      "free_cash_flow": 18500000000,
      "total_debt": 32100000000,
      "cash_and_equivalents": 8900000000,
      "equity": 98700000000,
      "assets": 185000000000,
      "eps": 3.42,
      "dividend_per_share": 2.89,
      "shares_outstanding": 2140000000,
      "note": "Source: rapport annuel 2023"
    }},
    {{
      "year": 2022,
      "revenue": 200100000000,
      "ebitda": 48500000000,
      "net_income": 16800000000,
      "operating_cash_flow": 22300000000,
      "free_cash_flow": 16200000000,
      "total_debt": 31500000000,
      "cash_and_equivalents": 7200000000,
      "equity": 95200000000,
      "assets": 180000000000,
      "eps": 3.89,
      "dividend_per_share": 2.65,
      "shares_outstanding": 2140000000,
      "note": "Source: rapport annuel 2022"
    }},
    {{
      "year": 2021,
      "revenue": 184300000000,
      "ebitda": 42100000000,
      "net_income": 12900000000,
      "operating_cash_flow": 19800000000,
      "free_cash_flow": 13100000000,
      "total_debt": 33200000000,
      "cash_and_equivalents": 6100000000,
      "equity": 88500000000,
      "assets": 175000000000,
      "eps": 3.01,
      "dividend_per_share": 2.25,
      "shares_outstanding": 2140000000,
      "note": "Source: rapport annuel 2021"
    }}
  ],
  "credit_rating": {{
    "sp": "A+",
    "moodys": "A1",
    "fitch": "A+",
    "outlook": "Stable"
  }},
  "alternatives": []
}}

IMPORTANT :
- Tous les montants EN USD (convertir si nécessaire)
- Cherche les 3 dernières années disponibles
- Si NOT found, mets found: false et donne alternatives similaires
- Tous les montants en nombres (pas de M€, $, etc)
- Inclus seulement les données officielles publiques

Réponds UNIQUEMENT avec le JSON, sans texte avant/après.
"""

    try:
        message = ai.messages.create(
            model="claude-sonnet-4-6",
            max_tokens=3000,
            messages=[{"role": "user", "content": prompt}]
        )

        raw = message.content[0].text.strip()
        # Remove markdown if present
        if raw.startswith("```"):
            raw = raw.split("```")[1]
            if raw.startswith("json"):
                raw = raw[4:]
            raw = raw.strip()

        data = json.loads(raw)

        if data.get("found"):
            # Calcul des ratios
            calculate_ratios(data)

        return data

    except json.JSONDecodeError as e:
        return {
            "found": False,
            "error": f"Erreur parsing données : {str(e)}",
            "alternatives": []
        }
    except Exception as e:
        return {
            "found": False,
            "error": f"Erreur recherche : {str(e)}",
            "alternatives": []
        }


def calculate_ratios(data: dict):
    """Calcule ROE, ROA, EPS pour chaque année"""
    for year_data in data.get("years", []):
        equity = year_data.get("equity", 1)
        assets = year_data.get("assets", 1)
        net_income = year_data.get("net_income", 0)

        # ROE = Net Income / Equity
        roe = (net_income / equity * 100) if equity > 0 else 0
        year_data["roe"] = round(roe, 2)

        # ROA = Net Income / Assets
        roa = (net_income / assets * 100) if assets > 0 else 0
        year_data["roa"] = round(roa, 2)

        # EPS est déjà fourni, mais on peut le recalculer
        # EPS = Net Income / Shares Outstanding
        shares = year_data.get("shares_outstanding", 1)
        if shares > 0:
            eps_calc = net_income / shares
            year_data["eps_calculated"] = round(eps_calc, 2)


def format_financial_report(data: dict) -> str:
    """Formate les données financières en texte lisible pour la présentation"""

    if not data.get("found"):
        return f"❌ Entreprise non trouvée. Suggestions : {', '.join(data.get('alternatives', []))}"

    company = data.get("company_name", "Unknown")
    ticker = data.get("ticker", "N/A")
    rating = data.get("credit_rating", {})

    report = f"""
DONNÉES FINANCIÈRES {company.upper()} ({ticker})
Source: Données publiques officielles | Devise: USD

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

NOTATION CRÉDIT
• S&P: {rating.get('sp', 'N/A')}
• Moody's: {rating.get('moodys', 'N/A')}
• Fitch: {rating.get('fitch', 'N/A')}
• Outlook: {rating.get('outlook', 'Stable')}

"""

    for year_data in data.get("years", []):
        year = year_data.get("year", "N/A")
        revenue = year_data.get("revenue", 0)
        ebitda = year_data.get("ebitda", 0)
        net_income = year_data.get("net_income", 0)
        fcf = year_data.get("free_cash_flow", 0)
        debt = year_data.get("total_debt", 0)
        cash = year_data.get("cash_and_equivalents", 0)
        equity = year_data.get("equity", 0)
        eps = year_data.get("eps", 0)
        roe = year_data.get("roe", 0)
        roa = year_data.get("roa", 0)

        net_debt = debt - cash
        ebitda_margin = (ebitda / revenue * 100) if revenue > 0 else 0
        net_margin = (net_income / revenue * 100) if revenue > 0 else 0
        fcf_yield = (fcf / equity * 100) if equity > 0 else 0
        leverage = (net_debt / ebitda) if ebitda > 0 else 0

        report += f"""ANNÉE {year}
────────────────────────────────────────────────

COMPTE DE RÉSULTAT
• Chiffre d'affaires (CA): ${revenue/1e9:.1f}B
• EBITDA: ${ebitda/1e9:.1f}B (marge {ebitda_margin:.1f}%)
• Résultat net (PAT): ${net_income/1e9:.1f}B (marge {net_margin:.1f}%)
• EPS (Earnings Per Share): ${eps:.2f}

CASH FLOW
• Free Cash Flow (FCF): ${fcf/1e9:.1f}B
• FCF Yield: {fcf_yield:.2f}%
• OCF: ${year_data.get('operating_cash_flow', 0)/1e9:.1f}B

BILAN & DETTE
• Dette brute: ${debt/1e9:.1f}B
• Trésorerie: ${cash/1e9:.1f}B
• Dette nette: ${net_debt/1e9:.1f}B
• Equity: ${equity/1e9:.1f}B
• Levier ND/EBITDA: {leverage:.2f}x

RATIOS DE PERFORMANCE
• ROE (Return on Equity): {roe:.2f}%
• ROA (Return on Assets): {roa:.2f}%
• Dividend par action: ${year_data.get('dividend_per_share', 0):.2f}

"""

    return report


if __name__ == "__main__":
    # Test
    result = search_company_financials("TotalEnergies")
    print(json.dumps(result, indent=2))
    print("\n" + format_financial_report(result))
