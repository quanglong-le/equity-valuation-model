# Equity Valuation Model – DCF + Comps
**Author:** Quang Long LE | Paris II Panthéon-Assas  
**Stack:** Python · pandas · numpy

---

## Objectif

Modèle de valorisation d'entreprise combinant deux approches complémentaires :
- **DCF (Discounted Cash Flow)** — valorisation intrinsèque par actualisation des flux futurs
- **Comps (Trading Multiples)** — valorisation relative via EV/EBITDA et P/E des pairs cotés

---

## Fonctionnalités

| Module | Description |
|---|---|
| `project_fcf()` | Projection des Free Cash Flows sur 5 ans |
| `dcf_valuation()` | Calcul WACC, Terminal Value, Equity Value/share |
| `comps_valuation()` | Valorisation par multiples EV/EBITDA et P/E |
| `football_field()` | Synthèse comparée de toutes les méthodes |
| Sensitivity table | Matrice Prix/action selon WACC × taux de croissance terminal |

---

## Méthodologie

### DCF
```
FCF = NOPAT + D&A - CAPEX - ΔBFR
Enterprise Value = Σ PV(FCF) + PV(Terminal Value)
Terminal Value = FCF_n × (1 + g) / (WACC - g)
Equity Value = EV - Net Debt
```

### Comps
```
EV (EV/EBITDA) = Médiane pairs × EBITDA société
P/E implied price = Médiane pairs × EPS
```

---

## Output exemple

```
════════════════════════════════════════════════════════════
  EQUITY VALUATION MODEL  |  Example Corp
════════════════════════════════════════════════════════════

  1. PROJECTED FREE CASH FLOWS  (M€)
  ────────────────────────────────────────────────────────
  Year  Revenue  EBITDA   EBIT   NOPAT    FCF
    Y1   1080.0   270.0  200.0   150.0   140.0
    Y2   1155.6   288.9  218.9   164.2   154.2
   ...

  2. DCF VALUATION  (M€)
  ────────────────────────────────────────────────────────
  PV of FCFs          :      673.4 M€
  Terminal Value (TV) :    2 134.8 M€
  PV of TV            :    1 386.7 M€
  Enterprise Value    :    2 060.1 M€
  Equity Value        :    1 860.1 M€
  ➜  Price per share  :      37.20 €

  3. SENSITIVITY – Price/share (€)  |  WACC × Terminal Growth
  ────────────────────────────────────────────────────────
         1.0%   2.0%   3.0%   4.0%
  7.0%   45.2   50.1   56.8   66.4
  8.0%   38.4   42.1   47.0   53.7
  9.0%   33.1   36.0   39.7   44.6
  ...

  5. FOOTBALL FIELD – Valuation Summary  (€/share)
  ────────────────────────────────────────────────────────
                        Low    Mid    High
  DCF                  33.5   37.2   40.9
  EV/EBITDA            28.6   31.8   34.9
  P/E                  42.9   47.6   52.4
  Implied range (avg)  28.6   38.9   52.4
```

---

## Lancer le modèle

```bash
# Cloner le repo
git clone https://github.com/quanglong-le/equity-valuation-model
cd equity-valuation-model

# Installer les dépendances
pip install pandas numpy

# Lancer
python equity_valuation.py
```

---

## Personnaliser

Modifiez les dictionnaires `company`, `dcf_params` et `comps` en tête de fichier :

```python
company = {
    "name"       : "Votre société",
    "revenue"    : 500,      # M€
    "ebitda"     : 120,
    "wacc"       : 0.09,
    ...
}
```

---

## Extensions possibles

- [ ] Import automatique des données financières via `yfinance`
- [ ] Dashboard interactif avec `plotly` / `dash`
- [ ] Export Excel des résultats (`openpyxl`)
- [ ] LBO model

---

## Compétences illustrées

`Python` · `pandas` · `numpy` · `DCF` · `Valorisation` · `Analyse financière` · `Modélisation quantitative`
