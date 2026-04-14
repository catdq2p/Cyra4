# TPCRA v3.0 Risk Assessment Dashboard

Streamlit dashboard built for the **TPCRA Questionnaire v3.0** Excel format with 5 sheets and 14 domains (A–N).

## Repo structure

```
├── app.py
├── requirements.txt
├── .streamlit/
│   └── config.toml
└── README.md
```

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

Opens at `http://localhost:8501`.

## Deploy to Streamlit Community Cloud

1. Push this repo to GitHub.
2. Go to [share.streamlit.io](https://share.streamlit.io) → **New app**.
3. Select repo, branch `main`, main file `app.py`.
4. Click **Deploy** — live in ~60 seconds.

Every `git push` to `main` triggers an automatic redeploy.

## Expected Excel format (v3.0)

| Sheet | Purpose |
|-------|---------|
| `Part 1` | Contact person, engagement info, application scope |
| `Part 2` | Security questionnaire (A–N domains, responses, risk tiers) |
| `Evidence` | Evidence checklist with status tracking |
| `Score Summary` | Rating reference (auto-calculated by dashboard) |
| `Selection` | Dropdown reference lists |

### Part 2 column structure

| Column | Content |
|--------|---------|
| A (`#`) | Question key e.g. `A.1`, `C.5.1`, `N.2.1` |
| B | Statement / Question text |
| C | Response: `Yes` / `No` / `Partial` / `N/A` |
| D | Other Information (remarks, evidence references) |
| E | Risk Tier: `Critical` / `High` / `Medium` / `Low` |
| F | Comments Required flag |

### Domain coverage (14 domains)

| Letter | Domain |
|--------|--------|
| A | Organizational Management |
| B | Human Resource Management |
| C | Infrastructure Security |
| D | Data Protection |
| E | Access Management |
| F | Application Security |
| G | System Security |
| H | Email Security |
| I | Mobile Devices |
| J | Incident Response |
| K | Cloud Services |
| L | Business Continuity |
| M | Supply Chain & Physical Security |
| N | AI & Emerging Technology Risk |

## Dashboard features

| Tab | Contents |
|-----|----------|
| **Overview** | Stacked bar by domain, donut chart, score bars, gaps-by-tier chart |
| **By domain** | Per-domain drill-down with question cards, tier badges, and free-text responses |
| **Gap analysis** | All No / Partial / Unanswered items — filtered by tier, response, and domain — sortable and exportable |
| **Evidence checklist** | Evidence items with submission status from the Evidence sheet |
| **Part 1 — Engagement** | Contact person and engagement information from Part 1 |

## Scoring

`Compliance Score = (Yes × 100 + Partial × 50) ÷ (Total − N/A) %`

| Score | Rating |
|-------|--------|
| ≥ 90% | ✅ Low |
| 70–89% | 🟡 Medium |
| 50–69% | 🟠 High |
| < 50% | 🔴 Critical |
