# Quote Analyzer — Internal Benefits Dashboard

An internal tool for benefits brokers to upload carrier quote files (PDF / Excel / CSV), auto-extract plan data, score and rank plans, and export client-ready PowerPoint or Excel reports — all from a single-page dashboard.

---

## Features

| Feature | Details |
|---------|---------|
| **File Upload** | Drag-and-drop or browse; PDF, XLSX, XLS, CSV; up to 50 MB per file |
| **Auto Extraction** | Regex + Excel-column parsing for 25+ benefit fields per plan |
| **Scoring Engine** | Premium efficiency (40%), risk protection (30%), copay usability (20%), network (10%) |
| **Editable Table** | Click any cell in the extracted-plans table to edit values inline |
| **Recommendations** | Top-3 scored plans with score bars and natural-language rationale |
| **PPTX Export** | 8-slide branded deck: title, summary, top 3, comparison table, premiums, appendix |
| **XLSX Export** | Two-sheet workbook: full Data sheet + Summary with census totals |
| **Token Auth** | Simple `X-API-Token` header; configurable from the UI |
| **Embeddable** | Vanilla JS frontend — works in any `<iframe>` or WordPress page |

---

## Repository Layout

```
Quote-Analyzer/
├── backend/
│   ├── package.json
│   └── server.js          # Express API (Node 18+)
├── frontend/
│   ├── index.html         # Single-page dashboard
│   ├── styles.css         # Professional CSS with custom properties
│   └── app.js             # Vanilla JS — no framework dependencies
├── .gitignore
└── README.md
```

---

## Local Development

### Prerequisites
- Node.js 18+ and npm
- Modern browser (Chrome, Firefox, Edge, Safari)

### 1 — Start the backend

```bash
cd backend
npm install
npm start        # production
# or
npm run dev      # nodemon hot-reload (install nodemon first: npm i -D nodemon)
```

The API listens on **http://localhost:3001** by default.  
Override with `PORT=8080 npm start`.

Change the API token with `API_TOKEN=mysecret npm start`.

### 2 — Open the frontend

Open `frontend/index.html` directly in your browser, **or** serve it:

```bash
# Python (no install needed)
cd frontend && python3 -m http.server 8080

# Node (npx)
cd frontend && npx serve .
```

In the dashboard's **Case Information** card, set:
- **API URL** → `http://localhost:3001`
- **API Token** → `internal-token-2024` (or whatever you set via `API_TOKEN`)

### 3 — Health check

```bash
curl http://localhost:3001/health
# {"status":"ok"}
```

---

## API Reference

All endpoints (except `GET /health`) require the `X-API-Token` header.

| Method | Endpoint | Description |
|--------|----------|-------------|
| `GET`  | `/health` | Liveness check |
| `POST` | `/upload` | Upload files (multipart `files[]`); returns `{ caseId, files }` |
| `POST` | `/parse` | Extract plans from uploaded files; returns `{ caseId, plans, warnings }` |
| `POST` | `/recommend` | Score and rank plans; returns `{ recommendations, allPlans }` |
| `POST` | `/export/pptx` | Generate PPTX; returns binary file |
| `POST` | `/export/xlsx` | Generate XLSX; returns binary file |

### Plan Object Schema

```js
{
  id, carrier, planName, planCode, networkType, metalLevel,
  deductibleIndividual, deductibleFamily, oopMaxIndividual, oopMaxFamily,
  coinsurance, copayPCP, copaySpecialist, copayUrgentCare, copayER,
  rxDeductible, rxTier1, rxTier2, rxTier3,
  premiumEE, premiumES, premiumEC, premiumEF,
  effectiveDate, ratingArea, underwritingNotes,
  extractionConfidence,   // 0.0–1.0
  sourceFile
}
```

---

## Carrier File Formatting Tips

For best extraction confidence:

### PDFs
- Include clear field labels: `Deductible: $1,500`, `PCP Copay: $25`, `OOP Maximum: $6,000`
- Put the plan name on its own line (e.g. `Anthem Gold PPO 1500`)
- Include keywords: `HMO`, `PPO`, `EPO`, `HDHP`, `Platinum`, `Gold`, `Silver`, `Bronze`
- Premium tables should label tiers: `EE`, `ES`, `EC`, `EF` or `Employee Only`, `Employee + Spouse`, etc.

### Excel / CSV
- **Row-per-plan layout** (recommended): one row per plan, columns named after fields  
  (`Carrier`, `Plan Name`, `Network Type`, `Deductible`, `OOP Max`, `PCP Copay`, `EE Premium`, …)
- **Column-per-plan layout**: first column = field labels, each subsequent column = plan values
- Column headers are matched case-insensitively; partial matches work (`"Ded"` matches `"Deductible"`)

---

## Scoring Weights

| Category | Weight | Logic |
|----------|--------|-------|
| Premium Efficiency | 40% | Lower monthly total for census → higher score; normalized min–max |
| Risk Protection | 30% | Lower individual deductible + OOP max → higher; $0 = 1.0, ≥$15k = 0 |
| Copay Usability | 20% | Lower PCP copay + copay-first model (no deductible for office visits) |
| Network Breadth | 10% | PPO=1.0, EPO=0.8, HMO=0.7, HDHP/HSA=0.5 |

---

## Environment Variables

| Variable | Default | Description |
|----------|---------|-------------|
| `PORT` | `3001` | API listen port |
| `API_TOKEN` | `internal-token-2024` | Bearer token for `X-API-Token` header |

---

## WordPress Embedding

The frontend is fully self-contained (vanilla JS, single HTML file). To embed in WordPress:

1. **Deploy the backend** to any Node.js host (Railway, Render, Fly.io, EC2, etc.)
2. **Host the frontend** files on any static host (S3, Netlify, same server)
3. Embed via iframe in a WordPress page:

```html
<iframe
  src="https://your-host.example.com/quote-analyzer/"
  width="100%"
  height="900"
  frameborder="0"
  style="border:none;border-radius:12px"
></iframe>
```

4. Set the **API URL** field in the dashboard to your deployed backend URL.

Alternatively, install the [WP iFrame](https://wordpress.org/plugins/iframe/) plugin and use `[iframe src="https://..."]`.

---

## Security Notes

- The `API_TOKEN` is a shared secret — rotate it regularly in production.
- Add HTTPS (TLS) in production; the backend does not terminate TLS itself.
- The in-memory case store is process-local and not persistent; restart clears all data.
- For multi-instance or persistent storage, replace `caseStore` (Map) with Redis or a database.
- `pdf-parse` runs in-process; consider sandboxing for untrusted PDFs in production.

---

## License

Internal use only — not for public distribution.
tool to analyze and organize quotes for account managers to use internally
