# Part B: AI contract check, handover for the SPFx developer

Part B checks the counterparty's draft against a **published contract standard** and returns findings in the same
format the web part already saves to `LDC_Comparisons.LDC_FindingsJson`. The web part keeps doing everything else
(upload, wizard, saving, review tasks). Part B only replaces the "deterministic-rules" analysis step when the
baseline is a contract standard.

## When to call it

New comparison → Baseline tab **Contract standard** → a Published standard is selected → Run comparison.
After the draft is uploaded to `LDC_ContractDocuments`, call Part B instead of the local rules engine.
(Baseline tab **Your own file** stays on the existing redline engine.)

## Request

```http
POST https://func-ldc-clause-dev-oter.azurewebsites.net/api/compare?code=<function key>
Content-Type: application/json

{
  "documentUrl": "https://veeleaddev.sharepoint.com/sites/LegalDocComparison/LDC_ContractDocuments/<file>.docx",
  "standardId": 1,
  "ourRole": "Customer"
}
```

| Field | Rules |
|---|---|
| `documentUrl` | Full link to the uploaded draft. Must be inside the LegalDocComparison site. Word (.docx), PDF, .txt or .md, up to 20 MB. The **current** version of the file is checked. |
| `standardId` | `LDC_ContractStandards` item ID. Must be **Published**. |
| `ourRole` | Optional, up to 40 characters (for example "Customer" or "Supplier"). |

No other fields are accepted. Use `fetch` (not `SPHttpClient`), because the function is outside SharePoint.
Typical time: 6-10 seconds. Set a client timeout of 60 seconds.

## Response (200)

Save these fields to the comparison row:

| Response field | Save to `LDC_Comparisons` column |
|---|---|
| `findingsJson` | `LDC_FindingsJson` (JSON.stringify) |
| `summary.riskScore` (0-100) | `LDC_RiskScore` |
| `summary.riskLevel` (Low, Medium, High, Critical) | `LDC_RiskLevel` |
| `summary.aiSummary` | `LDC_AiSummary` |
| `run.runId` | `LDC_CorrelationId` |
| `standard.version` | `LDC_StandardVersion` |

The risk score uses the web part's own formula (critical 30, high 20, medium 10, low 4, +10 for a red line;
`100 × (1 − e^(−points/70))`) and the same bands, so approval-stage thresholds keep working.

`findingsJson.findings[]` has every field the web part already reads (`id, clause, target, title, severity,
ruleCode, category, redLine, confidence, proposed, required, description, why, routedTeam, routeNote, activity`)
plus five extra fields that can be shown or ignored:

| Extra field | Meaning |
|---|---|
| `verdict` | `meets_preferred`, `within_fallback`, `outside_fallback`, `crosses_red_line`, `missing`, `unclear` |
| `categoryKey` | Review category key, for `LDC_ChangeReviewTasks` routing |
| `aiReasoning` | Short reasoning from the AI |
| `needsReview` | `true` when the system is not sure the AI is right |
| `reviewReasons` | Why it needs review |

Other sections (`run`, `document`, `checks`, `notInStandard`) are for display or audit and need not be saved.
`notInStandard` lists the draft's clauses that no rule covers, with a suggested team.

## Errors

| Status | Meaning | Show the user |
|---|---|---|
| 400 | Bad request (missing or extra fields, bad link) | The `error` message |
| 404 | Standard or document not found | The `error` message |
| 409 | Standard is not published | The `error` message |
| 413 / 415 / 422 | File too large, unsupported type, or no text (scanned PDF) | The `error` message |
| 401 | Wrong or missing key | "AI analysis is not available" and fall back to the rules engine |
| 500 / network | Service problem | Same as 401 |

All errors return `{ "error": "<plain message>" }`.

## The function key

- The key is a **function-level key for `compare` only**. It cannot call any other endpoint.
- Keep it in the web part property pane or the `LDC_AppSettings` list, not in source code.
- Anyone who can open the page can see the key in the browser's network tab. If it leaks, the administrator
  rotates it in the Azure portal (Function App → Functions → compare → Function keys) and updates the setting.
- Calls are only accepted from browsers on `https://veeleaddev.sharepoint.com` (CORS).

## Privacy

The function reads the draft into memory, checks it and drops it. It stores nothing: no file, no text.
