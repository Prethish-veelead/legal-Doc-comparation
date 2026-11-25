# Clause Standard Builder (Part A)

Turns a contract template stored in SharePoint into a Draft clause standard, using Azure OpenAI.
The document is read in memory and never stored. Only metadata is saved to SharePoint.

## How it works

1. `POST /api/process` receives a SharePoint document link, a contract type and (optionally) our role.
2. The file is downloaded into memory through Microsoft Graph (app registration, client secret).
3. Text is extracted (Word .docx, PDF, .txt, .md / README) and split into clauses.
4. Each clause is sent on its own to Azure OpenAI (`gpt-4o-mini`), with strict JSON output limited to the allowed categories, departments and risk levels.
5. Every answer is checked: values must be on the lists, and any number in the preferred, fallback or red line positions must appear in the clause. Problems flag the clause for Legal. Unusable answers are drafted by a simple keyword method instead.
6. Results are saved to SharePoint:
   - `LDC_ProcessingRuns`: one row per run (link, file name, fingerprint, counts, tokens, cost, result JSON)
   - `LDC_ContractStandards`: one Draft standard (also fills `LDC_DraftClausesJson` in the existing app's format)
   - `LDC_Clauses`: one row per clause
7. `POST /api/publish` rebuilds the rules from the Clauses list, so edits made there are kept, then sets the standard to Published.

The same file processed again (same SHA-256 fingerprint) returns the saved result without calling the AI, unless `force` is true.

## Using it

Open `https://func-ldc-clause-dev-oter.azurewebsites.net/api/app?code=<function key>`.
Get the key from the Azure portal: Function App → App keys → `default`.

API:

```http
POST /api/process?code=<key>
{ "documentUrl": "https://veeleaddev.sharepoint.com/sites/LegalDocComparison/LDC_ContractDocuments/<file>",
  "contractType": "Master services", "ourRole": "Customer", "force": false }

POST /api/publish?code=<key>
{ "standardId": 11 }
```

## Azure resources (Azure Learning subscription, East US 2)

| Resource | Name |
|---|---|
| Resource group | `rg-legal-doc-compare-dev` |
| Azure OpenAI | `aoai-legal-doc-compare-dev` (deployment `gpt-4o-mini`) |
| Function App (Flex Consumption, Node 24) | `func-ldc-clause-dev-oter` |
| Storage account (required by Functions) | `stldcclauseoter` |

## Local development

```bash
npm install
cp .env.example .env              # fill in values
npm run setup-lists               # creates LDC_Clauses and LDC_ProcessingRuns (safe to re-run)
npm run run-local -- "<document link>" "Master services" "Customer" [--force]
func start                        # needs local.settings.json with the same values as .env
```

Deploy: `az account set --subscription "Azure Learning"` then `func azure functionapp publish func-ldc-clause-dev-oter`.

## Settings

`AZURE_TENANT_ID`, `AZURE_CLIENT_ID`, `AZURE_CLIENT_SECRET` (SharePoint app registration),
`SHAREPOINT_SITE_URL`, `SHAREPOINT_LIBRARY`, `AZURE_OPENAI_ENDPOINT`, `AZURE_OPENAI_API_KEY`,
`AZURE_OPENAI_DEPLOYMENT`, `AZURE_OPENAI_API_VERSION`. Optional: `MAX_FILE_MB` (20), `AI_CONCURRENCY` (3).

The lists the AI chooses from are maintained in SharePoint:
departments in `LDC_ReviewCategories` (active rows), categories in `LDC_AppSettings` → `clauseCategories`.

## Part B: Contract check (local testing, not deployed yet)

`POST /api/compare` checks the other party's draft against a **published** standard. It is read-only: nothing is saved yet.

```http
POST /api/compare?code=<key>
{ "documentUrl": "https://veeleaddev.sharepoint.com/sites/LegalDocComparison/LDC_ContractDocuments/<their draft>",
  "standardId": 1, "ourRole": "Customer" }
```

1. The draft is read in memory and split into clauses (same code as Part A).
2. For each rule in the standard, the 3 most similar clauses are picked by word overlap and sent to Azure OpenAI
   (`AZURE_OPENAI_COMPARE_DEPLOYMENT`, `gpt-4.1-mini`). The AI writes its reasoning, then one verdict:
   `meets_preferred`, `within_fallback`, `outside_fallback`, `crosses_red_line`, `missing` or `unclear`.
3. Checks: the quote must appear word for word in the clause, confidence must be at least 0.6, and a red line whose
   reasoning shows an equal value or "not above/under" is flagged for review.
4. Findings are routed to a team: the rule's sign-off team if it is a review category, otherwise the Classification Rules list.
   Clauses that match no rule are listed under `notInStandard`.
5. `findingsJson` uses the same format as `LDC_Comparisons.LDC_FindingsJson`.
6. `summary.riskScore` (0-100) and `summary.riskLevel` use the SPFx web part's own formula and bands.

The SPFx web part calls this endpoint with a function key for `compare` only. See `PART-B-SPFX-HANDOVER.md`.

Local testing:

```bash
npm run run-compare -- "<their document link>" 1 "Customer"
npm run run-compare -- --file test/fixtures/counterparty-msa-bad.md 1 "Customer"
node test/compare-check.js 3      # 3 drafts with known answers, 3 runs each (calls Azure OpenAI)
```

## Limits

- Old .doc files and scanned PDFs without text are refused with a message.
- Links must be inside the LegalDocComparison site.
- Files up to 20 MB.
