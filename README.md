# Compliance Tools: MSME Verification & Struck-Off Companies

A small web app with two tools for checking vendors:

1. **MSME Search & Verification** (`/`): Udyam registration status by PAN or Udyam number.
2. **Struck-Off Companies** (`/struck-off`): finds suppliers that MCA has struck off, or that are
   under process of striking off, dormant or otherwise not active, from their PAN / GSTIN / CIN.

**Windows:** double-click `Start Compliance Tools.bat`. It installs what the tool needs and opens it
in your browser. You need Python installed first (python.org, tick "Add python.exe to PATH").

---

## 1. MSME Search & Verification

A web tool for checking the MSME (Udyam) registration status of vendors, so you can comply
with the **MSMED Act, 2006** (45-day payment rule, half-yearly MSME Form-1) and
**Section 43B(h) of the Income-tax Act**.

## Features

| What it does | Details |
|---|---|
| **Search by PAN** | Checks one PAN for a Udyam registration |
| **Search by Udyam number** | Checks one `UDYAM-XX-00-0000000` number |
| **Bulk verification** | Upload an Excel/CSV vendor list (PAN and Udyam numbers can be mixed) or paste a list |
| **Details returned** | Registration status, name of enterprise, type of enterprise (Micro/Small/Medium), major activity, type of organisation, date of incorporation, Udyam registration date, state/district |
| **Compliance flag** | Shows whether the 45-day rule and Sec 43B(h) apply (Micro/Small, not traders) |
| **Excel export** | Formatted report with colour-coded statuses and a summary sheet |
| **Format validation** | Bad PAN/Udyam formats are flagged before any lookup is made |
| **Result cache** | Repeat lookups come from a local SQLite cache (default 30 days), so you don't pay for the same paid-API call twice |

## Quick start

```bash
pip install -r requirements.txt
python -m msme_verifier.app          # http://127.0.0.1:5000
```

It starts in **demo mode**, which returns made-up sample data so you can try the screens.
To check it against the bundled sample master file:

```bash
MSME_PROVIDER=local MSME_MASTER_FILE=sample_data/msme_master.csv python -m msme_verifier.app
```

Then upload `sample_data/vendors_to_verify.csv` on the **Bulk Verification** tab.

## Data sources (providers)

The government Udyam portal has no free public API, and its verification page is protected by a
CAPTCHA. Commercial tools like this one get their data from a **licensed data provider**
(KYC/verification API vendors that offer "Udyam by PAN" and "Udyam verification" APIs). You pick
the source with `MSME_PROVIDER`:

| `MSME_PROVIDER` | Use |
|---|---|
| `demo` (default) | Made-up, repeatable data for demos and training |
| `local` | Looks numbers up in your own vendor master (CSV/XLSX), e.g. built from Udyam certificates your vendors gave you. Set `MSME_MASTER_FILE`. |
| `http` | Calls any REST verification API; set it up with the variables below |

### Connecting a verification API (`http`)

```bash
export MSME_PROVIDER=http
export MSME_API_URL_PAN="https://api.your-vendor.com/udyam/pan?pan={id}"
export MSME_API_URL_UDYAM="https://api.your-vendor.com/udyam/verify?urn={id}"
export MSME_API_METHOD=GET                      # or POST (sends {"id_number": "<id>"})
export MSME_API_KEY="Bearer <your-token>"
export MSME_API_KEY_HEADER=Authorization        # or x-api-key, etc.
# Map the API's JSON response (dotted paths) to the tool's fields:
export MSME_API_FIELD_MAP='{
  "udyam_number": "data.udyam_number",
  "pan": "data.pan",
  "enterprise_name": "data.enterprise_name",
  "enterprise_type": "data.enterprise_type",
  "major_activity": "data.major_activity",
  "organisation_type": "data.organisation_type",
  "date_of_incorporation": "data.date_of_incorporation",
  "date_of_udyam_registration": "data.registration_date",
  "state": "data.state",
  "district": "data.district"
}'
# Optional: how the API says "not found" when it doesn't return HTTP 404
export MSME_API_NOT_FOUND_PATH=data.status
export MSME_API_NOT_FOUND_VALUE=NOT_FOUND
```

If your vendor needs something these settings can't express (OAuth, signed requests), subclass
`BaseProvider` in `msme_verifier/providers.py`.

## 2. Struck-Off Companies

Upload your supplier list (sheet **Suppliers**: `S.No, Company Name, GST, PAN`, the same layout as
the template on the page). The tool will:

- take the PAN from the GSTIN when the PAN is missing, and flag a PAN that doesn't match the GSTIN
- skip PANs that can't be on MCA (individuals, HUFs, trusts...; only companies `C` and LLPs `F` are checked)
- look up each company's MCA master data and classify its **Company Status (for efiling)**:
  Active / **Struck Off** / **Under Process of Striking Off** / Other - Not Active (dormant,
  amalgamated, under liquidation, dissolved...) / Not Found on MCA
- detect **double status**, where one PAN is linked to several CINs, and pick the CIN whose name best
  matches your supplier name
- flag name mismatches and ACTIVE (INC-22A) non-compliance

**Download Report** produces the Excel format you use:

| Sheet | Contents |
|---|---|
| `Suppliers` | Your input (S.No, Company Name, GST, PAN) plus MCA Status, Result and Remarks |
| `Supplier-Final Sheet` | One row per supplier: Company Name, PAN, CIN, Company Name As Per CIN, Company Status (for efiling), ROC Code, Registration Number, Category, SubCategory, Class, Paid up Capital, Number of Members, Date of Incorporation, Registered Address, Other Address, Email, Listed, Suspended, ACTIVE Compliance, Date of last AGM, Date of Balance Sheet |
| `Supplier- Double Status` | Every CIN found for suppliers whose PAN maps to more than one company |
| `Summary` | Counts by result |

The status cells are colour-coded: red for struck off or under process, amber for other inactive
statuses, green for active.

### Data source for MCA data

MCA's "View Company / LLP Master Data" has no free API (it is behind a CAPTCHA), and MCA does not
publish a PAN-to-CIN mapping. Choose the source with `MCA_PROVIDER`:

| `MCA_PROVIDER` | Use |
|---|---|
| `demo` (default) | Made-up, repeatable data |
| `local` | A company master file you maintain (CSV/XLSX) with a `PAN` column and the report's column headers. See `sample_data/company_master.csv`. Several rows with the same PAN make a double status. Set `MCA_MASTER_FILE`. |
| `http` | A licensed company-data API (several KYC/data vendors offer "company details by PAN / CIN") |

```bash
export MCA_PROVIDER=http
export MCA_API_URL_PAN="https://api.your-vendor.com/company/by-pan?pan={id}"
export MCA_API_URL_CIN="https://api.your-vendor.com/company/master-data?cin={id}"
export MCA_API_KEY="Bearer <your-token>"         # MCA_API_KEY_HEADER defaults to Authorization
export MCA_API_RESULTS_PATH=data.companies       # where the record list sits in the JSON
export MCA_API_FIELD_MAP='{"cin": "cin", "company_name": "company_name", "status": "company_status", "roc_code": "roc", "date_of_last_agm": "last_agm_date"}'
```

The field keys you can map are: `cin, company_name, status, roc_code, registration_number, category,
subcategory, class_of_company, paid_up_capital, number_of_members, date_of_incorporation,
registered_address, other_address, email, listed, suspended, active_compliance, date_of_last_agm,
date_of_balance_sheet`.

Try it with the sample files:

```bash
MCA_PROVIDER=local MCA_MASTER_FILE=sample_data/company_master.csv python -m msme_verifier.app
# open http://127.0.0.1:5000/struck-off and upload sample_data/suppliers_to_check.csv
```

On Windows Command Prompt, set each variable on its own line first:
`set MCA_PROVIDER=local` and `set MCA_MASTER_FILE=sample_data\company_master.csv`.

## Other settings

| Variable | Default | Meaning |
|---|---|---|
| `MSME_CACHE_DAYS` | `30` | How long results stay cached (`0` turns the cache off) |
| `MSME_CACHE_DB` | `msme_cache.sqlite3` | Path of the cache database |
| `MSME_BULK_LIMIT` | `5000` | Maximum records in one bulk request |
| `HOST` / `PORT` | `127.0.0.1` / `5000` | Address the server listens on |

## API

| Endpoint | Body | Returns |
|---|---|---|
| `POST /api/verify` | `{"id": "AAAPL1234C"}` | One result |
| `POST /api/bulk` | multipart `file`, or `{"ids": [...]}` | `{results, summary}` |
| `POST /api/export` | `{"results": [...]}` | `.xlsx` report |
| `GET /api/template` | – | Bulk upload template |
| `POST /api/struck-off/check` | `{"id": "<PAN, GSTIN or CIN>"}` | One supplier result |
| `POST /api/struck-off/bulk` | multipart `file`, or `{"suppliers": [{"name", "gst", "pan", "cin"}]}` | `{results, summary}` |
| `POST /api/struck-off/report` | `{"results": [...]}` | `.xlsx` report |
| `GET /api/struck-off/template` | – | Supplier list template |
| `GET /api/struck-off/sample-report` | – | Sample report (demo data) |

## Development

```bash
pip install -r requirements-dev.txt
python -m pytest
```

## Compliance notes (MSME)

- The 45-day / Sec 43B(h) flag is "Yes" only for **Micro** and **Small** enterprises whose major
  activity is manufacturing or services. Medium enterprises and traders are marked "No".
- Classification is as of the date the result was fetched. Re-verify vendors from time to time,
  because an enterprise's category can change.
- For important decisions, confirm the result on the official Udyam portal.
