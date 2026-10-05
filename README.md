# MSME Search & Verification Tool

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

## Development

```bash
pip install -r requirements-dev.txt
python -m pytest
```

## Compliance notes

- The 45-day / Sec 43B(h) flag is "Yes" only for **Micro** and **Small** enterprises whose major
  activity is manufacturing or services. Medium enterprises and traders are marked "No".
- Classification is as of the date the result was fetched. Re-verify vendors from time to time,
  because an enterprise's category can change.
- For important decisions, confirm the result on the official Udyam portal.
