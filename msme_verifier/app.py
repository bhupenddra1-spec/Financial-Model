"""Flask web app: single search, bulk verification and Excel export."""
import io
import os
from datetime import datetime

from flask import Flask, jsonify, render_template, request, send_file

from .excel_io import build_report, build_template, extract_identifiers
from .providers import ProviderError, build_provider
from .service import ResultCache, VerificationService, summarise
from . import struck_off as so


def create_app(env=None, provider=None, company_provider=None):
    env = os.environ if env is None else env
    app = Flask(__name__)
    app.config["MAX_CONTENT_LENGTH"] = 10 * 1024 * 1024
    app.json.sort_keys = False  # keep summaries in their logical order
    bulk_limit = int(env.get("MSME_BULK_LIMIT", "5000"))
    cache_days = int(env.get("MSME_CACHE_DAYS", "30"))
    cache = ResultCache(env.get("MSME_CACHE_DB", "msme_cache.sqlite3"), cache_days) if cache_days > 0 else None
    service = VerificationService(provider or build_provider(env), cache)
    so_service = so.StruckOffService(company_provider or so.build_company_provider(env), cache)

    def xlsx_response(data, name):
        return send_file(io.BytesIO(data), as_attachment=True, download_name=name,
                         mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    @app.get("/")
    def index():
        return render_template("index.html", provider=service.provider.name, page="msme")

    @app.get("/struck-off")
    def struck_off_page():
        return render_template("struck_off.html", provider=so_service.provider.name, page="struck-off")

    @app.get("/api/health")
    def health():
        return jsonify(status="ok", provider=service.provider.name)

    @app.post("/api/verify")
    def verify():
        body = request.get_json(silent=True) or {}
        identifier = body.get("id") or request.form.get("id", "")
        if not str(identifier).strip():
            return jsonify(error="Enter a PAN or Udyam number"), 400
        return jsonify(service.verify(identifier))

    @app.post("/api/bulk")
    def bulk():
        if "file" in request.files:
            f = request.files["file"]
            try:
                identifiers = extract_identifiers(f.filename, f.read())
            except Exception as exc:  # unreadable/corrupt upload
                return jsonify(error=f"Could not read file: {exc}"), 400
        else:
            body = request.get_json(silent=True) or {}
            ids = body.get("ids", [])
            identifiers = ids.split() if isinstance(ids, str) else [str(i) for i in ids]
        identifiers = [i for i in identifiers if i.strip()]
        if not identifiers:
            return jsonify(error="No PAN / Udyam numbers found"), 400
        if len(identifiers) > bulk_limit:
            return jsonify(error=f"Bulk limit is {bulk_limit} records per request"), 400
        results = service.verify_many(identifiers)
        return jsonify(results=results, summary=summarise(results))

    @app.post("/api/export")
    def export():
        body = request.get_json(silent=True) or {}
        results = body.get("results") or []
        if not results:
            return jsonify(error="Nothing to export"), 400
        stamp = datetime.now().strftime("%Y%m%d_%H%M")
        return xlsx_response(build_report(results, summarise(results)), f"MSME_Verification_{stamp}.xlsx")

    @app.get("/api/template")
    def template():
        return xlsx_response(build_template(), "MSME_Bulk_Upload_Template.xlsx")

    # ---- Struck-off companies -------------------------------------------

    @app.post("/api/struck-off/check")
    def so_check():
        body = request.get_json(silent=True) or {}
        value = so.normalise(body.get("id", ""))
        if not value:
            return jsonify(error="Enter a PAN, GSTIN or CIN"), 400
        if so.GSTIN_RE.match(value):
            return jsonify(so_service.check(gst=value, sno=1))
        if so.CIN_RE.match(value) or so.LLPIN_RE.match(value):
            return jsonify(so_service.check(cin=value, sno=1))
        return jsonify(so_service.check(pan=value, sno=1))

    @app.post("/api/struck-off/bulk")
    def so_bulk():
        if "file" in request.files:
            f = request.files["file"]
            try:
                rows = so.parse_supplier_file(f.filename, f.read())
            except Exception as exc:  # unreadable/corrupt upload or missing columns
                return jsonify(error=f"Could not read file: {exc}"), 400
        else:
            rows = (request.get_json(silent=True) or {}).get("suppliers") or []
        if not rows:
            return jsonify(error="No suppliers with a PAN, GST or CIN were found"), 400
        if len(rows) > bulk_limit:
            return jsonify(error=f"Bulk limit is {bulk_limit} suppliers per request"), 400
        results = so_service.check_many(rows)
        return jsonify(results=results, summary=so.summarise(results))

    @app.post("/api/struck-off/report")
    def so_report():
        results = (request.get_json(silent=True) or {}).get("results") or []
        if not results:
            return jsonify(error="Nothing to export"), 400
        stamp = datetime.now().strftime("%Y%m%d_%H%M")
        return xlsx_response(so.build_struck_off_report(results), f"Struck_Off_Companies_Report_{stamp}.xlsx")

    @app.get("/api/struck-off/template")
    def so_template():
        return xlsx_response(so.build_supplier_template(), "Struck_Off_Supplier_List_Template.xlsx")

    @app.get("/api/struck-off/sample-report")
    def so_sample_report():
        demo = so.StruckOffService(so.DemoCompanyProvider())
        results = [demo.check(sno=i, name=n, gst=g, pan=p) for i, (n, g, p) in enumerate(so.SAMPLE_SUPPLIERS, 1)]
        return xlsx_response(so.build_struck_off_report(results), "Sample_Struck_Off_Companies_Report.xlsx")

    @app.errorhandler(413)
    def too_large(_):
        return jsonify(error="File too large (max 10 MB)"), 413

    return app


def main():
    try:
        app = create_app()
    except ProviderError as exc:
        raise SystemExit(f"Configuration error: {exc}")
    app.run(host=os.environ.get("HOST", "127.0.0.1"), port=int(os.environ.get("PORT", "5000")))


if __name__ == "__main__":
    main()
