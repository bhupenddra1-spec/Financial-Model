"""Flask web app: struck-off company search, bulk check and Excel report."""
import io
import os
from datetime import datetime

from flask import Flask, jsonify, render_template, request, send_file

from .excel_io import build_report, build_template, extract_suppliers
from .providers import ProviderError, build_provider
from .service import CheckService, ResultCache, summarise


def create_app(env=None, provider=None):
    env = os.environ if env is None else env
    app = Flask(__name__)
    app.config["MAX_CONTENT_LENGTH"] = 10 * 1024 * 1024
    bulk_limit = int(env.get("STRUCKOFF_BULK_LIMIT", "2000"))
    cache_days = int(env.get("STRUCKOFF_CACHE_DAYS", "7"))
    cache = ResultCache(env.get("STRUCKOFF_CACHE_DB", "struckoff_cache.sqlite3"), cache_days) if cache_days > 0 else None
    service = CheckService(provider or build_provider(env), cache)

    @app.get("/")
    def index():
        return render_template("index.html", provider=service.provider.name,
                               info=getattr(service.provider, "info", ""))

    @app.get("/api/health")
    def health():
        return jsonify(status="ok", provider=service.provider.name)

    @app.post("/api/check")
    def check():
        body = request.get_json(silent=True) or {}
        row = {k: str(body.get(k) or "") for k in ("name", "gstin", "pan", "cin")}
        if not any(v.strip() for v in row.values()):
            return jsonify(error="Enter a company name, PAN, GSTIN or CIN"), 400
        return jsonify(service.check(**row))

    @app.post("/api/bulk")
    def bulk():
        if "file" in request.files:
            f = request.files["file"]
            try:
                rows = extract_suppliers(f.filename, f.read())
            except Exception as exc:  # unreadable / corrupt upload
                return jsonify(error=f"Could not read file: {exc}"), 400
        else:
            body = request.get_json(silent=True) or {}
            rows = [{k: str(r.get(k) or "") for k in ("name", "gstin", "pan", "cin")}
                    for r in body.get("rows", []) if isinstance(r, dict)]
        if not rows:
            return jsonify(error="No suppliers found in the input"), 400
        if len(rows) > bulk_limit:
            return jsonify(error=f"Bulk limit is {bulk_limit} suppliers per request"), 400
        results = service.check_many(rows)
        return jsonify(results=results, summary=summarise(results))

    @app.post("/api/export")
    def export():
        results = (request.get_json(silent=True) or {}).get("results") or []
        if not results:
            return jsonify(error="Nothing to export"), 400
        stamp = datetime.now().strftime("%Y%m%d_%H%M")
        return send_file(io.BytesIO(build_report(results, summarise(results))), as_attachment=True,
                         download_name=f"Struck_Off_Companies_{stamp}.xlsx",
                         mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    @app.get("/api/template")
    def template():
        return send_file(io.BytesIO(build_template()), as_attachment=True,
                         download_name="Struck_Off_Upload_Template.xlsx",
                         mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    @app.errorhandler(413)
    def too_large(_):
        return jsonify(error="File too large (max 10 MB)"), 413

    return app


def main():
    try:
        app = create_app()
    except ProviderError as exc:
        raise SystemExit(f"Configuration error: {exc}")
    app.run(host=os.environ.get("HOST", "127.0.0.1"), port=int(os.environ.get("PORT", "5001")))


if __name__ == "__main__":
    main()
