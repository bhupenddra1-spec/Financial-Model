"""Verification workflow: validate -> cache -> provider lookup -> compliance enrichment."""
import json
import sqlite3
import threading
import time
from datetime import datetime

from .providers import FIELDS, ProviderError
from .validators import detect_id_type, normalise, pan_holder_type

STATUS_REGISTERED = "Registered"
STATUS_NOT_REGISTERED = "Not Registered"
STATUS_INVALID = "Invalid Format"
STATUS_ERROR = "Error"


def payment_rule(enterprise_type, major_activity):
    """Whether the 45-day payment limit (MSMED Act s.15) and IT Act s.43B(h) apply.

    They protect micro and small suppliers; medium enterprises and traders
    (whose Udyam registration is only for priority-sector lending) are excluded.
    """
    etype = (enterprise_type or "").strip().lower()
    activity = (major_activity or "").strip().lower()
    if etype not in ("micro", "small"):
        return "No"
    if activity.startswith("trad"):
        return "No (Trader)"
    return "Yes"


class ResultCache:
    """SQLite cache so repeat checks don't hit a paid API again."""

    def __init__(self, path, ttl_days):
        self.ttl = ttl_days * 86400
        self._lock = threading.Lock()
        self._db = sqlite3.connect(path, check_same_thread=False)
        self._db.execute("CREATE TABLE IF NOT EXISTS results (key TEXT PRIMARY KEY, data TEXT, ts REAL)")

    def get(self, key):
        if self.ttl <= 0:
            return None
        with self._lock:
            row = self._db.execute("SELECT data, ts FROM results WHERE key = ?", (key,)).fetchone()
        if row and time.time() - row[1] < self.ttl:
            return json.loads(row[0])
        return None

    def put(self, key, data):
        with self._lock:
            self._db.execute("REPLACE INTO results VALUES (?, ?, ?)", (key, json.dumps(data), time.time()))
            self._db.commit()


class VerificationService:
    def __init__(self, provider, cache=None):
        self.provider = provider
        self.cache = cache

    def verify(self, raw):
        query = normalise(raw)
        id_type = detect_id_type(query)
        result = {"query": query, "id_type": id_type or "", **{f: "" for f in FIELDS},
                  "status": "", "msme_payment_rule": "", "remarks": "",
                  "source": self.provider.name, "checked_at": datetime.now().isoformat(timespec="seconds")}
        if not id_type:
            result.update(status=STATUS_INVALID,
                          remarks="Not a valid PAN (AAAAA9999A) or Udyam number (UDYAM-XX-00-0000000)")
            return result

        cache_key = f"{self.provider.name}:{query}"
        cached = self.cache.get(cache_key) if self.cache else None
        if cached:
            cached["remarks"] = "Served from cache"
            return cached

        try:
            record = self.provider.lookup(query, id_type)
        except ProviderError as exc:
            result.update(status=STATUS_ERROR, remarks=str(exc))
            return result  # errors are not cached so they can be retried

        if record is None:
            result["status"] = STATUS_NOT_REGISTERED
            result["remarks"] = f"No Udyam registration found for this {id_type}"
            if id_type == "PAN":
                result["pan"] = query
                result["organisation_type"] = pan_holder_type(query) or ""
            else:
                result["udyam_number"] = query
            result["msme_payment_rule"] = "No"
        else:
            result.update({f: record.get(f, "") or "" for f in FIELDS})
            result["status"] = STATUS_REGISTERED
            if not result["organisation_type"] and result["pan"]:
                result["organisation_type"] = pan_holder_type(result["pan"]) or ""
            result["msme_payment_rule"] = payment_rule(result["enterprise_type"], result["major_activity"])

        if self.cache:
            self.cache.put(cache_key, result)
        return result

    def verify_many(self, identifiers):
        """Verify a list, de-duplicating while preserving input order."""
        seen, results = {}, []
        for raw in identifiers:
            key = normalise(raw)
            if not key:
                continue
            if key not in seen:
                seen[key] = self.verify(key)
            results.append(seen[key])
        return results


def summarise(results):
    summary = {"total": len(results)}
    for r in results:
        summary[r["status"]] = summary.get(r["status"], 0) + 1
        if r["enterprise_type"]:
            summary[r["enterprise_type"]] = summary.get(r["enterprise_type"], 0) + 1
    return summary
