"""Check workflow: validate -> cache -> provider lookup -> struck-off classification."""
from datetime import datetime

from msme_verifier.service import ResultCache  # generic SQLite result cache

from .providers import ProviderError
from .validators import classify_input

FLAG_STRUCK_OFF = "Struck-Off"
FLAG_UNDER_PROCESS = "Under Strike-Off Process"
FLAG_ACTIVE = "Active"
FLAG_OTHER = "Other Status"
FLAG_DOUBLE = "Double Status"
FLAG_NOT_FOUND = "Not Found"
FLAG_INVALID = "Invalid Input"
FLAG_ERROR = "Error"

__all__ = ["ResultCache", "CheckService", "classify_status", "summarise"]


def classify_status(status):
    """Map an MCA company status string to one of the FLAG_* groups."""
    s = (status or "").lower()
    if "strik" in s or "struck" in s:
        return FLAG_UNDER_PROCESS if ("process" in s or "under" in s) else FLAG_STRUCK_OFF
    if s.startswith("active"):
        return FLAG_ACTIVE
    return FLAG_OTHER if s else FLAG_NOT_FOUND


class CheckService:
    def __init__(self, provider, cache=None):
        self.provider = provider
        self.cache = cache

    def _search(self, kind, value):
        key = f"{self.provider.name}:{kind}:{value}"
        cached = self.cache.get(key) if self.cache else None
        if cached is not None:
            return cached
        records = self.provider.search(kind, value)
        if records is not None and self.cache:
            self.cache.put(key, records)
        return records

    def check(self, name="", gstin="", pan="", cin=""):
        clean, errors = classify_input(name, gstin, pan, cin)
        result = {"input": clean, "matched_by": "", "records": [], "flag": "", "status": "",
                  "remarks": "", "source": self.provider.name,
                  "checked_at": datetime.now().isoformat(timespec="seconds")}
        if errors:
            result.update(flag=FLAG_INVALID, remarks="; ".join(errors))
            return result

        # Most specific identifier first; fall back if a provider can't search by it.
        attempts = [(k, v) for k, v in (("CIN", clean["cin"]), ("PAN", clean["pan"]), ("NAME", clean["name"])) if v]
        records, notes = [], []
        try:
            for kind, value in attempts:
                found = self._search(kind, value)
                if found is None:
                    notes.append(f"Provider does not support {kind} search")
                    continue
                if found:
                    records, result["matched_by"] = found, kind
                    break
        except ProviderError as exc:
            result.update(flag=FLAG_ERROR, remarks=str(exc))
            return result

        if not records:
            result.update(flag=FLAG_NOT_FOUND, remarks="; ".join(notes) or "No company found in MCA master data")
            return result

        for r in records:
            r["group"] = classify_status(r["status"])
            if not r["pan"]:
                r["pan"] = clean["pan"]
        groups = {r["group"] for r in records}
        result["records"] = records
        result["status"] = "; ".join(dict.fromkeys(r["status"] for r in records if r["status"]))
        if len(groups) > 1:
            result["flag"] = FLAG_DOUBLE
            result["remarks"] = f"{len(records)} matches with different statuses - verify manually"
        else:
            result["flag"] = records[0]["group"]
            if len(records) > 1:
                result["remarks"] = f"{len(records)} matches (same status) - confirm the right entity"
        if result["matched_by"] == "NAME":
            result["remarks"] = ("Matched by name only - confirm with CIN/PAN. " + result["remarks"]).strip()
        return result

    def check_many(self, rows):
        return [self.check(**{k: r.get(k, "") for k in ("name", "gstin", "pan", "cin")}) for r in rows]


def summarise(results):
    summary = {"total": len(results)}
    for r in results:
        summary[r["flag"]] = summary.get(r["flag"], 0) + 1
    return summary
