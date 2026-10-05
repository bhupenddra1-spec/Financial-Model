"""Format checks and normalisation for CIN, PAN, GSTIN and company names."""
import re

CIN_RE = re.compile(r"^[LU][0-9]{5}[A-Z]{2}[0-9]{4}[A-Z]{3}[0-9]{6}$")
LLPIN_RE = re.compile(r"^[A-Z]{3}-[0-9]{4}$")
PAN_RE = re.compile(r"^[A-Z]{5}[0-9]{4}[A-Z]$")
GSTIN_RE = re.compile(r"^[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z][1-9A-Z]Z[0-9A-Z]$")

_NAME_WORDS = {"PVT": "PRIVATE", "LTD": "LIMITED", "CO": "COMPANY", "CORP": "CORPORATION", "&": "AND"}


def normalise(value):
    """Upper-case and strip all whitespace (for ID numbers)."""
    return re.sub(r"\s+", "", str(value or "").upper())


def normalise_name(value):
    """Canonical company name: upper-case, no punctuation, PVT->PRIVATE, LTD->LIMITED."""
    text = re.sub(r"[^A-Z0-9& ]", " ", str(value or "").upper().replace("&", " & "))
    return " ".join(_NAME_WORDS.get(w, w) for w in text.split())


def pan_from_gstin(gstin):
    """Characters 3-12 of a valid GSTIN are the holder's PAN."""
    g = normalise(gstin)
    return g[2:12] if GSTIN_RE.match(g) else ""


def classify_input(name="", gstin="", pan="", cin=""):
    """Clean a supplier row and report format problems. Returns (clean_dict, [errors])."""
    clean = {"name": " ".join(str(name or "").split()), "gstin": normalise(gstin),
             "pan": normalise(pan), "cin": normalise(cin)}
    errors = []
    if clean["cin"] and not (CIN_RE.match(clean["cin"]) or LLPIN_RE.match(clean["cin"])):
        errors.append("Invalid CIN/LLPIN format")
    if clean["gstin"] and not GSTIN_RE.match(clean["gstin"]):
        errors.append("Invalid GSTIN format")
    if clean["pan"] and not PAN_RE.match(clean["pan"]):
        errors.append("Invalid PAN format")
    if not clean["pan"] and clean["gstin"] and not errors:
        clean["pan"] = pan_from_gstin(clean["gstin"])
    if clean["pan"] and clean["gstin"] and GSTIN_RE.match(clean["gstin"]) \
            and pan_from_gstin(clean["gstin"]) != clean["pan"]:
        errors.append("PAN does not match the PAN embedded in the GSTIN")
    if not (clean["name"] or clean["pan"] or clean["cin"]) and not errors:
        errors.append("Provide at least a company name, PAN, GSTIN or CIN")
    return clean, errors
