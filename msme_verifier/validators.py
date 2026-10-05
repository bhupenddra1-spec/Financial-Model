"""Format validation and normalisation for PAN and Udyam Registration Numbers."""
import re

PAN_RE = re.compile(r"^[A-Z]{5}[0-9]{4}[A-Z]$")
# UDYAM-<state code>-<district code>-<7 digit serial>, e.g. UDYAM-MH-26-0012345
UDYAM_RE = re.compile(r"^UDYAM-[A-Z]{2}-[0-9]{2}-[0-9]{7}$")

# The 4th character of a PAN identifies the holder's status.
PAN_HOLDER_TYPES = {
    "P": "Individual / Proprietorship",
    "C": "Company",
    "H": "Hindu Undivided Family (HUF)",
    "F": "Firm / LLP",
    "A": "Association of Persons (AOP)",
    "T": "Trust",
    "B": "Body of Individuals (BOI)",
    "L": "Local Authority",
    "J": "Artificial Juridical Person",
    "G": "Government",
}


def normalise(value):
    """Upper-case and strip whitespace; tolerate spaces/underscores in Udyam numbers."""
    if value is None:
        return ""
    text = str(value).strip().upper()
    text = re.sub(r"\s+", "", text)
    return text.replace("_", "-")


def is_valid_pan(value):
    return bool(PAN_RE.match(normalise(value)))


def is_valid_udyam(value):
    return bool(UDYAM_RE.match(normalise(value)))


def detect_id_type(value):
    """Return 'PAN', 'UDYAM' or None for an identifier."""
    text = normalise(value)
    if PAN_RE.match(text):
        return "PAN"
    if UDYAM_RE.match(text):
        return "UDYAM"
    return None


def pan_holder_type(pan):
    pan = normalise(pan)
    if not PAN_RE.match(pan):
        return None
    return PAN_HOLDER_TYPES.get(pan[3], "Unknown")
