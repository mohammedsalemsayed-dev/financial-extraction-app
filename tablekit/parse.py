"""Cell text -> value.  `parse_number` is the tested workhorse; `coerce_cell`
is the fast path used while cleaning a freshly-extracted table."""
from __future__ import annotations

import math
import re

NUM_RE = re.compile(r"^\(?-?[\d,]+(?:\.\d+)?\)?%?$")
_SUPERSCRIPT = re.compile(r"[⁰¹²³⁴-⁹]")


def normspace(s):
    """Collapse whitespace and turn the U+FFFD 'replacement char' that some
    older PDFs use as an inter-word space back into a space."""
    s = str(s).replace("�", " ").replace("\xa0", " ").replace("\n", " ")
    return re.sub(r"\s+", " ", s).strip()


def parse_number(raw):
    """Parse a figure the way a report might print it.  Returns int / float, or
    None if `raw` is not a number.  Handles:
      1,234   1,234.56   (1,234)   -1,234   1,234-   1 234   1.234,56 (EU)
      12.5%   $1,234 / AED 1,234   1,234 CR / DR   1,234*  1,234(1)  1,234 (a)
    A bare '-' / en-dash / 'nil' / 'n/a' is a blank, not a number -> None.
    """
    if raw is None:
        return None
    s = str(raw).replace("\xa0", " ").replace("−", "-").strip()
    s = _SUPERSCRIPT.sub("", s)
    if s.lower() in ("", "-", "–", "—", "--", "nil", "n/a", "na"):
        return None
    neg = False
    if s.startswith("(") and s.rstrip("*").rstrip().endswith(")"):
        neg = True
        s = s[s.index("(") + 1: s.rindex(")")]
    m = re.search(r"\b(cr|dr)\b\.?$", s, re.I)               # credit / debit
    if m:
        neg = neg ^ (m.group(1).lower() == "cr")
        s = s[:m.start()].strip()
    had_currency_prefix = bool(re.match(r"^(aed|usd|eur|gbp|egp|sar|rs\.?|\$|£|€)\s*", s, re.I))
    s = re.sub(r"^(aed|usd|eur|gbp|egp|sar|rs\.?|\$|£|€)\s*", "", s, flags=re.I)
    # "AED 000" / "AED'000" / "USD 000" etc. is the standard "figures in
    # thousands" unit disclaimer printed once near a statement's header --
    # never a real data value -- so a bare "000" straight after a stripped
    # currency prefix is not a number, even though a real zero elsewhere is.
    if had_currency_prefix and re.fullmatch(r"'?0{2,3}", s.strip()):
        return None
    s = re.sub(r"[%\s]*$", "", s)
    s = re.sub(r"\s*[\*†‡]$", "", s)               # * dagger etc.
    s = re.sub(r"\s*\([a-z0-9]{1,3}\)$", "", s, flags=re.I)  # trailing (a) / (1)
    s = s.strip()
    # trailing minus -- only '1234-' (digit right before '-'), never '<ref> -'
    if re.search(r"\d-$", s):
        neg = not neg
        s = s[:-1].strip()
    s = s.lstrip("+-")
    sign = -1 if str(raw).strip().startswith("-") else 1
    if not re.search(r"\d", s):
        return None
    flat = s.replace(" ", "")
    if re.fullmatch(r"[\d.]*,\d{1,2}", flat):                # European 1.234,56
        flat = flat.replace(".", "").replace(",", ".")
    elif "," in flat and "." in flat:                        # US 1,234.56
        flat = flat.replace(",", "")
    else:
        flat = flat.replace(",", "")
    if len(re.sub(r"\D", "", flat)) > 18:                    # absurd digit run
        return None
    try:
        val = float(flat)
    except (ValueError, OverflowError):
        return None
    if not math.isfinite(val):
        return None
    val = -val if neg else val * sign
    try:
        if val == int(val):
            val = int(val)
    except (OverflowError, ValueError):
        return None
    return round(val, 4) if isinstance(val, float) else val


def coerce_cell(v):
    """Fast cell coercion used by the cleaner: obvious numbers straight through,
    a bare dash kept as a spacer, everything else via `parse_number`."""
    if v is None:
        return None
    s = normspace(v)
    if s == "" or s == "-" or s == "–":
        return None if s == "" else s
    if NUM_RE.match(s):
        neg = s.startswith("(") and s.endswith(")")
        num = s.strip("()%").replace(",", "")
        try:
            val = float(num)
            return (-1 if neg else 1) * (int(val) if val.is_integer() else val)
        except ValueError:
            return s
    n = parse_number(s)
    return n if n is not None else s
