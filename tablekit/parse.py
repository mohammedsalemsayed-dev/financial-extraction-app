"""Cell text -> value.  `parse_number` is the tested workhorse; `coerce_cell`
is the fast path used while cleaning a freshly-extracted table."""
from __future__ import annotations

import math
import re

NUM_RE = re.compile(r"^\(?-?[\d,]+(?:\.\d+)?\)?%?$")
_SUPERSCRIPT = re.compile(r"[⁰¹²³⁴-⁹]")
_CURRENCY_RE = re.compile(r"^(aed|usd|eur|gbp|egp|sar|rs\.?|\$|£|€)\s*", re.I)


class FormattedNumber(float):
    """A number whose source text carried a currency symbol and/or a
    trailing '%' -- behaves as a plain float everywhere it matters
    (isinstance checks, arithmetic, footing, health scoring, JSON encoding)
    since it IS one, but remembers the original prefix/suffix so a caller
    that wants to show the user what was actually printed (not just the
    bare number every other numeric value collapses to) still can.

    Deliberately NOT surfaced through the normal `str`/`repr` -- every
    existing call site that treats a cell as a number (sums, comparisons,
    JSON) must keep seeing exactly that, unaffected. Read `.prefix` /
    `.suffix` explicitly where the formatting is wanted."""
    def __new__(cls, value, prefix="", suffix=""):
        obj = super().__new__(cls, value)
        obj.prefix = prefix
        obj.suffix = suffix
        return obj

    def formatted(self):
        n = float(self)
        s = f"{int(n):,}" if n.is_integer() else f"{n:,.10f}".rstrip("0").rstrip(".")
        # a currency CODE reads naturally with a space ("AED 1,234"); a
        # currency SYMBOL doesn't ("$1,234", not "$ 1,234")
        sep = " " if self.prefix.isalpha() else ""
        return f"{self.prefix}{sep}{s}{self.suffix}"


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
    cur_m = _CURRENCY_RE.match(s)
    currency = cur_m.group(1) if cur_m else ""
    s = _CURRENCY_RE.sub("", s)
    # "AED 000" / "AED'000" / "USD 000" etc. is the standard "figures in
    # thousands" unit disclaimer printed once near a statement's header --
    # never a real data value -- so a bare "000" straight after a stripped
    # currency prefix is not a number, even though a real zero elsewhere is.
    if currency and re.fullmatch(r"'?0{2,3}", s.strip()):
        return None
    had_percent = s.rstrip().endswith("%")
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
    val = round(val, 4) if isinstance(val, float) else val
    if currency or had_percent:
        val = FormattedNumber(val, prefix=currency, suffix="%" if had_percent else "")
    return val


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
        had_percent = s.rstrip(")").endswith("%")
        num = s.strip("()%").replace(",", "")
        try:
            val = float(num)
            val = (-1 if neg else 1) * (int(val) if val.is_integer() else val)
            return FormattedNumber(val, suffix="%") if had_percent else val
        except ValueError:
            return s
    n = parse_number(s)
    return n if n is not None else s
