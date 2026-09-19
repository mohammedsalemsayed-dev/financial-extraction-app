"""Cell text -> value.  `parse_number` is the tested workhorse; `coerce_cell`
is the fast path used while cleaning a freshly-extracted table."""
from __future__ import annotations

import math
import re

NUM_RE = re.compile(r"^\(?-?[\d,]+(?:\.\d+)?\)?%?$")
_SUPERSCRIPT = re.compile(r"[⁰¹²³⁴-⁹]")
_CURRENCY_RE = re.compile(r"^(aed|usd|eur|gbp|egp|sar|rs\.?|\$|£|€)\s*", re.I)
_LEAD_REV_PAREN_RE = re.compile(r"^\)\s*\(\s*([\d,]+(?:\.\d+)?)\s*$")


class FormattedNumber(float):
    """A number whose source text carried a currency symbol, a trailing '%',
    and/or parentheses-as-negative -- behaves as a plain float everywhere it
    matters (isinstance checks, arithmetic, footing, health scoring, JSON
    encoding) since it IS one, but remembers the original prefix/suffix/
    paren-style so a caller that wants to show the user what was actually
    printed (not just the bare number every other numeric value collapses
    to) still can.

    `paren_negative`: the source used "(1,234)", not "-1,234" -- accounting
    notation for negative, not a different value. The tool still needs the
    real signed number for every arithmetic check (footing, health, Δ), so
    this only changes how a negative value is DISPLAYED back
    (`.formatted()`), never what it IS.

    Deliberately NOT surfaced through the normal `str`/`repr` -- every
    existing call site that treats a cell as a number (sums, comparisons,
    JSON) must keep seeing exactly that, unaffected. Read `.prefix` /
    `.suffix` / `.paren_negative` explicitly where the formatting is
    wanted."""
    def __new__(cls, value, prefix="", suffix="", paren_negative=False):
        obj = super().__new__(cls, value)
        obj.prefix = prefix
        obj.suffix = suffix
        obj.paren_negative = paren_negative
        return obj

    def formatted(self):
        n = float(self)
        neg = n < 0 and self.paren_negative
        n = -n if neg else n
        s = f"{int(n):,}" if n.is_integer() else f"{n:,.10f}".rstrip("0").rstrip(".")
        # a currency CODE reads naturally with a space ("AED 1,234"); a
        # currency SYMBOL doesn't ("$1,234", not "$ 1,234")
        sep = " " if self.prefix.isalpha() else ""
        body = f"{self.prefix}{sep}{s}{self.suffix}"
        return f"({body})" if neg else body


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
    paren_negative = False
    if s.startswith("(") and s.rstrip("*").rstrip().endswith(")"):
        neg = True
        paren_negative = True
        s = s[s.index("(") + 1: s.rindex(")")]
    elif s.startswith(")") and s.rstrip("*").rstrip().endswith("("):
        # Some reports (Etisalat / e& especially, but also seen in du's own
        # 2019 report) come out of PDF text extraction with the parenthesis
        # PAIR itself reversed -- ")1,234(" -- a bidi-reordering artifact of
        # the source PDF, not a different notation. The digits/commas inside
        # stay in correct reading order; only the two paren characters swap
        # visual position, so this is safe to treat exactly like the normal
        # "(1,234)" negative marker. telecom_extract.clean_cell already
        # special-cases this for the auto-detect path; this brings the same
        # handling to parse_number, which the manual/box-select path (and
        # everything else that calls coerce_cell) relies on instead.
        neg = True
        paren_negative = True
        s = s[s.index(")") + 1: s.rindex("(")]
    elif _LEAD_REV_PAREN_RE.match(s):
        # A step further than the swap above: img2table's own cell-text
        # assembly (not pdfplumber's -- the underlying PDF word order is
        # verified correct, see below) can carry BOTH reversed parens all
        # the way to the front instead of one on each side -- ")(417,358"
        # for a source "(417,358)" -- found live on e&'s FY2022 cash flow
        # statement (en-2022-1-eand-group-annual-report.pdf p50), where
        # pdfplumber's own extract_words() reads the same words as the
        # perfectly normal "(417,358)". Same bidi-reordering family as the
        # swap above, just with the closing paren's reordering carrying it
        # past the opening one; the digits/commas themselves stay in
        # correct reading order.
        neg = True
        paren_negative = True
        s = _LEAD_REV_PAREN_RE.match(s).group(1)
    m = re.search(r"\b(cr|dr)\b\.?$", s, re.I)               # credit / debit
    if m:
        neg = neg ^ (m.group(1).lower() == "cr")
        s = s[:m.start()].strip()
    cur_m = _CURRENCY_RE.match(s)
    currency = cur_m.group(1) if cur_m else ""
    s = _CURRENCY_RE.sub("", s)
    # US-GAAP-style statements often print the currency symbol only on a
    # value block's first and total/last row, not every row -- text
    # extraction can pick up the NEXT column's leading symbol and glue it
    # onto the end of THIS cell instead of its own ("$ 72,732 $"; confirmed
    # widespread on a real report -- 162 cells on one file alone). Strip a
    # trailing symbol the same way as the leading one.
    trail_m = re.search(r"\s*([$£€])\s*$", s)
    if trail_m:
        currency = currency or trail_m.group(1)
        s = s[:trail_m.start()]
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
    tokens = s.split()
    if len(tokens) > 1:
        # A real space-grouped number's thousands groups are always exactly
        # 3 digits after the leading group ("1 234 567"). When the tokens
        # don't fit that shape -- e.g. "2020 2020 2019 2019", four whole
        # 4-digit years -- this isn't one grouped figure, it's several
        # distinct numbers that ended up in the same cell (a wrapped
        # multi-line cell flattened to spaces by normspace, or several rows
        # collapsed into one upstream). Gluing those together would silently
        # fabricate a number nobody printed, so refuse instead, same as
        # _looks_garbled does for the newline-delimited version of this.
        grouped = re.fullmatch(r"\d{1,3}", tokens[0]) and \
            all(re.fullmatch(r"\d{3}", t) for t in tokens[1:-1]) and \
            re.fullmatch(r"\d{3}(?:\.\d+)?", tokens[-1])
        # A number that's ALREADY comma-grouped can pick up a stray extra
        # space (img2table's own cell-text assembly turning an internal
        # newline into a space via normspace, not the source PDF) ANYWHERE
        # inside it, not just right after one of its own commas -- confirmed
        # live on two different splits, both on en-2021-etisalat-group-
        # annual-report.pdf p61: "11, 180,517" (splits right after a comma)
        # and "1,1\n12,374" for a source "1,112,374" (splits mid-group,
        # after 4 characters with no comma adjacent at all -- pdfplumber's
        # own extract_words() reads the same spot as the single, ordinary
        # word "1,112,374"). Concatenating with no separator recovers the
        # original figure whenever that concatenation is EXACTLY one
        # strictly comma-grouped number -- which two genuinely distinct
        # numbers glued into one cell can't produce by accident: every
        # group after the first must start with a literal comma, so the
        # match can only continue past the first token's own digits if the
        # very next token happens to begin with a bare "," character, never
        # true for a real second number. Never fires on real French/EU
        # space-grouping either (that format never has commas inside its
        # groups, so the whole-concatenation pattern can't match it).
        if not grouped and re.fullmatch(r"\d{1,3}(?:,\d{3})+(?:\.\d+)?", "".join(tokens)):
            grouped = True
        if not grouped and sum(1 for t in tokens if NUM_RE.match(t)) >= 2:
            return None
    flat = re.sub(r"\s+", "", s)
    if re.fullmatch(r"[\d.]*,\d{1,2}", flat):                # European 1.234,56
        flat = flat.replace(".", "").replace(",", ".")
    elif "," in flat and "." in flat:                        # US 1,234.56
        flat = flat.replace(",", "")
    else:
        flat = flat.replace(",", "")
    # No real statement line item runs anywhere near this many digits (the
    # largest figure across every hand-verified fixture is 9 digits); a run
    # this long is almost always several numbers glued with no separator at
    # all to catch -- e.g. by _join_side_by_side reusing one row's text for
    # several rows it best-overlaps -- so there's no space left for the
    # check above to catch it on. Never guess where the real boundary was;
    # refuse instead. 16 keeps the existing +-10**15 round-trip contract
    # (test_round_trips_through_thousands_formatting) intact.
    if len(re.sub(r"\D", "", flat)) > 16:                     # absurd digit run
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
    if currency or had_percent or paren_negative:
        val = FormattedNumber(val, prefix=currency, suffix="%" if had_percent else "",
                              paren_negative=paren_negative)
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
        # NOT s.endswith(")") -- NUM_RE's own pattern allows a trailing
        # "%" AFTER the closing paren ("(11)%"), which is the majority
        # shape for a negative percentage in a real statement. Requiring
        # ")" as literally the last character silently dropped the
        # negative sign on every one of those: found live on a Micron
        # 10-K average-selling-price table, "(11)%" -> coerce_cell
        # returned a bare positive 11.0, an actively wrong figure (not
        # just an incomplete one) shown with total confidence.
        neg = s.startswith("(") and ")" in s
        had_percent = s.rstrip(")").endswith("%")
        num = s.strip("()%").replace(",", "")
        try:
            val = float(num)
            val = (-1 if neg else 1) * (int(val) if val.is_integer() else val)
            if had_percent or neg:
                return FormattedNumber(val, suffix="%" if had_percent else "",
                                       paren_negative=neg)
            return val
        except ValueError:
            return s
    n = parse_number(s)
    return n if n is not None else s
