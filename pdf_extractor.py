"""
PDF Invoice Extractor

Provides `PDFInvoiceExtractor` which extracts key fields from a single PDF invoice.

Dependencies: `pdfplumber` recommended. If not available, falls back to `PyMuPDF` (fitz).

Usage:
    extractor = PDFInvoiceExtractor("/path/to/invoice.pdf")
    data = extractor.extract_data()

Returned dict keys: `date` (datetime.date or None), `amount` (float or None), `vendor` (str or None)
"""
from __future__ import annotations

import logging
import re
from datetime import datetime, date
from typing import Optional, Dict, Tuple

try:
    import pdfplumber
    _HAS_PDFPLUMBER = True
except Exception:
    pdfplumber = None
    _HAS_PDFPLUMBER = False

try:
    import fitz  # PyMuPDF
    _HAS_PYMUPDF = True
except Exception:
    fitz = None
    _HAS_PYMUPDF = False

try:
    from dateutil import parser as _dateutil_parser
    _HAS_DATEUTIL = True
except Exception:
    _HAS_DATEUTIL = False

logger = logging.getLogger(__name__)
logging.basicConfig(level=logging.INFO)


class PDFInvoiceExtractor:
    """Extract Invoice Date, Total Amount, and Vendor Name from a PDF file.

    Initialization: provide a path to a single PDF file.

    Public API:
        extract_data() -> dict with keys: `date`, `amount`, `vendor`.
    """

    # Regex patterns
    _DATE_RE = re.compile(
        r"(?P<date>\b(?:\d{1,2}[/-]\d{1,2}[/-]\d{2,4})\b|\b(?:"
        r"Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|"
        r"Jul(?:y)?|Aug(?:ust)?|Sep(?:tember)?|Oct(?:ober)?|Nov(?:ember)?|"
        r"Dec(?:ember)?)\.?\s+\d{1,2},?\s+\d{4})",
        flags=re.IGNORECASE,
    )

    # Amount pattern: currency symbol optional, digits with optional grouping commas, optional decimals
    _AMOUNT_RE = re.compile(r"(?P<amount>[€$£¥]?\s?[-+]?\d{1,3}(?:[,\d]{0,15})?(?:\.\d{1,2})?)")

    # Keywords to search for amounts (case-insensitive)
    _AMOUNT_KEYWORDS = re.compile(
        r"(?:total|amount due|grand total|balance due|invoice total|amount:)", re.IGNORECASE
    )

    # Vendor anchors
    _VENDOR_ANCHORS = re.compile(r"(?:vendor|supplier|from|bill from|billed by)", re.IGNORECASE)

    def __init__(self, pdf_path: str):
        self.pdf_path = pdf_path

    def extract_data(self) -> Dict[str, Optional[object]]:
        """Extracts and returns a dictionary: {'date': date|None, 'amount': float|None, 'vendor': str|None}.

        If any error occurs while opening/reading the PDF, logs a warning and returns all None values.
        """
        try:
            full_text, first_page_text = self._extract_text()
        except Exception as exc:
            logger.warning("Failed to read PDF '%s': %s", self.pdf_path, exc)
            return {"date": None, "amount": None, "vendor": None}

        # Normalize whitespace
        text = full_text or ""
        text = re.sub(r"\u00A0|\r\n|\r", " ", text)

        found_date = self._find_date(text)
        found_amount = self._find_amount(text)
        found_vendor = self._find_vendor(first_page_text or text)

        return {"date": found_date, "amount": found_amount, "vendor": found_vendor}

    def _extract_text(self) -> Tuple[str, str]:
        """Extracts text from the PDF using pdfplumber or PyMuPDF as fallback.

        Returns a tuple (full_text, first_page_text).
        """
        if _HAS_PDFPLUMBER:
            try:
                with pdfplumber.open(self.pdf_path) as pdf:
                    pages = list(pdf.pages)
                    texts = [p.extract_text() or "" for p in pages]
                    full_text = "\n".join(texts)
                    first_page_text = texts[0] if texts else ""
                    return full_text, first_page_text
            except Exception:
                # Fall through to try PyMuPDF if available
                logger.debug("pdfplumber failed, trying PyMuPDF if available")

        if _HAS_PYMUPDF:
            try:
                doc = fitz.open(self.pdf_path)
                texts = []
                for page in doc:
                    texts.append(page.get_text("text") or "")
                full_text = "\n".join(texts)
                first_page_text = texts[0] if texts else ""
                doc.close()
                return full_text, first_page_text
            except Exception as exc:
                raise RuntimeError(f"Failed to read PDF with PyMuPDF: {exc}")

        raise RuntimeError("No PDF backend available: install pdfplumber or PyMuPDF (fitz)")

    def _find_date(self, text: str) -> Optional[date]:
        """Search for dates using regex and parse to a `datetime.date` when possible.

        Pattern covers numeric dates like 01/31/2023 or 31-01-2023 and textual dates like "January 31, 2023".
        """
        for match in self._DATE_RE.finditer(text):
            date_str = match.group("date")
            parsed = self._try_parse_date(date_str)
            if parsed:
                return parsed
        return None

    def _try_parse_date(self, s: str) -> Optional[date]:
        s = s.strip().replace("\n", " ")
        # Try dateutil if available for robust parsing
        if _HAS_DATEUTIL:
            try:
                dt = _dateutil_parser.parse(s, dayfirst=False, fuzzy=True)
                return dt.date()
            except Exception:
                pass

        # Fallback: try common formats explicitly
        formats = [
            "%d/%m/%Y",
            "%m/%d/%Y",
            "%d-%m-%Y",
            "%m-%d-%Y",
            "%Y-%m-%d",
            "%B %d, %Y",
            "%b %d, %Y",
            "%B %d %Y",
            "%b %d %Y",
        ]
        for fmt in formats:
            try:
                dt = datetime.strptime(s, fmt)
                return dt.date()
            except Exception:
                continue
        return None

    def _find_amount(self, text: str) -> Optional[float]:
        """Find monetary amounts using the most likely invoice total anchor.

        Strategy:
        - Prefer the last occurrence of a total-related keyword so subtotals and address numbers are avoided.
        - Look for numeric values after the keyword, or in the same localized window.
        - Only use broader fallbacks if invoice keywords are absent.
        """
        # Prefer the last total-related keyword occurrence, since invoice totals are usually at the bottom.
        for k in reversed(list(self._AMOUNT_KEYWORDS.finditer(text))):
            window_start = max(0, k.start() - 20)
            window_end = min(len(text), k.end() + 200)
            window = text[window_start:window_end]
            keyword_end = k.end() - window_start
            amount_candidate = self._find_amount_in_text(window, keyword_end=keyword_end)
            if amount_candidate is not None:
                return amount_candidate

        # Fallback: search for any currency-style amount in the bottom section of the document.
        lines = [ln for ln in text.splitlines() if ln.strip()]
        for line in reversed(lines[-20:]):
            amount_candidate = self._find_amount_in_text(line)
            if amount_candidate is not None:
                return amount_candidate

        # Last fallback: search for the cleanest amount across the whole document.
        all_amounts = [
            self._clean_amount(m.group("amount"))
            for m in self._AMOUNT_RE.finditer(text)
        ]
        all_amounts = [a for a in all_amounts if a is not None]
        if all_amounts:
            return max(all_amounts)
        return None

    def _find_amount_in_text(self, text: str, keyword_end: Optional[int] = None) -> Optional[float]:
        """Search a text snippet for amount values.

        If keyword_end is provided, first search only after that position.
        """
        if keyword_end is not None and 0 <= keyword_end < len(text):
            after_segment = text[keyword_end:]
            for m in self._AMOUNT_RE.finditer(after_segment):
                cleaned = self._clean_amount(m.group("amount"))
                if cleaned is not None:
                    return cleaned

        for m in self._AMOUNT_RE.finditer(text):
            cleaned = self._clean_amount(m.group("amount"))
            if cleaned is not None:
                return cleaned
        return None

    def _clean_amount(self, amount_str: str) -> Optional[float]:
        if not amount_str:
            return None
        # Remove common currency symbols and whitespace
        cleaned = amount_str.strip()
        cleaned = re.sub(r"[€$£¥\s]", "", cleaned)
        cleaned = cleaned.replace(",", "")
        cleaned = cleaned.replace("(", "-").replace(")", "")
        # Ignore obvious non-monetary numbers like long zip/postal codes without decimals.
        if re.fullmatch(r"\d{5,}", cleaned):
            return None
        try:
            return float(cleaned)
        except Exception:
            return None

    def _find_vendor(self, first_page_text: str) -> Optional[str]:
        """Attempt to detect vendor/company name.

        Strategy:
        - Look for explicit anchors like 'Vendor:' or 'From:' and capture the adjacent text.
        - Otherwise, inspect the top lines of the first page and pick the first prominent non-invoice header line.
        """
        if not first_page_text:
            return None

        # Attempt to find anchored vendor lines like 'Vendor: ACME Corp'
        anchored = re.search(r"(?:Vendor|Supplier|From|Billed By)[:\s]+(.{2,120})", first_page_text, re.IGNORECASE)
        if anchored:
            vendor = anchored.group(1).strip()
            # Trim trailing junk after comma or newline
            vendor = vendor.split("\n")[0].strip()
            return vendor or None

        # Heuristic: take the first several non-empty lines and choose the first that isn't an invoice header
        lines = [ln.strip() for ln in first_page_text.splitlines() if ln.strip()]
        if not lines:
            return None

        # Common header words to ignore
        ignore_re = re.compile(r"^(invoice|tax invoice|statement|bill|date|invoice no|page)\b", re.IGNORECASE)

        for ln in lines[:12]:
            if ignore_re.search(ln):
                continue
            # Avoid lines that look like addresses or key-value pairs
            if len(ln) > 2 and "," not in ln and ":" not in ln and len(ln) < 80:
                # Probably a company name
                return ln

        # Last resort: return the first non-empty line
        return lines[0]


if __name__ == "__main__":
    # Demo usage - replace the path below with a real invoice PDF path when running.
    sample_path = "./sample-invoice.pdf"
    extractor = PDFInvoiceExtractor(sample_path)
    data = extractor.extract_data()
    print("Extracted:")
    print("  date:  ", data.get("date"))
    print("  amount:", data.get("amount"))
    print("  vendor:", data.get("vendor"))
