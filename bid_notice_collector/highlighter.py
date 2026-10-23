"""Explainable, recall-oriented title matching; never filters out records."""

import re
import unicodedata


def normalize(text):
    return re.sub(r"\s+", "", unicodedata.normalize("NFKC", str(text or "")).casefold())


def highlight_reasons(title, keywords=None, rules=None):
    """Any matching term is sufficient; return all matched groups and terms."""
    text = normalize(title)
    reasons = []
    groups = [("지정 키워드", keywords or [])] + list((rules or {}).items())
    for label, terms in groups:
        matches = []
        seen = set()
        for term in terms:
            normalized = normalize(term)
            if normalized and normalized in text and normalized not in seen:
                matches.append(term.strip())
                seen.add(normalized)
        if matches:
            reasons.append(f"{label}: {', '.join(matches)}")
    return reasons
