# 파일 역할: 공고명의 관심 분야·과업 표현을 비교하여 노란색 강조 여부와 근거를 생성합니다.
# 마지막 업데이트: 2026-10-23
# 작업자: SMI 장지수

"""Explainable, recall-oriented title matching; never filters out records."""

import re
import unicodedata

DEFENSE_ISSUER_TERMS = (
    "국방", "방위", "방산", "방위사업청", "군", "군사", "무기", "육군", "해군", "공군",
)


def normalize(text):
    # 전각 문자·대소문자·띄어쓰기 차이를 제거해 AI/ＡＩ, R&D/R& D 등을 같은 표현으로 봅니다.
    return re.sub(r"\s+", "", unicodedata.normalize("NFKC", str(text or "")).casefold())


def highlight_reasons(title, keywords=None, rules=None, issuer=None):
    """Apply a looser field/keyword rule for defense-related issuers."""
    text = normalize(title)
    issuer_text = normalize(issuer)
    groups = [("지정 키워드", keywords or [])] + list((rules or {}).items())
    matched_groups = []
    for label, terms in groups:
        matches = []
        # 동일 분류 안에서 표기만 다른 키워드가 근거에 반복되지 않도록 정규화한 값을 기록합니다.
        seen = set()
        for term in terms:
            normalized = normalize(term)
            if normalized and normalized in text and normalized not in seen:
                matches.append(term.strip())
                seen.add(normalized)
        if matches:
            matched_groups.append((label, matches))

    explicit_match = any(label == "지정 키워드" for label, _ in matched_groups)
    has_field = any(label.startswith("관심 분야:") for label, _ in matched_groups)
    has_task = any(
        label.startswith("과업") or label.startswith("검토 대상:")
        for label, _ in matched_groups
    )
    is_defense_issuer = any(normalize(term) in issuer_text for term in DEFENSE_ISSUER_TERMS)
    if is_defense_issuer:
        if not explicit_match and not (has_field or has_task):
            return []
    elif not explicit_match and not (has_field and has_task):
        return []
    return [f"{label}: {', '.join(matches)}" for label, matches in matched_groups]
