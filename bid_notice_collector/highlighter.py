# 파일 역할: 공고명의 관심 분야·과업 표현을 비교하여 노란색 강조 여부와 근거를 생성합니다.
# 마지막 업데이트: 2026-10-23
# 작업자: SMI 장지수

"""Explainable, recall-oriented title matching; never filters out records."""

import re
import unicodedata


def normalize(text):
    # 전각 문자·대소문자·띄어쓰기 차이를 제거해 AI/ＡＩ, R&D/R& D 등을 같은 표현으로 봅니다.
    return re.sub(r"\s+", "", unicodedata.normalize("NFKC", str(text or "")).casefold())


def highlight_reasons(title, keywords=None, rules=None):
    """Any matching term is sufficient; return all matched groups and terms."""
    text = normalize(title)
    reasons = []
    # 분야와 과업 조건은 OR입니다. 하나만 맞아도 검토 후보이며, 실제 참여 적합성을 확정하지 않습니다.
    groups = [("지정 키워드", keywords or [])] + list((rules or {}).items())
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
            reasons.append(f"{label}: {', '.join(matches)}")
    return reasons
