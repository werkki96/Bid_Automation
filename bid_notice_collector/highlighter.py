# 파일 역할: 공고명의 관심 분야·과업 표현을 비교하여 노란색 강조 여부와 근거를 생성합니다.
# 마지막 업데이트: 2026-11-09
# 작업자: SMI 장지수

"""Explainable, recall-oriented title matching; never filters out records."""

import re
import unicodedata

# '군' 한 글자는 강화군·영광군 같은 지자체와 군포시 기관까지 국방기관으로 오인하므로 쓰지 않습니다.
# '사단'도 '사단법인'과 겹치므로 제외합니다.
DEFENSE_ISSUER_TERMS = (
    "국방", "방위", "방산", "방위사업청", "군사", "무기", "육군", "해군", "공군",
    "국군", "해병대", "합동참모", "군단", "부대", "사령부", "군수", "비행단",
)


def normalize(text):
    # 전각 문자·대소문자·띄어쓰기 차이를 제거해 AI/ＡＩ, R&D/R& D 등을 같은 표현으로 봅니다.
    return re.sub(r"\s+", "", unicodedata.normalize("NFKC", str(text or "")).casefold())


ASCII_ALNUM = re.compile(r"[a-z0-9]")


def term_pattern(normalized_term):
    """Space-insensitive pattern; English/number edges must not touch other English/numbers."""
    # 예: '6G'는 '5G 6G'·'6G 특화망'에는 맞지만 공고번호 '26G193'이나 '6GHz' 안에서는 맞지 않습니다.
    body = r"\s*".join(re.escape(char) for char in normalized_term)
    prefix = r"(?<![a-z0-9])" if ASCII_ALNUM.match(normalized_term[0]) else ""
    suffix = r"(?![a-z0-9])" if ASCII_ALNUM.match(normalized_term[-1]) else ""
    return re.compile(prefix + body + suffix)


def highlight_reasons(title, keywords=None, rules=None, issuer=None):
    """Apply a looser field/keyword rule for defense-related issuers."""
    # 띄어쓰기를 남긴 텍스트에서 비교해야 영문·숫자 키워드의 앞뒤 경계를 판단할 수 있습니다.
    text = unicodedata.normalize("NFKC", str(title or "")).casefold()
    issuer_text = normalize(issuer)
    groups = [("지정 키워드", keywords or [])] + list((rules or {}).items())
    matched_groups = []
    for label, terms in groups:
        matches = []
        # 동일 분류 안에서 표기만 다른 키워드가 근거에 반복되지 않도록 정규화한 값을 기록합니다.
        seen = set()
        for term in terms:
            normalized = normalize(term)
            if normalized and normalized not in seen and term_pattern(normalized).search(text):
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
