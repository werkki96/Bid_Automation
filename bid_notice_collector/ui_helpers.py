# 파일 역할: 웹·데스크톱에서 공유하는 표시 형식 변환을 제공합니다.
# 마지막 업데이트: 2026-10-27
# 작업자: SMI 장지수

from decimal import Decimal, InvalidOperation


def format_price(value):
    """원본 데이터를 변경하지 않고 금액에 천 단위 쉼표를 붙입니다."""
    if value is None or value == '':
        return ''
    try:
        number = Decimal(str(value).replace(',', '').strip())
        if not number.is_finite():
            return str(value)
        return format(number, ',f')
    except (InvalidOperation, ValueError):
        return str(value)
