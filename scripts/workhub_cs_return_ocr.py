from __future__ import annotations

import re
from datetime import date, datetime
from io import BytesIO
from typing import Iterable


TRACKING_PATTERN = re.compile(r"(?<!\d)(\d{4})[\s-]?(\d{4})[\s-]?(\d{4})(?!\d)")
MASKED_NAME_PATTERN = re.compile(r"(?<![가-힣])([가-힣][*＊Xx○●][가-힣])(?![가-힣])")
PHONE_PATTERN = re.compile(r"(?<!\d)(01\d|0\d{1,2})[\s-]?(\d{3,4})[\s-]?(\d{4})(?!\d)")
TOKEN_PATTERN = re.compile(r"[가-힣A-Za-z0-9]{2,}")
MATCH_STOPWORDS = {
    "상품명", "보내는분", "받는분", "운송장번호", "수량", "경기도", "서울시", "광역시",
    "대한통운", "반품", "고객", "주소", "전화", "배송", "회수", "양품",
}


def compact_text(value: object) -> str:
    return re.sub(r"[^가-힣A-Za-z0-9*＊Xx○●]", "", str(value or "")).lower()


def tracking_display(value: str) -> str:
    digits = re.sub(r"\D", "", value or "")
    if len(digits) == 12:
        return f"{digits[:4]}-{digits[4:8]}-{digits[8:]}"
    return value.strip()


def _tracking_candidates(text: str) -> list[tuple[str, int]]:
    found: list[tuple[str, int]] = []
    seen: set[str] = set()
    for match in TRACKING_PATTERN.finditer(text or ""):
        number = "-".join(match.groups())
        if number in seen:
            continue
        seen.add(number)
        found.append((number, match.start()))
    return found


def extract_return_label_fields(text: str) -> dict[str, object]:
    raw_text = str(text or "").strip()
    tracking_candidates = _tracking_candidates(raw_text)
    heading = re.search(r"운\s*송\s*장\s*번\s*호", raw_text)
    if heading and tracking_candidates:
        tracking_candidates.sort(key=lambda item: (abs(item[1] - heading.end()), item[1]))
    tracking_number = tracking_candidates[0][0] if tracking_candidates else ""
    masked_names = list(dict.fromkeys(MASKED_NAME_PATTERN.findall(raw_text)))
    phones = ["-".join(match.groups()) for match in PHONE_PATTERN.finditer(raw_text)]
    quantity_match = re.search(r"(?:수량|총)\s*[:：-]?\s*(\d+)\s*(?:개|박스|세트)?", raw_text)
    condition = ""
    for keyword in ("바닥면스크래치", "스크래치불량", "코팅불량", "구멍뚫림", "양품", "파손", "불량"):
        if compact_text(keyword) in compact_text(raw_text):
            condition = keyword
            break
    return {
        "tracking_number": tracking_number,
        "tracking_candidates": [item[0] for item in tracking_candidates],
        "masked_names": masked_names,
        "phones": phones,
        "quantity": quantity_match.group(1) if quantity_match else "",
        "condition": condition,
        "raw_text": raw_text,
    }


def _masked_name_matches(masked: str, full_name: str) -> bool:
    masked_clean = re.sub(r"\s+", "", masked or "")
    full_clean = re.sub(r"\s+", "", full_name or "")
    return (
        len(masked_clean) >= 3
        and len(full_clean) >= 2
        and masked_clean[0] == full_clean[0]
        and masked_clean[-1] == full_clean[-1]
    )


def _match_tokens(value: object) -> set[str]:
    return {
        token.lower()
        for token in TOKEN_PATTERN.findall(str(value or ""))
        if token not in MATCH_STOPWORDS and not token.isdigit()
    }


def _candidate_score(parsed: dict[str, object], case: dict[str, object]) -> int:
    raw_compact = compact_text(parsed.get("raw_text"))
    score = 0
    name = str(case.get("receiver_name") or case.get("orderer_name") or "").strip()
    if name and compact_text(name) in raw_compact:
        score += 75
    elif any(_masked_name_matches(masked, name) for masked in parsed.get("masked_names", []) or []):
        score += 50

    phone_digits = re.sub(r"\D", "", str(case.get("receiver_phone") or case.get("orderer_phone") or ""))
    if len(phone_digits) >= 4 and phone_digits[-4:] in re.sub(r"\D", "", str(parsed.get("raw_text") or "")):
        score += 55

    product_tokens = _match_tokens(case.get("product_name"))
    raw_tokens = _match_tokens(parsed.get("raw_text"))
    score += min(40, len(product_tokens & raw_tokens) * 10)

    address_tokens = _match_tokens(case.get("receiver_address"))
    score += min(20, len(address_tokens & raw_tokens) * 5)
    return score


def match_return_label(parsed: dict[str, object], cases: Iterable[dict[str, object]]) -> dict[str, object]:
    ranked = []
    for case in cases:
        score = _candidate_score(parsed, case)
        if score:
            ranked.append({**case, "match_score": score})
    ranked.sort(key=lambda item: (-int(item["match_score"]), int(item.get("id") or 0)))
    top_score = int(ranked[0]["match_score"]) if ranked else 0
    second_score = int(ranked[1]["match_score"]) if len(ranked) > 1 else 0
    matched = bool(ranked and top_score >= 60 and top_score - second_score >= 15)
    return {
        "match_status": "matched" if matched else ("review" if ranked else "unregistered"),
        "case_id": int(ranked[0].get("id") or 0) if matched else 0,
        "match_score": top_score,
        "candidates": ranked[:5],
    }


def _korean_date_title(value: str) -> str:
    try:
        parsed = datetime.strptime(value, "%Y-%m-%d").date()
    except (TypeError, ValueError):
        parsed = date.today()
    weekdays = "월화수목금토일"
    return f"★{parsed.month}월 {parsed.day}일({weekdays[parsed.weekday()]}) 택배 회수송장★"


def format_kakao_return_text(rows: Iterable[dict[str, object]], title_date: str) -> str:
    blocks = [_korean_date_title(title_date)]
    for row in rows:
        if row.get("selected") is False:
            continue
        name = str(row.get("receiver_name") or "미확인").strip()
        product = str(row.get("product_name") or "상품 미확인").strip()
        quantity = str(row.get("quantity") or "1").strip()
        condition = str(row.get("condition") or "상태 미확인").strip()
        tracking = tracking_display(str(row.get("tracking_number") or ""))
        lines = [f"{name} / {product} - {quantity}개", condition]
        if tracking:
            lines.append(f"회수송장: {tracking}")
        blocks.append("\n".join(lines))
    return "\n\n".join(blocks).strip() + "\n"


def ocr_return_label_image(data: bytes) -> dict[str, object]:
    try:
        import pytesseract  # type: ignore
        from PIL import Image, ImageEnhance, ImageOps  # type: ignore
    except Exception as exc:  # pragma: no cover - deployment dependency
        raise RuntimeError(f"OCR 구성요소를 불러오지 못했습니다: {exc}") from exc
    try:
        image = ImageOps.exif_transpose(Image.open(BytesIO(data))).convert("RGB")
    except Exception as exc:
        raise ValueError("이미지 파일을 읽을 수 없습니다.") from exc
    languages = set(pytesseract.get_languages(config=""))
    lang = "kor+eng" if {"kor", "eng"}.issubset(languages) else ("eng" if "eng" in languages else None)
    best_text = ""
    best_rank = (-1, -1)
    for angle in (0, 90, 270, 180):
        rotated = image.rotate(angle, expand=True)
        scale = min(1.0, 2200 / max(rotated.size))
        if scale < 1:
            rotated = rotated.resize(
                (
                    max(1, round(rotated.width * scale)),
                    max(1, round(rotated.height * scale)),
                ),
                Image.Resampling.LANCZOS,
            )
        gray = ImageOps.autocontrast(rotated.convert("L"))
        gray = ImageEnhance.Contrast(gray).enhance(1.4)
        gray = ImageEnhance.Sharpness(gray).enhance(2.0)
        text = pytesseract.image_to_string(gray, lang=lang, config="--psm 11")
        parsed = extract_return_label_fields(text)
        rank = (1 if parsed["tracking_number"] else 0, len(compact_text(text)))
        if rank > best_rank:
            best_rank = rank
            best_text = text
    return extract_return_label_fields(best_text)

