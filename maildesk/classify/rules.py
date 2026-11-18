"""Maildesk 규칙 기반 분류기 v1 (기준선 / before)

AI 없이 키워드와 정규식만으로 메일을 분류하고 정보를 뽑는다.
나중에 AI(LLM) 분류기와 같은 정답표로 비교하기 위한 '기준선'이다.

의도적으로 '실무자가 처음 떠올릴 법한 단순한 규칙'으로 만들었다.
정답표에 맞추려고 메일별 예외 규칙을 넣지 않는다(과적합 방지).
"""
import re
from datetime import date, datetime, timedelta

RULE_VERSION = "rule-v1"

# ── 1차 분류 단서 ──────────────────────────────────────────
AD_CUES = ["(광고)", "광고", "수신거부", "뉴스레터", "구독", "사전등록", "무료", "할인", "이벤트",
           "계정 확인", "계정을 확인", "비밀번호"]
NOTICE_SUBJECT_CUES = ["[공지]", "공지", "알림", "안내", "리포트"]
NOTICE_BODY_CUES = ["발신 전용", "자동 발송", "회신하지 않으셔도"]

# ── 2차 분류 단서 (업무 메일만) ─────────────────────────────
CAT2_CUES = {
    "견적요청": ["견적", "단가", "rfq", "quotation", "quote"],
    "납기문의": ["납기", "출하", "delivery status", "언제"],
    "납품일정": ["입고", "납품", "배차", "배송", "도착", "팔레트"],
    "주문·변경": ["변경", "취소", "발주", "new po", "주문"],
    "일반문의": ["성적서", "카탈로그", "휴무", "자료"],
}
CAT2_ORDER = list(CAT2_CUES)  # 점수가 같으면 앞쪽 우선

# ── 회신 필요 / 긴급 단서 ─────────────────────────────────
REPLY_CUES = ["부탁", "회신", "알려주", "주실 수", "주세요", "있을까요", "궁금합니다",
              "please", "could you", "kindly", "advise"]
NO_REPLY_CUES = ["회신하지 않으셔도", "발신 전용", "자동 발송", "경우에만"]
URGENT_CUES = ["긴급", "취소", "line will stop", "라인", "as soon as possible", "asap",
               "가능한 빨리", "오늘 중"]
DEADLINE_CUES = ["까지", "전까지", "중으로", "by ", "within"]

PART_RE = re.compile(r"\bHG-[A-Z]{3}-\d+\b")
QTY_RE = re.compile(r"([\d,]+)\s*(개|pcs)", re.I)

MONTHS = {m: i + 1 for i, m in enumerate(
    ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"])}
WEEKDAYS_KR = {"월": 0, "화": 1, "수": 2, "목": 3, "금": 4, "토": 5, "일": 6}


def _has(text, cues):
    t = text.lower()
    return [c for c in cues if c.lower() in t]


def _sentences(body):
    parts = re.split(r"\n|(?<=[.!?])\s+", body)
    return [p.strip() for p in parts if p.strip()]


def add_business_days(d, n):
    while n > 0:
        d += timedelta(days=1)
        if d.weekday() < 5:
            n -= 1
    return d


def _fix_year(month, day, base):
    d = date(base.year, month, day)
    if d < base - timedelta(days=60):
        d = date(base.year + 1, month, day)
    return d


def find_dates(text, base):
    """문장에서 날짜 표현을 찾아 (위치, 날짜, 원문) 목록으로 돌려준다. base = 메일 수신일."""
    found = []
    for m in re.finditer(r"(\d{1,2})월\s*(\d{1,2})일", text):
        found.append((m.start(), _fix_year(int(m[1]), int(m[2]), base), m[0]))
    for m in re.finditer(r"(?<![\d/-])(\d{1,2})/(\d{1,2})(?![\d/])", text):
        if 1 <= int(m[1]) <= 12 and 1 <= int(m[2]) <= 31:
            found.append((m.start(), _fix_year(int(m[1]), int(m[2]), base), m[0]))
    for m in re.finditer(r"\b(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\.?\s+(\d{1,2})\b", text, re.I):
        found.append((m.start(), _fix_year(MONTHS[m[1].lower()], int(m[2]), base), m[0]))
    for m in re.finditer(r"오늘", text):
        found.append((m.start(), base, m[0]))
    for m in re.finditer(r"내일", text):
        found.append((m.start(), base + timedelta(days=1), m[0]))
    for m in re.finditer(r"금주|이번\s*주", text):
        found.append((m.start(), base + timedelta(days=4 - base.weekday()), m[0]))
    for m in re.finditer(r"다음\s*주\s*([월화수목금토일])요일", text):
        next_monday = base + timedelta(days=7 - base.weekday())
        found.append((m.start(), next_monday + timedelta(days=WEEKDAYS_KR[m[1]]), m[0]))
    for m in re.finditer(r"end of next week", text, re.I):
        next_monday = base + timedelta(days=7 - base.weekday())
        found.append((m.start(), next_monday + timedelta(days=4), m[0]))
    for m in re.finditer(r"within\s+(\d+)\s+business days|(\d+)\s*영업일", text, re.I):
        n = int(m[1] or m[2])
        found.append((m.start(), add_business_days(base, n), m[0]))
    return sorted(found, key=lambda x: x[0])


def extract_due(body, base):
    """기한 단서(까지/by/within 등)가 있는 문장의 첫 날짜를 우선, 없으면 본문 첫 날짜."""
    first_any = None
    for s in _sentences(body):
        ds = find_dates(s, base)
        if not ds:
            continue
        if first_any is None:
            first_any = (ds[0], s)
        if _has(s, DEADLINE_CUES):
            return ds[0][1], ds[0][2], s
    if first_any:
        (_, d, txt), s = first_any
        return d, txt, s
    return None, None, None


def extract_items(body):
    items, seen = [], set()
    for line in body.split("\n"):
        for m in PART_RE.finditer(line):
            pn = m.group(0)
            if pn in seen:
                continue
            seen.add(pn)
            q = QTY_RE.search(line, m.end())
            items.append({"part_no": pn, "qty": int(q[1].replace(",", "")) if q else None})
    return items


def classify(mail, reference_time):
    """mail: outlook_reader 형식 dict. reference_time: 긴급도 계산 기준 datetime."""
    subj, body = mail["subject"], mail["body"]
    text = subj + "\n" + body
    base = datetime.fromisoformat(mail["received_at"]).date()
    reasons = []

    # 1차 분류
    ad = _has(text, AD_CUES)
    notice = _has(subj, NOTICE_SUBJECT_CUES) + _has(body, NOTICE_BODY_CUES)
    if ad:
        cat1 = "광고·기타"; reasons.append(f"광고 단서: {', '.join(ad)}")
    elif notice:
        cat1 = "참고·알림"; reasons.append(f"알림 단서: {', '.join(notice)}")
    else:
        cat1 = "업무"; reasons.append("광고/알림 단서 없음 → 업무")

    pred = {"cat1": cat1, "cat2": "-", "reply_needed": "N", "urgency": "하", "items": [],
            "due_date": None, "due_text": None, "evidence": None}

    if cat1 != "업무":
        hit = (ad or notice)[0]
        pred["evidence"] = next((s for s in _sentences(text) if hit.lower() in s.lower()), None)
        pred["reasons"] = reasons
        return pred

    # 2차 분류: 제목 키워드 2점, 본문 1점
    scores = {c: 2 * len(_has(subj, cues)) + len(_has(body, cues)) for c, cues in CAT2_CUES.items()}
    best = max(CAT2_ORDER, key=lambda c: (scores[c], -CAT2_ORDER.index(c)))
    pred["cat2"] = best if scores[best] > 0 else "일반문의"
    reasons.append("2차 점수: " + ", ".join(f"{c} {s}" for c, s in scores.items() if s))

    # 회신 필요
    req, noreq = _has(body, REPLY_CUES), _has(body, NO_REPLY_CUES)
    pred["reply_needed"] = "Y" if req and not noreq else "N"
    reasons.append(f"회신 단서: {req or '없음'}" + (f" / 회신 불필요 단서: {noreq}" if noreq else ""))

    # 품번·수량, 기한
    pred["items"] = extract_items(body)
    due, due_text, due_sentence = extract_due(body, base)
    pred["due_date"] = due.isoformat() if due else None
    pred["due_text"] = due_text

    # 긴급도
    urgent = _has(text, URGENT_CUES)
    limit = add_business_days(reference_time.date(), 2)
    if pred["reply_needed"] == "N":
        pred["urgency"] = "하"
    elif urgent or (due and due <= limit):
        pred["urgency"] = "상"
        reasons.append(f"긴급: {urgent or ''} {('기한 ' + due.isoformat()) if due and due <= limit else ''}".strip())
    else:
        pred["urgency"] = "중"

    pred["evidence"] = next((s for s in _sentences(body) if _has(s, REPLY_CUES)), due_sentence)
    pred["reasons"] = reasons
    return pred
