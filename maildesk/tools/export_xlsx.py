"""샘플 메일 JSON을 검사하고, 사람이 검토하기 쉬운 엑셀 정답표로 내보낸다.

사용법: python maildesk/tools/export_xlsx.py
"""
import json
import sys
from collections import Counter
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

ROOT = Path(__file__).resolve().parent.parent
SRC = ROOT / "data" / "sample_emails_v1.json"
OUT = ROOT / "data" / "sample_emails_v1_정답표.xlsx"

data = json.loads(SRC.read_text(encoding="utf-8"))
rules = data["meta"]["label_rules"]
emails = data["emails"]

# 1) 검사: 근거 문장이 본문에 실제로 있는지, 라벨 값이 규칙에 있는 값인지
problems = []
for m in emails:
    a = m["answer"]
    if a["evidence"] not in m["body"]:
        problems.append(f"{m['id']}: 근거 문장이 본문에 없음")
    if a["cat1"] not in rules["cat1"]:
        problems.append(f"{m['id']}: cat1 값 오류 {a['cat1']}")
    if a["cat2"] not in rules["cat2"]:
        problems.append(f"{m['id']}: cat2 값 오류 {a['cat2']}")
    if a["urgency"] not in rules["urgency"]:
        problems.append(f"{m['id']}: urgency 값 오류 {a['urgency']}")
    if (a["cat1"] == "업무") == (a["cat2"] == "-"):
        problems.append(f"{m['id']}: cat1/cat2 조합 오류")
    for it in a["items"]:
        if it["part_no"] not in m["body"]:
            problems.append(f"{m['id']}: 품번 {it['part_no']} 본문에 없음")

if problems:
    print("검사 실패:")
    print("\n".join(problems))
    sys.exit(1)

# 2) 엑셀 작성
wb = Workbook()
head_font = Font(bold=True, color="FFFFFF")
head_fill = PatternFill("solid", fgColor="2F5496")
wrap = Alignment(wrap_text=True, vertical="top")


def write_sheet(ws, headers, rows, widths):
    ws.append(headers)
    for c in ws[1]:
        c.font, c.fill, c.alignment = head_font, head_fill, wrap
    for r in rows:
        ws.append(r)
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.alignment = wrap
    for i, w in enumerate(widths):
        ws.column_dimensions[chr(65 + i)].width = w
    ws.freeze_panes = "B2"


def fmt_items(items):
    if not items:
        return ""
    return "\n".join(
        f"{it['part_no']} × {it['qty'] if it['qty'] is not None else '(수량 없음)'}" for it in items
    )


ws = wb.active
ws.title = "정답표"
write_sheet(
    ws,
    ["ID", "수신일시", "보낸 사람", "회사", "제목", "1차 분류", "2차 분류", "회신 필요",
     "긴급도", "품번 × 수량", "기한", "기한 원문", "요청 요약", "판단 근거(원문)", "함정·메모", "검토 의견"],
    [
        [m["id"], m["received_at"].replace("T", " "), m["from_name"], m["company"], m["subject"],
         m["answer"]["cat1"], m["answer"]["cat2"], m["answer"]["reply_needed"], m["answer"]["urgency"],
         fmt_items(m["answer"]["items"]), m["answer"]["due_date"] or "", m["answer"]["due_text"] or "",
         m["answer"]["request_summary"], m["answer"]["evidence"], m["answer"]["notes"], ""]
        for m in emails
    ],
    [6, 16, 14, 18, 30, 10, 10, 8, 7, 22, 12, 18, 30, 40, 36, 24],
)

ws2 = wb.create_sheet("메일 원문")
write_sheet(
    ws2,
    ["ID", "수신일시", "보낸 사람", "이메일", "제목", "본문"],
    [[m["id"], m["received_at"].replace("T", " "), m["from_name"], m["from_email"], m["subject"], m["body"]]
     for m in emails],
    [6, 16, 16, 34, 34, 80],
)

ws3 = wb.create_sheet("라벨 기준")
ws3.append(["항목", "값", "설명"])
for c in ws3[1]:
    c.font, c.fill = head_font, head_fill
ws3.append(["기준 시점", "", data["meta"]["reference_time"]])
ws3.append(["시나리오", "", data["meta"]["scenario"]])
for key in ("cat1", "cat2", "urgency"):
    for val, desc in rules[key].items():
        ws3.append([key, val, desc])
for key in ("reply_needed", "due_date", "evidence"):
    ws3.append([key, "", rules[key]])
for col, w in zip("ABC", (14, 12, 90)):
    ws3.column_dimensions[col].width = w

ws4 = wb.create_sheet("분포")
ws4.append(["구분", "값", "건수"])
for c in ws4[1]:
    c.font, c.fill = head_font, head_fill
for key, label in (("cat1", "1차 분류"), ("cat2", "2차 분류"), ("reply_needed", "회신 필요"), ("urgency", "긴급도")):
    for val, n in Counter(m["answer"][key] for m in emails).items():
        ws4.append([label, val, n])
ws4.column_dimensions["A"].width = 12

wb.save(OUT)
print(f"검사 통과: {len(emails)}건 / 저장: {OUT.name}")
for key in ("cat1", "cat2", "urgency"):
    print(key, dict(Counter(m["answer"][key] for m in emails)))
