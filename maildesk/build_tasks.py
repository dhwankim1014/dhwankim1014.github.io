"""Maildesk 할 일 목록 · 기한 알림 · 엑셀 업무 대장

AI 분류 결과(업무 메일)를 '누구에게 / 무엇을 / 언제까지' 표로 바꾸고, 알림 규칙으로 경고를 붙인다.

알림 규칙 (가정 — 실제 회사 규정에 맞게 바꿔야 함)
- 기한 지남      : 기한 < 기준일
- 오늘 마감      : 기한 = 기준일
- 기한 임박      : 기한 ≤ 기준일 + 1영업일
- 미답변 24h 초과 : 회신 필요 메일이 받은 지 24시간 넘게 지남
- 사람 확인 필요  : AI가 확신이 낮거나 근거 검증에 실패한 메일

사용법:
    python maildesk/build_tasks.py                    # results/tasks_<pred>.json + 업무대장_<pred>.xlsx
"""
import argparse
import json
import sys
from datetime import datetime, timedelta
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / "connect"))
RESULTS = ROOT / "results"
DATA = ROOT / "data" / "sample_emails_v1.json"
URG_ORDER = {"상": 0, "중": 1, "하": 2}
SLA_HOURS = 24


def add_bd(d, n):
    while n > 0:
        d += timedelta(days=1)
        if d.weekday() < 5:
            n -= 1
    return d


def alerts_for(pred, mail, ref):
    out = []
    due = datetime.fromisoformat(pred["due_date"]).date() if pred.get("due_date") else None
    today = ref.date()
    if due:
        if due < today:
            out.append("기한 지남")
        elif due == today:
            out.append("오늘 마감")
        elif due <= add_bd(today, 1):
            out.append("기한 임박")
    hours = (ref - datetime.fromisoformat(mail["received_at"])).total_seconds() / 3600
    if pred.get("reply_needed") == "Y" and hours > SLA_HOURS:
        out.append(f"미답변 {int(hours)}시간")
    if pred.get("needs_review"):
        out.append("사람 확인 필요")
    return out, due, hours


def build_tasks(preds, mails, ref):
    tasks = []
    for i, p in preds.items():
        if p.get("cat1") != "업무":
            continue
        m = mails[i]
        alerts, due, hours = alerts_for(p, m, ref)
        tasks.append({
            "id": i,
            "company": m["company"],
            "contact": m["from_name"],
            "subject": m["subject"],
            "type": p["cat2"],
            "todo": p.get("request_summary"),
            "items": p.get("items") or [],
            "due_date": p.get("due_date"),
            "due_text": p.get("due_text"),
            "d_day": (due - ref.date()).days if due else None,
            "urgency": p["urgency"],
            "reply_needed": p["reply_needed"],
            "received_at": m["received_at"],
            "hours_since": round(hours, 1),
            "alerts": alerts,
            "status": "미처리",
        })
    tasks.sort(key=lambda t: (URG_ORDER.get(t["urgency"], 9), t["due_date"] or "9999", t["received_at"]))
    return tasks


def write_excel(tasks, path, ref):
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill

    head = Font(bold=True, color="FFFFFF")
    fill = PatternFill("solid", fgColor="2F5496")
    red = PatternFill("solid", fgColor="F8D7DA")
    amber = PatternFill("solid", fgColor="FFF3CD")
    wrap = Alignment(wrap_text=True, vertical="top")

    wb = Workbook()
    ws = wb.active
    ws.title = "할 일 목록"
    cols = ["상태", "메일", "긴급도", "기한", "D-day", "거래처", "담당자", "유형", "할 일", "품번 × 수량",
            "회신 필요", "받은 지(시간)", "알림", "처리 메모"]
    ws.append(cols)
    for t in tasks:
        ws.append([t["status"], t["id"], t["urgency"], t["due_date"] or "", "" if t["d_day"] is None else t["d_day"],
                   t["company"], t["contact"], t["type"], t["todo"],
                   "\n".join(f"{x['part_no']} × {x['qty'] if x['qty'] is not None else '?'}" for x in t["items"]),
                   t["reply_needed"], t["hours_since"], ", ".join(t["alerts"]), ""])
    for c in ws[1]:
        c.font, c.fill, c.alignment = head, fill, wrap
    for row in ws.iter_rows(min_row=2):
        alerts = row[12].value or ""
        for c in row:
            c.alignment = wrap
        if any(a in alerts for a in ("기한 지남", "오늘 마감")):
            for c in row:
                c.fill = red
        elif "기한 임박" in alerts or "미답변" in alerts:
            for c in row:
                c.fill = amber
    for col, w in zip("ABCDEFGHIJKLMN", (8, 6, 7, 12, 7, 18, 12, 10, 40, 22, 8, 10, 30, 24)):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "C2"
    ws.auto_filter.ref = ws.dimensions

    ws2 = wb.create_sheet("품번별 요청")
    ws2.append(["품번", "수량", "거래처", "유형", "기한", "메일"])
    for c in ws2[1]:
        c.font, c.fill = head, fill
    rows = [(x["part_no"], x["qty"], t["company"], t["type"], t["due_date"] or "", t["id"])
            for t in tasks for x in t["items"]]
    for r in sorted(rows, key=lambda r: (r[0], r[4] or "9999")):
        ws2.append(list(r))
    for col, w in zip("ABCDEF", (14, 8, 20, 10, 12, 6)):
        ws2.column_dimensions[col].width = w
    ws2.auto_filter.ref = ws2.dimensions

    ws3 = wb.create_sheet("안내")
    for line in [
        f"기준 시각: {ref:%Y-%m-%d %H:%M} (가상 시나리오)",
        "AI 분류 결과로 자동 생성한 업무 대장입니다. 최종 판단은 담당자가 합니다.",
        "알림 규칙(가정): 기한 지남 / 오늘 마감 / 기한 임박(1영업일) / 미답변 24시간 초과 / 사람 확인 필요",
        "빨강 = 기한 지남·오늘 마감, 노랑 = 기한 임박·미답변",
    ]:
        ws3.append([line])
    ws3.column_dimensions["A"].width = 100
    wb.save(path)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--pred", default="cli-v1-medium")
    ap.add_argument("--source", choices=["outlook", "json"], default="json")
    args = ap.parse_args()
    import outlook_reader

    meta = json.loads(DATA.read_text(encoding="utf-8"))["meta"]
    ref = datetime.fromisoformat(meta["reference_time"].split(" ")[0])
    preds = json.loads((RESULTS / f"{args.pred}_predictions.json").read_text(encoding="utf-8"))["predictions"]
    mails = {m["id"]: m for m in (outlook_reader.read_from_outlook() if args.source == "outlook"
                                  else outlook_reader.read_from_json())}
    tasks = build_tasks(preds, mails, ref)
    (RESULTS / f"tasks_{args.pred}.json").write_text(json.dumps(tasks, ensure_ascii=False, indent=2), encoding="utf-8")
    xlsx = RESULTS / f"업무대장_{args.pred}.xlsx"
    write_excel(tasks, xlsx, ref)
    print(f"할 일 {len(tasks)}건 / 경고 있는 건 {sum(bool(t['alerts']) for t in tasks)}건 → {xlsx.name}")
    for t in tasks:
        print(f"  {t['id']} {t['urgency']} {t['due_date'] or '-':10} {', '.join(t['alerts'])}")


if __name__ == "__main__":
    main()
