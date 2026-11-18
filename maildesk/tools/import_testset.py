"""사람이 작성한 새테스트메일_양식.xlsx → data/sample_emails_v2.json (기존 30건과 같은 형식)

사용법:
    python maildesk/tools/import_testset.py maildesk/eval_kit/새테스트메일_양식.xlsx
그다음:
    python maildesk/connect/seed_test_folder.py --data maildesk/data/sample_emails_v2.json --all
    python maildesk/run_eval.py --data maildesk/data/sample_emails_v2.json                     # 규칙
    python maildesk/run_eval.py --data maildesk/data/sample_emails_v2.json --classifier cli    # AI
"""
import json
import re
import sys
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parent.parent
V1 = ROOT / "data" / "sample_emails_v1.json"
OUT = ROOT / "data" / "sample_emails_v2.json"


def text(v):
    if v is None:
        return ""
    if isinstance(v, datetime):
        return v.strftime("%Y-%m-%d %H:%M")
    return str(v).strip()


def parse_items(s):
    items = []
    for line in filter(None, (x.strip() for x in s.splitlines())):
        m = re.match(r"^(\S+)\s*(?:[×xX*]\s*([\d,]+))?", line)
        items.append({"part_no": m[1], "qty": int(m[2].replace(",", "")) if m and m[2] else None})
    return items


def main(path, out=OUT):
    out = Path(out)
    wb = load_workbook(path, data_only=True)
    guide = {text(r[0].value): text(r[1].value) for r in wb["작성 안내"].iter_rows() if r[0].value}
    ref = guide.get("기준 시각") or "2026-11-25 10:00"
    rules = json.loads(V1.read_text(encoding="utf-8"))["meta"]["label_rules"]

    emails, problems = [], []
    for row in wb["메일 작성"].iter_rows(min_row=2, values_only=True):
        (mid, recv, name, email, company, subject, body,
         c1, c2, reply, urg, items, due, ev, notes) = (list(row) + [None] * 15)[:15]
        if not text(body):
            continue
        mid, body = text(mid), text(body).replace("\r\n", "\n")
        a = {"cat1": text(c1), "cat2": text(c2) or "-", "reply_needed": text(reply), "urgency": text(urg),
             "items": parse_items(text(items)), "due_date": text(due)[:10] or None, "due_text": None,
             "request_summary": "", "evidence": text(ev), "notes": text(notes)}
        if a["cat1"] not in rules["cat1"]:
            problems.append(f"{mid}: 1차 분류 비어 있음/잘못됨")
        if a["cat2"] not in rules["cat2"]:
            problems.append(f"{mid}: 2차 분류 잘못됨")
        if a["reply_needed"] not in ("Y", "N"):
            problems.append(f"{mid}: 회신 필요 비어 있음")
        if a["urgency"] not in rules["urgency"]:
            problems.append(f"{mid}: 긴급도 비어 있음")
        if a["evidence"] and a["evidence"] not in body:
            problems.append(f"{mid}: 근거 문장이 본문과 다름(그대로 복사해 주세요)")
        try:
            received = datetime.strptime(text(recv), "%Y-%m-%d %H:%M").strftime("%Y-%m-%dT%H:%M")
        except ValueError:
            problems.append(f"{mid}: 받은 시각 형식 오류 → YYYY-MM-DD HH:MM")
            received = None
        emails.append({"id": mid, "received_at": received, "from_name": text(name), "from_email": text(email),
                       "company": text(company), "subject": text(subject), "body": body, "answer": a})

    if problems:
        print("확인이 필요한 칸:")
        print("\n".join("  " + p for p in problems))
        sys.exit(1)
    ref_iso = datetime.strptime(ref, "%Y-%m-%d %H:%M").strftime("%Y-%m-%dT%H:%M")
    meta = {"dataset": "Maildesk 사람 작성 테스트 메일 v2", "version": "2.0", "created": datetime.now().strftime("%Y-%m-%d"),
            "status": "사람이 작성한 별도 테스트셋 (규칙·프롬프트 수정에 쓰지 않음)",
            "reference_time": ref_iso + " — 긴급도·상대 기한 계산 기준", "label_rules": rules}
    out.write_text(json.dumps({"meta": meta, "emails": emails}, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"저장: {out} ({len(emails)}건)")


if __name__ == "__main__":
    main(*sys.argv[1:3])
