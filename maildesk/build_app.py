"""Maildesk 검토 화면(HTML 한 파일) 만들기

분류 결과·회신 초안·할 일·채점 결과를 app/template.html 에 넣어 app/maildesk.html 을 만든다.
만든 파일은 인터넷 없이 브라우저로 바로 열 수 있다(더블클릭).

사용법:
    python maildesk/build_app.py
    python maildesk/build_app.py --ai cli-v1-medium --rule rule-v1 --source outlook
"""
import argparse
import json
import sys
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / "connect"))

import outlook_reader  # noqa: E402
from build_tasks import build_tasks  # noqa: E402
from run_eval import DATA, FIELDS, RESULTS, evaluate  # noqa: E402

CAVEAT = ("가상 메일 30건 기준. 메일·정답표·규칙을 모두 Claude가 만들었고 AI 분류기도 Claude 계열이라 "
          "두 점수 모두 실제보다 좋게 나왔을 수 있습니다. 사람이 새로 쓴 별도 테스트 메일로 재검증이 필요합니다.")


def load_preds(name):
    return json.loads((RESULTS / f"{name}_predictions.json").read_text(encoding="utf-8"))["predictions"]


def summarize(name, truth, preds):
    n, acc, _errors, kpi, _cats, _conf = evaluate(truth, preds)
    return n, {"name": name, "acc": acc, "kpi": {k: [v[0], v[1]] for k, v in kpi.items()}}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--ai", default="cli-v1-medium")
    ap.add_argument("--rule", default="rule-v1")
    ap.add_argument("--source", choices=["outlook", "json"], default="json")
    args = ap.parse_args()

    data = json.loads(DATA.read_text(encoding="utf-8"))
    ref_str = data["meta"]["reference_time"].split(" ")[0]
    ref = datetime.fromisoformat(ref_str)
    truth = {m["id"]: m["answer"] for m in data["emails"]}
    mails = outlook_reader.read_from_outlook() if args.source == "outlook" else outlook_reader.read_from_json()
    mails = sorted(mails, key=lambda m: m["received_at"], reverse=True)
    ai, rule = load_preds(args.ai), load_preds(args.rule)
    tasks = build_tasks(ai, {m["id"]: m for m in mails}, ref)

    drafts_path = RESULTS / f"drafts_{args.ai}.json"
    drafts = json.loads(drafts_path.read_text(encoding="utf-8"))["drafts"] if drafts_path.exists() else {}

    n, rule_sum = summarize(args.rule, truth, rule)
    _, ai_sum = summarize(args.ai, truth, ai)

    payload = {
        "ref": ref_str,
        "mails": mails,
        "preds": ai,
        "rulePreds": rule,
        "drafts": drafts,
        "tasks": tasks,
        "taskAlerts": {t["id"]: t["alerts"] for t in tasks},
        "eval": {"n": n, "fields": FIELDS, "rule": rule_sum, "ai": ai_sum, "caveat": CAVEAT},
    }
    blob = json.dumps(payload, ensure_ascii=False).replace("</", "<\\/")
    html = (ROOT / "app" / "template.html").read_text(encoding="utf-8").replace("/*__DATA__*/null", blob, 1)
    out = ROOT / "app" / "maildesk.html"
    out.write_text(html, encoding="utf-8")
    print(f"저장: {out.relative_to(ROOT.parent)} (메일 {len(mails)}건, 초안 {len(drafts)}건, 할 일 {len(tasks)}건)")


if __name__ == "__main__":
    main()
