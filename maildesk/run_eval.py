"""Maildesk 분류 실행 + 정답표 채점

사용법:
    python maildesk/run_eval.py                  # Outlook 테스트 폴더에서 읽어 규칙 분류기 채점
    python maildesk/run_eval.py --source json    # Outlook 없이 JSON에서 읽기

결과: maildesk/results/<분류기>_predictions.json, <분류기>_report.md
나중에 AI 분류기도 같은 채점 함수로 비교한다.
"""
import argparse
import json
import sys
from collections import Counter
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / "connect"))
sys.path.insert(0, str(ROOT / "classify"))

import outlook_reader  # noqa: E402
import rules  # noqa: E402

DATA = ROOT / "data" / "sample_emails_v1.json"
RESULTS = ROOT / "results"
FIELDS = [("cat1", "1차 분류"), ("cat2", "2차 분류"), ("reply_needed", "회신 필요"),
          ("urgency", "긴급도"), ("due_date", "기한"), ("items", "품번·수량")]


def items_key(items):
    return sorted((i["part_no"], i["qty"]) for i in items)


def same(field, a, p):
    return items_key(a) == items_key(p) if field == "items" else a == p


def fmt(field, v):
    if field == "items":
        return ", ".join(f"{i['part_no']}×{i['qty'] if i['qty'] is not None else '?'}" for i in v) or "(없음)"
    return "(없음)" if v is None else str(v)


def evaluate(truth, preds):
    n = len(truth)
    acc = {f: sum(same(f, truth[i][f], preds[i][f]) for i in truth) for f, _ in FIELDS}
    errors = [(i, label, fmt(f, truth[i][f]), fmt(f, preds[i][f]))
              for i in sorted(truth) for f, label in FIELDS if not same(f, truth[i][f], preds[i][f])]
    need = [i for i in truth if truth[i]["reply_needed"] == "Y"]
    urgent = [i for i in truth if truth[i]["urgency"] == "상"]
    kpi = {
        "회신 필요 메일 놓침": ([i for i in need if preds[i]["reply_needed"] == "N"], len(need)),
        "긴급(상) 메일 놓침": ([i for i in urgent if preds[i]["urgency"] != "상"], len(urgent)),
        "회신 불필요인데 회신 필요로 표시": ([i for i in truth if truth[i]["reply_needed"] == "N"
                                    and preds[i]["reply_needed"] == "Y"],
                                   sum(truth[i]["reply_needed"] == "N" for i in truth)),
    }
    cats = ["업무", "참고·알림", "광고·기타"]
    conf = Counter((truth[i]["cat1"], preds[i]["cat1"]) for i in truth)
    return n, acc, errors, kpi, cats, conf


def write_report(name, source, n, acc, errors, kpi, cats, conf, extra=()):
    L = [f"# {name} 채점 결과", "",
         f"- 실행 시각: {datetime.now():%Y-%m-%d %H:%M}",
         f"- 입력: {source} / 메일 {n}건 / 정답표: {'사람 작성 테스트셋' if '-v2' in name else 'data/sample_emails_v1.json (사용자 검토 전 초안)'}",
         "- 주의: 샘플 메일과 규칙을 같은 쪽(Claude)이 만들었으므로 점수가 실제보다 좋게 나올 수 있음."
         " 새로 만든 별도 테스트 메일로 다시 확인이 필요함.", "",
         "## 항목별 정확도", "", "| 항목 | 맞음 | 정확도 |", "|---|---|---|"]
    for f, label in FIELDS:
        L.append(f"| {label} | {acc[f]}/{n} | {acc[f] / n:.0%} |")
    L += ["", "## 업무상 중요한 실수", "", "| 지표 | 건수 | 해당 메일 |", "|---|---|---|"]
    for k, (ids, total) in kpi.items():
        L.append(f"| {k} | {len(ids)}/{total} | {', '.join(ids) or '-'} |")
    L += ["", "## 1차 분류 혼동표 (행 = 정답, 열 = 예측)", "",
          "| 정답 \\ 예측 | " + " | ".join(cats) + " |", "|---|" + "---|" * len(cats)]
    for t in cats:
        L.append(f"| {t} | " + " | ".join(str(conf[(t, p)]) for p in cats) + " |")
    L += [""] + list(extra)
    L += [f"## 틀린 항목 전체 ({len(errors)}개)", "", "| 메일 | 항목 | 정답 | 예측 |", "|---|---|---|---|"]
    for e in errors:
        L.append("| " + " | ".join(e) + " |")
    path = RESULTS / f"{name}_report.md"
    path.write_text("\n".join(L) + "\n", encoding="utf-8")
    return path


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--source", choices=["outlook", "json"], default="outlook")
    ap.add_argument("--classifier", choices=["rules", "llm", "cli"], default="rules",
                    help="rules=규칙, llm=Claude API, cli=Claude Code CLI(구독)")
    ap.add_argument("--model", help="cli 방식에서 쓸 모델 (예: opus). 생략하면 계정 기본 모델")
    ap.add_argument("--limit", type=int, help="앞에서부터 N건만 (AI 비용 확인용 시험 실행)")
    ap.add_argument("--effort", choices=["low", "medium", "high"], default="medium", help="AI 생각 깊이")
    ap.add_argument("--data", type=Path, default=DATA, help="정답표 JSON (기본: sample_emails_v1.json)")
    args = ap.parse_args()

    data = json.loads(args.data.read_text(encoding="utf-8"))
    ref = datetime.fromisoformat(data["meta"]["reference_time"].split(" ")[0])
    truth_all = {m["id"]: m["answer"] for m in data["emails"]}

    mails = outlook_reader.read_from_outlook() if args.source == "outlook" else outlook_reader.read_from_json(args.data)
    if not args.limit:
        missing = set(truth_all) - {m["id"] for m in mails}
        if missing:
            print("정답표에 있는데 읽지 못한 메일:", sorted(missing))
            sys.exit(1)
    mails = [m for m in mails if m["id"] in truth_all][: args.limit or None]

    extra = []
    if args.classifier == "rules":
        name = rules.RULE_VERSION
        preds = {m["id"]: rules.classify(m, ref) for m in mails}
    else:
        if args.classifier == "llm":
            import llm
            clf = llm.LLMClassifier(data["meta"]["label_rules"], data["meta"]["reference_time"], effort=args.effort)
            name = f"{llm.CLASSIFIER_VERSION}-{args.effort}"
            cost_note = "API 요금"
        else:
            import cli
            clf = cli.CLIClassifier(data["meta"]["label_rules"], data["meta"]["reference_time"],
                                    effort=args.effort, model=args.model)
            name = f"{cli.CLASSIFIER_VERSION}-{args.effort}"
            cost_note = "CLI가 계산한 API 환산 추정치 — 구독 사용 시 별도 청구 없음"
        preds = {}
        for k, m in enumerate(mails, 1):
            preds[m["id"]] = clf.classify(m)
            print(f"  [{k}/{len(mails)}] {m['id']} → {preds[m['id']]['cat1']} / {preds[m['id']]['cat2']}", flush=True)
        u = clf.usage
        flagged = {i: p["checks"] for i, p in preds.items() if p["checks"]}
        review = [i for i, p in preds.items() if p["needs_review"]]
        extra = [
            "## AI 사용량·비용", "",
            f"- 모델: {dict(clf.served_models)} / effort: {args.effort}",
            f"- 토큰: 입력 {u['input']:,} · 출력(생각 포함) {u['output']:,} · 캐시 읽기 {u['cache_read']:,} · 캐시 쓰기 {u['cache_write']:,}",
            f"- 추정 비용: ${clf.cost_usd():.3f} (메일 1건당 ${clf.cost_usd() / len(mails):.4f}) — {cost_note}",
            "", "## 코드 재검증 (AI가 지어낸 근거·품번 잡기)", "",
            f"- 근거/품번 검증 실패: {flagged if flagged else '없음'}",
            f"- '사람 확인 필요' 표시: {len(review)}건 {review}", "",
        ]
    if args.data.resolve() != DATA.resolve():
        name += "-" + args.data.stem.replace("sample_emails_", "")
    if args.limit:
        name += f"-sample{args.limit}"

    truth = {i: truth_all[i] for i in preds}
    RESULTS.mkdir(exist_ok=True)
    (RESULTS / f"{name}_predictions.json").write_text(
        json.dumps({"classifier": name, "source": args.source, "predictions": preds},
                   ensure_ascii=False, indent=2), encoding="utf-8")
    n, acc, errors, kpi, cats, conf = evaluate(truth, preds)
    report = write_report(name, args.source, n, acc, errors, kpi, cats, conf, extra)
    for line in extra:
        if line.startswith("- "):
            print(line[2:])

    for f, label in FIELDS:
        print(f"{label:8} {acc[f]:2}/{n}  {acc[f] / n:.0%}")
    for k, (ids, total) in kpi.items():
        print(f"{k}: {len(ids)}/{total} {ids}")
    print("보고서:", report.relative_to(ROOT.parent))


if __name__ == "__main__":
    main()
