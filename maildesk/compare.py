"""두 분류기의 채점 결과를 나란히 비교하는 보고서를 만든다.

사용법:
    python maildesk/compare.py rule-v1 cli-v1-medium
결과: maildesk/results/compare_<A>_vs_<B>.md
"""
import json
import sys
from pathlib import Path

from run_eval import DATA, FIELDS, RESULTS, evaluate, fmt, same


def load(name):
    return json.loads((RESULTS / f"{name}_predictions.json").read_text(encoding="utf-8"))["predictions"]


def main(a, b):
    truth = {m["id"]: m["answer"] for m in json.loads(DATA.read_text(encoding="utf-8"))["emails"]}
    pa, pb = load(a), load(b)
    ra, rb = evaluate(truth, pa), evaluate(truth, pb)
    n = ra[0]

    L = [f"# 분류기 비교: {a} vs {b}", "",
         f"- 대상: 가상 업무 메일 {n}건 (Outlook 테스트 폴더) / 정답표: 사용자 검토 전 초안",
         "- 주의: 메일·정답표·규칙을 모두 Claude가 만들었고 AI 분류기도 Claude 계열이라, 두 점수 모두 실제보다"
         " 좋게 나왔을 수 있음. 사람이 새로 쓴 별도 테스트 메일로 다시 확인해야 포트폴리오 수치로 쓸 수 있음.", "",
         "## 항목별 정확도", "", f"| 항목 | {a} | {b} |", "|---|---|---|"]
    for f, label in FIELDS:
        L.append(f"| {label} | {ra[1][f]}/{n} ({ra[1][f] / n:.0%}) | {rb[1][f]}/{n} ({rb[1][f] / n:.0%}) |")
    L += ["", "## 업무상 중요한 실수", "", f"| 지표 | {a} | {b} |", "|---|---|---|"]
    for k in ra[3]:
        (ia, ta), (ib, tb) = ra[3][k], rb[3][k]
        L.append(f"| {k} | {len(ia)}/{ta} {', '.join(ia)} | {len(ib)}/{tb} {', '.join(ib)} |")

    L += ["", "## 메일별 차이 (한쪽이라도 틀린 항목)", "",
          f"| 메일 | 항목 | 정답 | {a} | {b} |", "|---|---|---|---|---|"]
    for i in sorted(truth):
        for f, label in FIELDS:
            ok_a, ok_b = same(f, truth[i][f], pa[i][f]), same(f, truth[i][f], pb[i][f])
            if not (ok_a and ok_b):
                mark = lambda ok, v: ("✅ " if ok else "❌ ") + fmt(f, v)
                L.append(f"| {i} | {label} | {fmt(f, truth[i][f])} | {mark(ok_a, pa[i][f])} | {mark(ok_b, pb[i][f])} |")

    flagged = [i for i in sorted(pb) if pb[i].get("needs_review")]
    if flagged:
        L += ["", f"## {b}가 '사람 확인 필요'로 표시한 메일", "", "| 메일 | 확신도 | AI가 적은 이유 |", "|---|---|---|"]
        for i in flagged:
            L.append(f"| {i} | {pb[i].get('confidence')} | {pb[i]['reasons'][0]} |")

    out = RESULTS / f"compare_{a}_vs_{b}.md"
    out.write_text("\n".join(L) + "\n", encoding="utf-8")
    print("저장:", out.relative_to(Path(__file__).resolve().parent.parent))


if __name__ == "__main__":
    main(*sys.argv[1:3])
