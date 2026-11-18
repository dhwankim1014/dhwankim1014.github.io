"""회신이 필요한 메일에 대해 회신 초안 3종을 만든다 (Claude Code CLI, 구독 사용).

사용법:
    python maildesk/run_drafts.py                         # AI 분류 결과에서 회신 필요(Y)인 메일 전부
    python maildesk/run_drafts.py --only M01 M03          # 특정 메일만
    python maildesk/run_drafts.py --pred cli-v1-medium    # 사용할 분류 결과 이름

결과: maildesk/results/drafts_<pred>.json  (이미 만든 메일은 건너뜀 — 다시 만들려면 --redo)
"""
import argparse
import json
import sys
import tempfile
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / "connect"))
sys.path.insert(0, str(ROOT / "classify"))

import cli  # noqa: E402
import drafts  # noqa: E402
import outlook_reader  # noqa: E402

RESULTS = ROOT / "results"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--pred", default="cli-v1-medium")
    ap.add_argument("--source", choices=["outlook", "json"], default="outlook")
    ap.add_argument("--only", nargs="+")
    ap.add_argument("--redo", action="store_true")
    ap.add_argument("--effort", default="medium", choices=["low", "medium", "high"])
    ap.add_argument("--recheck", action="store_true", help="AI 호출 없이 저장된 초안의 숫자 검사만 다시")
    args = ap.parse_args()

    preds = json.loads((RESULTS / f"{args.pred}_predictions.json").read_text(encoding="utf-8"))["predictions"]
    mails = {m["id"]: m for m in (outlook_reader.read_from_outlook() if args.source == "outlook"
                                  else outlook_reader.read_from_json())}
    targets = args.only or [i for i in sorted(preds) if preds[i].get("reply_needed") == "Y"]

    out_path = RESULTS / f"drafts_{args.pred}.json"
    store = json.loads(out_path.read_text(encoding="utf-8")) if out_path.exists() else {"drafts": {}}
    if args.recheck:
        for i, e in store["drafts"].items():
            raw = {"drafts": e["drafts"], "recommended": e["recommended"],
                   "recommend_reason": e["recommend_reason"], "to_confirm": e["to_confirm"]}
            e.update(drafts.postprocess_drafts(raw, mails[i], preds[i]))
            if e["warnings"]:
                print(f"  {i} ⚠ {e['warnings']}")
        out_path.write_text(json.dumps(store, ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"재검사 완료: {len(store['drafts'])}건, 경고 {sum(bool(e['warnings']) for e in store['drafts'].values())}건")
        return
    exe, workdir = cli.find_claude(), tempfile.mkdtemp(prefix="maildesk-draft-")
    total_cost = 0.0

    for k, i in enumerate(targets, 1):
        if i in store["drafts"] and not args.redo:
            print(f"  [{k}/{len(targets)}] {i} 건너뜀(이미 있음)")
            continue
        t0 = time.time()
        res = cli.call_claude(exe, drafts.DRAFT_SYSTEM, drafts.TASK,
                              drafts.build_draft_input(mails[i], preds[i]),
                              drafts.DRAFT_SCHEMA, workdir, effort=args.effort)
        total_cost += res.get("total_cost_usd") or 0
        if "_error" in res:
            print(f"  [{k}/{len(targets)}] {i} 실패: {res['_error']}")
            continue
        entry = drafts.postprocess_drafts(res["structured_output"], mails[i], preds[i])
        entry["model"] = ",".join((res.get("modelUsage") or {}).keys())
        entry["seconds"] = round(time.time() - t0, 1)
        store["drafts"][i] = entry
        out_path.write_text(json.dumps(store, ensure_ascii=False, indent=2), encoding="utf-8")  # 1건마다 저장
        warn = f" ⚠ {entry['warnings']}" if entry["warnings"] else ""
        print(f"  [{k}/{len(targets)}] {i} 완료 ({entry['seconds']}초, 추천: {entry['recommended']}){warn}", flush=True)

    print(f"저장: {out_path.relative_to(ROOT.parent)} / 이번 실행 API 환산 추정 ${total_cost:.3f}")


if __name__ == "__main__":
    main()
