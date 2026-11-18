"""Maildesk → Outlook 반영: 분류 색깔 태그(범주)와 기한 깃발

- 'Maildesk-테스트' 폴더 안의 가상 메일(MDId가 있는 항목)에만 적용한다. 다른 폴더는 건드리지 않는다.
- 메일을 보내거나 지우거나 옮기지 않는다. 범주(Categories)와 후속 작업 깃발만 바꾼다.
- 기본은 '미리보기'(아무것도 바꾸지 않음). 실제로 바꾸려면 --apply, 되돌리려면 --undo.

사용법 (클래식 Outlook을 켜 둔 상태에서):
    python maildesk/connect/outlook_apply.py                 # 미리보기
    python maildesk/connect/outlook_apply.py --apply         # 적용
    python maildesk/connect/outlook_apply.py --undo          # Maildesk가 붙인 범주·깃발 제거
"""
import argparse
import json
import sys
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
RESULTS = ROOT / "results"
PREFIX = "Maildesk"
OL_MARK_NO_DATE = 4


def planned(pred):
    """메일 1건에 붙일 범주 목록과 깃발(기한) 계획."""
    cats = [f"{PREFIX}: {pred['cat2'] if pred['cat1'] == '업무' else pred['cat1']}"]
    if pred.get("urgency") == "상":
        cats.append(f"{PREFIX}: 긴급")
    if pred.get("needs_review"):
        cats.append(f"{PREFIX}: 사람 확인")
    flag = None
    if pred.get("reply_needed") == "Y":
        flag = {"request": "회신 필요", "due": pred.get("due_date")}
    return cats, flag


def strip_maildesk(categories):
    return [c.strip() for c in (categories or "").split(",") if c.strip() and not c.strip().startswith(PREFIX)]


def main():
    ap = argparse.ArgumentParser()
    g = ap.add_mutually_exclusive_group()
    g.add_argument("--apply", action="store_true")
    g.add_argument("--undo", action="store_true")
    ap.add_argument("--pred", default="cli-v1-medium")
    args = ap.parse_args()

    import win32com.client

    from outlook_check import find_test_folder

    preds = json.loads((RESULTS / f"{args.pred}_predictions.json").read_text(encoding="utf-8"))["predictions"]
    folder = find_test_folder(win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI"))
    if folder is None:
        print("'Maildesk-테스트' 폴더를 찾지 못했습니다.")
        sys.exit(1)

    mode = "적용" if args.apply else "되돌리기" if args.undo else "미리보기(변경 없음)"
    print(f"[{mode}] {folder.FolderPath}")
    changed = 0
    skipped_flags = []
    for item in folder.Items:
        p = item.UserProperties.Find("MDId")
        if getattr(item, "Class", None) != 43 or p is None or p.Value not in preds:
            continue
        mid = p.Value
        keep = strip_maildesk(item.Categories)
        if args.undo:
            if keep != [c.strip() for c in (item.Categories or "").split(",") if c.strip()] or item.IsMarkedAsTask:
                item.Categories = ", ".join(keep)
                if item.IsMarkedAsTask:
                    item.ClearTaskFlag()
                item.Save()
                changed += 1
            continue
        cats, flag = planned(preds[mid])
        flag_txt = f"깃발 '{flag['request']}' 기한 {flag['due'] or '없음'}" if flag else "깃발 없음"
        print(f"  {mid}: 범주 {cats} / {flag_txt}")
        if args.apply:
            item.Categories = ", ".join(keep + cats)
            if flag and not item.Sent:
                # 스크립트로 넣은 항목이 '미발송(초안)' 상태면 Outlook이 깃발을 허용하지 않음 → 태그만
                skipped_flags.append(mid)
            elif flag:
                item.MarkAsTask(OL_MARK_NO_DATE)
                item.FlagRequest = flag["request"]
                if flag["due"]:
                    # 정오로 넘겨야 시간대 변환(UTC↔KST)에도 날짜가 바뀌지 않음 (자정으로 넘기면 하루 당겨짐)
                    d = datetime.fromisoformat(flag["due"]).replace(hour=12)
                    item.TaskStartDate = d
                    item.TaskDueDate = d
            item.Save()
            changed += 1
    print(f"변경한 메일: {changed}건" if (args.apply or args.undo) else "미리보기만 했습니다. 적용하려면 --apply")
    if skipped_flags:
        print(f"깃발 생략 {len(skipped_flags)}건 (미발송 상태 항목이라 Outlook이 깃발을 허용하지 않음): {skipped_flags}")


if __name__ == "__main__":
    main()
