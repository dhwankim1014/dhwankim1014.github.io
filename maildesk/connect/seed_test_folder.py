"""가상 샘플 메일을 'Maildesk-테스트' 폴더에 넣는 테스트 데이터 도구

- 메일을 '보내지' 않는다. 테스트 폴더 안에 메일 모양의 항목을 '만들기'만 한다.
- 테스트 폴더 외의 폴더에는 아무것도 하지 않는다. 삭제 기능 없음.
- 각 항목에 숨김 속성(MDId 등)을 붙여 정답표와 짝지을 수 있게 한다.
  (Outlook 속성 이름에는 [ ] _ # 를 쓸 수 없음)
- 같은 MDId가 이미 있으면 건너뛴다(중복 방지).

사용법:
    python maildesk/connect/seed_test_folder.py --only M01        # 1건만
    python maildesk/connect/seed_test_folder.py --only M02 M03    # 여러 건
    python maildesk/connect/seed_test_folder.py --all             # 30건 전부
"""
import argparse
import json
import sys
from datetime import datetime, timedelta, timezone
from pathlib import Path

import win32com.client

from outlook_check import find_test_folder

ROOT = Path(__file__).resolve().parent.parent
SRC = ROOT / "data" / "sample_emails_v1.json"

OL_POST_ITEM = 6
PR_MESSAGE_FLAGS = "http://schemas.microsoft.com/mapi/proptag/0x0E070003"
PR_MESSAGE_DELIVERY_TIME = "http://schemas.microsoft.com/mapi/proptag/0x0E060040"
PR_CLIENT_SUBMIT_TIME = "http://schemas.microsoft.com/mapi/proptag/0x00390040"
KST = timezone(timedelta(hours=9))
MSGFLAG_READ = 1  # 미발송(8) 비트 없이 저장 → '받은 메일'로 인식. 이후 UnRead=True로 안 읽음 표시
OL_TEXT = 1
HEADER_PREFIX = "[가상 메일 "  # Maildesk가 본문을 읽을 때 이 줄은 떼어낸다


def existing_ids(folder):
    ids = set()
    for item in folder.Items:
        try:
            p = item.UserProperties.Find("MDId")
            if p is not None:
                ids.add(p.Value)
        except Exception:
            pass
    return ids


def make_item(folder, m):
    item = folder.Items.Add(OL_POST_ITEM)
    # 메시지 플래그를 '받은 메일'로 지정 — 반드시 MessageClass를 바꾸기 전, 첫 저장 전에 해야 먹힌다.
    # (이게 없으면 Outlook이 '미발송 초안'으로 보고 기한 깃발을 거부함. 2026-11-19 학교 Exchange 계정에서 확인)
    pa = item.PropertyAccessor
    pa.SetProperty(PR_MESSAGE_FLAGS, MSGFLAG_READ)
    # Outlook 목록의 '받은 날짜'를 시나리오 수신 시각으로 (MAPI 시각은 UTC로 지정)
    kst = datetime.fromisoformat(m["received_at"]).replace(tzinfo=KST)
    # pywin32는 시간대 정보가 있으면 현지 시각으로 한 번 더 바꿔 버리므로, UTC 숫자를 '시간대 없이' 넘긴다
    utc = kst.astimezone(timezone.utc).replace(tzinfo=None)
    for prop in (PR_MESSAGE_DELIVERY_TIME, PR_CLIENT_SUBMIT_TIME):
        try:
            pa.SetProperty(prop, utc)
        except Exception:  # noqa: BLE001
            pass
    item.Subject = m["subject"]
    header = (
        f"{HEADER_PREFIX}{m['id']} | 보낸 사람: {m['from_name']} <{m['from_email']}> | "
        f"{m['company']} | 시나리오 수신: {m['received_at'].replace('T', ' ')}]"
    )
    item.Body = header + "\n\n" + m["body"]
    for name, value in (
        ("MDId", m["id"]),
        ("MDFromName", m["from_name"]),
        ("MDFromEmail", m["from_email"]),
        ("MDCompany", m["company"]),
        ("MDReceivedAt", m["received_at"]),
    ):
        item.UserProperties.Add(name, OL_TEXT).Value = value
    item.MessageClass = "IPM.Note"  # 게시물 대신 일반 메일처럼 보이게
    item.Save()
    # 안 읽음 표시는 첫 저장 '뒤에' — 저장 전에 하면 Outlook이 미발송 플래그를 되살림
    item.UnRead = True
    item.Save()


def main():
    ap = argparse.ArgumentParser()
    g = ap.add_mutually_exclusive_group(required=True)
    g.add_argument("--only", nargs="+", metavar="ID")
    g.add_argument("--all", action="store_true")
    g.add_argument("--clear", action="store_true",
                   help="테스트 폴더에서 MDId가 붙은 가상 메일만 지움(지운 편지함으로 이동)")
    ap.add_argument("--data", type=Path, default=SRC, help="넣을 메일 JSON (기본: sample_emails_v1.json)")
    args = ap.parse_args()

    emails = json.loads(args.data.read_text(encoding="utf-8"))["emails"]
    if args.only:
        wanted = set(args.only)
        unknown = wanted - {m["id"] for m in emails}
        if unknown:
            print("없는 ID:", ", ".join(sorted(unknown)))
            sys.exit(1)
        emails = [m for m in emails if m["id"] in wanted]

    ns = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    folder = find_test_folder(ns)
    if folder is None:
        print("'Maildesk-테스트' 폴더를 찾지 못했습니다.")
        sys.exit(1)

    if args.clear:
        items = folder.Items
        removed = 0
        for k in range(items.Count, 0, -1):  # 지우면 번호가 당겨지므로 뒤에서부터
            item = items.Item(k)
            try:
                p = item.UserProperties.Find("MDId")
            except Exception:  # noqa: BLE001
                p = None
            if p is not None:
                item.Delete()  # 영구 삭제가 아니라 '지운 편지함'으로 이동
                removed += 1
        print(f"지움: {removed}건 (지운 편지함으로 이동) / 남은 항목 {folder.Items.Count}개")
        return

    have = existing_ids(folder)
    added = 0
    for m in emails:
        if m["id"] in have:
            print(f"건너뜀(이미 있음): {m['id']}")
            continue
        make_item(folder, m)
        added += 1
        print(f"추가: {m['id']} {m['subject']}")
    print(f"완료: {added}건 추가 / 폴더 항목 {folder.Items.Count}개 ({folder.FolderPath})")


if __name__ == "__main__":
    main()
