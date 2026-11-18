"""Maildesk 1단계: 클래식 Outlook 연결 확인 (읽기 전용)

- 'Maildesk-테스트' 폴더 하나만 연다. 받은편지함 등 다른 폴더의 메일은 읽지 않는다.
- 메일을 보내거나, 지우거나, 옮기는 코드는 이 파일에 없다.
- 폴더 안 메일의 수신시각 / 보낸 사람 / 제목만 화면에 보여준다.

사용법 (클래식 Outlook을 먼저 켜 둔 상태에서):
    python maildesk/connect/outlook_check.py
"""
import sys

import win32com.client

TEST_FOLDER = "Maildesk-테스트"  # 이 이름의 폴더만 읽는다
MAX_SHOW = 50


def find_test_folder(ns):
    """Outlook에 연결된 각 계정의 최상위 폴더 목록에서 테스트 폴더를 이름으로 찾는다.
    (계정이 여러 개라도) 다른 폴더는 '이름'만 비교하고, 그 안의 메일은 열지 않는다."""
    found = []
    for store_root in ns.Folders:  # 계정(메일함)별 최상위
        for f in store_root.Folders:
            if f.Name == TEST_FOLDER:
                found.append(f)
    if len(found) > 1:
        print(f"주의: '{TEST_FOLDER}' 폴더가 {len(found)}개 있습니다. 첫 번째를 사용합니다.")
    return found[0] if found else None


def main():
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
    except Exception as e:
        print("클래식 Outlook에 연결하지 못했습니다. 클래식 Outlook이 켜져 있는지 확인하세요.")
        print("상세:", e)
        sys.exit(1)

    ns = outlook.GetNamespace("MAPI")
    folder = find_test_folder(ns)
    if folder is None:
        print(f"'{TEST_FOLDER}' 폴더를 찾지 못했습니다.")
        print("클래식 Outlook에서 계정 바로 아래(받은 편지함과 같은 줄)에 이 이름으로 폴더를 만들어 주세요.")
        sys.exit(1)

    items = folder.Items
    items.Sort("[ReceivedTime]", True)
    print(f"연결 성공: '{folder.FolderPath}' 폴더, 항목 {items.Count}개")
    for i, item in enumerate(items):
        if i >= MAX_SHOW:
            break
        if getattr(item, "Class", None) != 43:  # 43 = 일반 메일
            continue
        print(f"- {item.ReceivedTime:%Y-%m-%d %H:%M} | {item.SenderName} | {item.Subject}")


if __name__ == "__main__":
    main()
