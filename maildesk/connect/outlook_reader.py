"""Maildesk 메일 읽기 모듈 (읽기 전용)

'Maildesk-테스트' 폴더의 메일을 읽어 분류기가 쓰기 좋은 형태(dict 목록)로 돌려준다.
- Outlook이 저장하며 바꾼 형식(줄바꿈 \\r\\n, 줄 끝 공백)을 정리한다.
- 본문 첫 줄의 '[가상 메일 ...]' 안내 줄은 떼어낸다.
- 가상 보낸 사람·시나리오 수신 시각은 숨김 속성(MDId 등)에서 가져온다.
- 졸업 후 학교 계정이 없어져도 데모할 수 있도록, JSON 파일에서 같은 형태로 읽는 함수도 둔다.
"""
import json
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SAMPLE_JSON = ROOT / "data" / "sample_emails_v1.json"
HEADER_PREFIX = "[가상 메일 "


def normalize_body(text):
    lines = text.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    lines = [ln.rstrip() for ln in lines]
    if lines and lines[0].startswith(HEADER_PREFIX):
        lines = lines[1:]
    return "\n".join(lines).strip("\n")


def _prop(item, name):
    p = item.UserProperties.Find(name)
    return p.Value if p is not None else None


def read_from_outlook():
    import win32com.client

    from outlook_check import find_test_folder

    ns = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    folder = find_test_folder(ns)
    if folder is None:
        raise RuntimeError("'Maildesk-테스트' 폴더를 찾지 못했습니다.")
    mails = []
    for item in folder.Items:
        if getattr(item, "Class", None) != 43:  # 일반 메일만
            continue
        mails.append({
            "id": _prop(item, "MDId") or item.EntryID[-8:],
            "received_at": _prop(item, "MDReceivedAt") or item.ReceivedTime.strftime("%Y-%m-%dT%H:%M"),
            "from_name": _prop(item, "MDFromName") or item.SenderName,
            "from_email": _prop(item, "MDFromEmail") or "",
            "company": _prop(item, "MDCompany") or "",
            "subject": item.Subject,
            "body": normalize_body(item.Body),
            "source": "outlook",
        })
    return sorted(mails, key=lambda m: m["id"])


def read_from_json(path=SAMPLE_JSON):
    emails = json.loads(Path(path).read_text(encoding="utf-8"))["emails"]
    return [{k: m[k] for k in ("id", "received_at", "from_name", "from_email", "company", "subject", "body")}
            | {"source": "json"} for m in emails]


if __name__ == "__main__":
    # 확인용: Outlook에서 읽은 결과가 원본 JSON과 같은지 대조
    out = {m["id"]: m for m in read_from_outlook()}
    src = {m["id"]: m for m in read_from_json()}
    diff = [i for i in src if i not in out or any(out[i][k] != src[i][k] for k in
            ("received_at", "from_name", "from_email", "company", "subject", "body"))]
    print(f"Outlook {len(out)}건 / 원본 {len(src)}건 / 불일치: {diff if diff else '없음'}")
