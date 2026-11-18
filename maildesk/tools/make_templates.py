"""측정용 엑셀 양식 2개를 만든다.

1) eval_kit/새테스트메일_양식.xlsx  — 사람이 직접 쓰는 새 테스트 메일 + 정답 (재검증용)
2) eval_kit/시간측정_기록지.xlsx    — 사람 단독 vs Maildesk 보조 처리 시간 비교 실험 기록

사용법: python maildesk/tools/make_templates.py
"""
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation

OUT = Path(__file__).resolve().parent.parent / "eval_kit"
HEAD = Font(bold=True, color="FFFFFF")
FILL_MAIL = PatternFill("solid", fgColor="2F5496")
FILL_ANS = PatternFill("solid", fgColor="7B2C2C")
WRAP = Alignment(wrap_text=True, vertical="top")


def header(ws, cols, fills):
    ws.append(cols)
    for c, f in zip(ws[1], fills):
        c.font, c.fill, c.alignment = HEAD, f, WRAP


def dropdown(ws, col, values, rows):
    dv = DataValidation(type="list", formula1='"' + ",".join(values) + '"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(f"{col}2:{col}{rows + 1}")


def testset():
    wb = Workbook()
    ws = wb.active
    ws.title = "메일 작성"
    mail_cols = ["ID", "받은 시각 (YYYY-MM-DD HH:MM)", "보낸 사람", "이메일(가상)", "회사(가상)", "제목", "본문"]
    ans_cols = ["정답: 1차 분류", "정답: 2차 분류", "정답: 회신 필요", "정답: 긴급도",
                "정답: 품번 × 수량 (줄바꿈으로 여러 개)", "정답: 기한 (YYYY-MM-DD)", "정답: 판단 근거 문장(본문 그대로)", "메모(함정 설명 등)"]
    header(ws, mail_cols + ans_cols, [FILL_MAIL] * len(mail_cols) + [FILL_ANS] * len(ans_cols))
    n = 20
    for k in range(1, n + 1):
        ws.append([f"T{k:02d}"])
    dropdown(ws, "H", ["업무", "참고·알림", "광고·기타"], n)
    dropdown(ws, "I", ["견적요청", "납기문의", "납품일정", "주문·변경", "일반문의", "-"], n)
    dropdown(ws, "J", ["Y", "N"], n)
    dropdown(ws, "K", ["상", "중", "하"], n)
    for col, w in zip("ABCDEFGHIJKLMNO", (6, 18, 12, 26, 16, 30, 60, 11, 11, 9, 8, 22, 14, 40, 30)):
        ws.column_dimensions[col].width = w
    for row in ws.iter_rows(min_row=2, max_row=n + 1):
        for c in row:
            c.alignment = WRAP
    ws.freeze_panes = "B2"

    g = wb.create_sheet("작성 안내", 0)
    lines = [
        ("기준 시각", "2026-11-25 10:00"),
        ("", "↑ 긴급도·상대 기한(내일, 다음 주 등)을 계산하는 기준. 바꿔도 됩니다. 받은 시각은 이보다 앞서야 합니다."),
        ("", ""),
        ("목적", "Maildesk가 처음 보는 메일로 다시 채점하기 위한 '시험지'입니다. 기존 30건은 AI가 만들어서 점수가 부풀려졌을 수 있습니다."),
        ("누가", "가능하면 Maildesk를 만들지 않은 사람(친구, 연구실 동료)이 쓰면 가장 좋습니다. 본인이 써도 됩니다."),
        ("주의", "실제 회사명·사람 이름·이메일·거래 정보는 쓰지 마세요. 모두 가상으로. 이메일은 ...@example.com 형태로."),
        ("분량", "15~20건. 업무 메일 2/3, 참고·알림과 광고 1/3 정도. 견적·납기·납품·주문변경·일반문의가 골고루."),
        ("팁", "실제 업무에서 헷갈렸던 상황을 넣어 주세요: 제목과 내용이 다른 메일, 수량 변경, 여러 요청이 섞인 메일, 영어 메일, 회신 스레드 등."),
        ("정답", "빨간 칸은 메일을 쓴 사람이 직접 정답을 적습니다. 판단 근거 문장은 본문에서 그대로 복사하세요."),
        ("", ""),
        ("1차 분류", "업무 = 실제 업무 요청/확인 / 참고·알림 = 읽고 알면 되는 메일(자동 알림, 사내 공지, 회신 불필요한 감사) / 광고·기타 = 홍보, 뉴스레터, 스팸·피싱"),
        ("2차 분류", "견적요청 / 납기문의(기존 주문의 출하·도착 시점) / 납품일정(입고·배송·장소 조율) / 주문·변경(신규 발주, 수량·사양 변경, 취소) / 일반문의 / 업무가 아니면 -"),
        ("회신 필요", "Y = 우리 쪽 회신·확인 답장이 필요 / N = 불필요(조건부 회신 포함)"),
        ("긴급도", "상 = 기한이 기준 시각부터 2영업일 이내, 또는 라인 정지·취소처럼 즉시 대응 / 중 = 회신·처리 필요하지만 상이 아님 / 하 = 회신·처리 불필요"),
        ("기한", "우리가 행동(회신·처리)해야 하는 날짜. 상대 표현은 받은 날짜 기준으로 계산. 없으면 비워 두기."),
        ("", ""),
        ("다 쓰면", "python maildesk/tools/import_testset.py eval_kit/새테스트메일_양식.xlsx  → data/sample_emails_v2.json 생성"),
    ]
    for a, b in lines:
        g.append([a, b])
    g["A1"].font = Font(bold=True)
    g["B1"].font = Font(bold=True, color="C00000")
    g.column_dimensions["A"].width = 12
    g.column_dimensions["B"].width = 120
    for row in g.iter_rows():
        for c in row:
            c.alignment = WRAP
    wb.save(OUT / "새테스트메일_양식.xlsx")


def timing():
    wb = Workbook()
    ws = wb.active
    ws.title = "시간 기록"
    cols = ["참가자", "메일 세트(A/B)", "조건", "메일 ID", "시작 (HH:MM:SS)", "끝 (HH:MM:SS)", "걸린 시간(초)",
            "1차 분류", "2차 분류", "회신 필요", "기한", "초안 수정 정도(보조 조건)", "메모"]
    header(ws, cols, [FILL_MAIL] * len(cols))
    for r in range(2, 202):
        ws.cell(r, 7, f'=IF(AND(E{r}<>"",F{r}<>""),ROUND((F{r}-E{r})*86400,0),"")')
        ws.cell(r, 5).number_format = "hh:mm:ss"
        ws.cell(r, 6).number_format = "hh:mm:ss"
    dropdown(ws, "C", ["사람 단독", "Maildesk 보조"], 200)
    dropdown(ws, "B", ["A", "B"], 200)
    dropdown(ws, "L", ["수정 없음", "조금(단어 몇 개)", "많이(문장 단위)", "새로 씀"], 200)
    for col, w in zip("ABCDEFGHIJKLM", (10, 10, 14, 8, 12, 12, 12, 10, 10, 9, 12, 18, 30)):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A2"

    s = wb.create_sheet("요약")
    s.append(["조건", "처리한 메일 수", "평균 시간(초)", "총 시간(분)"])
    for c in s[1]:
        c.font, c.fill = HEAD, FILL_MAIL
    for k, cond in enumerate(["사람 단독", "Maildesk 보조"], start=2):
        s.cell(k, 1, cond)
        s.cell(k, 2, f"=COUNTIFS('시간 기록'!C:C,A{k},'시간 기록'!G:G,\">0\")")
        s.cell(k, 3, f"=IFERROR(AVERAGEIFS('시간 기록'!G:G,'시간 기록'!C:C,A{k}),\"\")")
        s.cell(k, 4, f"=IFERROR(SUMIFS('시간 기록'!G:G,'시간 기록'!C:C,A{k})/60,\"\")")
    s.append([])
    s.append(["시간 절감률", "", "=IFERROR(1-C3/C2,\"\")"])
    s["C5"].number_format = "0%"
    s.append(["주의", "정확도·놓친 메일 수는 정답표와 비교해 따로 계산합니다. 숫자는 '참가자 n명, 메일 m건 기준'과 함께 적습니다."])
    for col, w in zip("ABCD", (16, 16, 16, 16)):
        s.column_dimensions[col].width = w
    wb.save(OUT / "시간측정_기록지.xlsx")


if __name__ == "__main__":
    OUT.mkdir(exist_ok=True)
    testset()
    timing()
    print("저장:", OUT / "새테스트메일_양식.xlsx", "/", OUT / "시간측정_기록지.xlsx")
