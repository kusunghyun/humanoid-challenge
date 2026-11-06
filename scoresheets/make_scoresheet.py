"""Generate the TASK A / TASK B scoring sheets (채점표) as .xlsx files, in the TASK A layout.

Usage: python3 scoresheets/make_scoresheet.py
  -> scoresheets/scoresheet.xlsx (TASK A + TASK B tabs)

Layout (TASK A 양식):
  1  TASK X 채점표
  2  통과 규칙 한 줄
  3  팀이름: / 팀 대표 / 채점자 / 촬영자 / 안전관리자
  4  채점 항목 | 채점 항목 | 시도 # ... (배점, 체크 두 칸씩)
  5  배점 / 체크
  .. 항목 (대분류 칸은 세로 병합)
  .. 합계 행 (배점 합계, 체크=O 인 배점 합계)
  .. 추가사항.
"""
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

TASK_A = {
    "title": "TASK A 채점표",
    "rule": "A-(n)을 모두 통과하지 못하면 A-(n+1,+2,+3)를 획득하지 못한다. (A-4 는 A-2 까지 통과하면 획득할 수 있다.)",
    "attempts": 5,
    "items": [
        ("A-1 집기", "그리퍼가 한 쪽(Left or Right)이라도 바구니에 한 번이라도 닿았다", 1),
        ("A-1 집기", "바구니를 쥔 채로, 바구니 밑면과 탁자 윗면 사이에 틈이 보였다", 2),
        ("A-1 집기", "로봇이 든 바구니가, 한 번이라도 탁자 전체에서 벗어났다.", 3),
        ("A-2 옮기기", "로봇 바퀴나 베이스 일부가 바닥의 도착 구역 선(책상 근처 30cm 반경 다각형) 한 바퀴라도 안쪽으로 들어왔다", 4),
        ("A-2 옮기기", "바로 위 항목을 받았을 당시에 로봇이 바구니를 들고 있었다.", 2),
        ("A-3 놓기", "판 내내 책상 다리가 바닥 표시선에서 5cm 이상 밀리지 않았다. (크게 밀린 경우, A-3 전체를 받지 못함)", 1),
        ("A-3 놓기", "바구니 밑면이 책상 상판에 한번이라도 닿았다 ", 2),
        ("A-3 놓기", "그리퍼를 열고 손을 뗀 뒤 6초 동안, 바구니가 상판에서 떨어지거나 넘어지지 않았다 (떨어졌을 경우, A-3 모든 점수를 받지 못한다)", 4),
        ("A-4 부가점수", "판 내내 로봇과 바구니가 매장 가구(진열대·냉장고·계산대·스툴)에 닿지 않았다 (A-2 를 모두 통과한 시도에만)", 4),
    ],
    "notes": [
        "가구에 살짝 스쳤지만 가구가 움직이지 않았으면 A-4 의 4점만 잃고 시도는 계속한다.",
        "6초는 마지막으로 손을 뗀 때부터 센다. 다시 쥐고 고쳐 놓으면 마지막 놓기부터 다시 6초. 1초 미만의 순간 접촉은 다시 쥔 것으로 보지 않는다. 손을 끝까지 떼지 않으면 6초 항목은 0점.",
        "바닥 표시: 도착 구역 = 책상 상판 중심 기준 반지름 1.0 m 원.       ***책상 밀림 허용 = 다리 표시선에서 5 cm. ",
        "6번 시도 후에 좌석 바꾸기",
    ],
}

# TASK B: 원래 채점표 내용 그대로 (시도 6회, 시도 1만 일부 배점이 다름), TASK A 양식으로.
TASK_B = {
    "title": "TASK B 채점표",
    "rule": "B-(n)을 모두 통과하지 못하면 B-(n+1,+2,+3)를 획득하지 못한다.",
    "attempts": 6,
    "problem_no": True,  # 추가사항에 '문제 번호 기입'이 있어 칸을 둠
    "items": [
        ("B-1 집기", "그리퍼가 상품에 한 번이라도 닿았다", 1),
        ("B-1 집기", "상품을 쥔 채로, 조금이라도 들어올렸다.", 2),
        ("B-1 집기", "상품을 쥔 채로, 상품 전체가 (위에서 볼 때) 검은 탁자 가장자리 바깥으로 나왔다", 3),
        ("B-2 옮기기", "상품을 쥔 채로, 상품 일부가 바닥의 흰 테두리 선 안쪽으로 들어왔다", 2),
        ("B-2 옮기기", "상품을 쥔 채로, 상품 일부가 목표 판 윗면보다 위로 올라왔다", 2),
        ("B-3 놓기", "목표 층 판 위에 놓였다", 2),
        ("B-3 놓기", "바닥에 닿은 상품 면이 기준 사각형과 겹친다", [3, 2, 2, 2, 2, 2]),
        ("B-3 놓기", "라벨 기준 상품 상하가 뒤집히지 않았다", [3, 1, 1, 1, 1, 1]),
        ("B-3 놓기", "정면에서 상품 앞면이 조금이라도 보인다", 4),
        ("B-3 놓기", "놓은 뒤 상품이 스스로 구르거나, 미끄러지거나, 넘어지지 않았다", [3, 2, 2, 2, 2, 2]),
        ("B-3 놓기", "다른 상품들이 모두 서 있고, 진열 위치 표시에서 밀린 틈이 보이지 않는다", 3),
        ("B-4 부가점수", "상품이 바닥에 떨어지지 않고 선반 판 위에 놓였다", 2),
        ("B-4 부가점수", "상품을 끝까지 떨어뜨리지 않았는가", 6),
    ],
    "notes": [
        "항상 놓기 항목 관련 평가를 할 경우, 뒤에 있는 상품을 제거하고 준비된 주황도구로 채점할 것.",
        "항상 문제 번호를 제대로 기입할 것.",
        "시도 6번이 끝나면 상품을 바꿀 것.",
        "문제지 3번을 모두 풀고 나면, 작업 공간 층의 상품들의 배열을 바꿀 것.",
    ],
}

THIN = Side(style="thin")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
HEADER_FILL = PatternFill("solid", fgColor="D9D9D9")
BOLD = Font(bold=True)


def build_sheet(wb, task):
    n = task["attempts"]
    last_col = 2 + 2 * n
    ws = wb.create_sheet(task["title"].replace(" 채점표", ""))

    # 1: title, 2: rule
    ws["A1"] = task["title"]
    ws["A1"].font = Font(bold=True, size=14)
    if task.get("problem_no"):
        ws.cell(row=1, column=8, value="문제번호:").font = BOLD
    ws["A2"] = task["rule"]
    ws["A2"].font = BOLD
    ws["A2"].alignment = LEFT
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=last_col)
    ws.row_dimensions[2].height = 30

    # 3: team / roles
    for col, label in ((1, "팀이름:"), (3, "팀 대표"), (5, "채점자"), (7, "촬영자"), (9, "안전관리자")):
        ws.cell(row=3, column=col, value=label).font = BOLD

    # 4-5: header
    ws.cell(row=4, column=1, value="채점 항목")
    ws.cell(row=4, column=2, value="채점 항목")
    ws.merge_cells("A4:B5")
    for i in range(n):
        col = 3 + 2 * i
        ws.cell(row=4, column=col, value="시도 #")
        ws.merge_cells(start_row=4, start_column=col, end_row=4, end_column=col + 1)
        ws.cell(row=5, column=col, value="배점")
        ws.cell(row=5, column=col + 1, value="체크")

    # items
    first = 6
    items = task["items"]
    dv = DataValidation(type="list", formula1='"O,X"', allow_blank=True)
    ws.add_data_validation(dv)
    for r, (_, text, pts) in enumerate(items, start=first):
        pts = pts if isinstance(pts, list) else [pts] * n
        ws.cell(row=r, column=2, value=text).alignment = LEFT
        for i in range(n):
            col = 3 + 2 * i
            ws.cell(row=r, column=col, value=pts[i])
            dv.add(ws.cell(row=r, column=col + 1))
        ws.row_dimensions[r].height = 45 if len(text) > 40 else 32
    last = first + len(items) - 1

    # merge category cells
    start = 0
    for idx in range(1, len(items) + 1):
        if idx == len(items) or items[idx][0] != items[start][0]:
            ws.cell(row=first + start, column=1, value=items[start][0])
            if idx - 1 > start:
                ws.merge_cells(start_row=first + start, start_column=1, end_row=first + idx - 1, end_column=1)
            start = idx

    # 합계 row
    total = last + 1
    ws.cell(row=total, column=2, value="합계")
    for i in range(n):
        col = 3 + 2 * i
        p, c = get_column_letter(col), get_column_letter(col + 1)
        ws.cell(row=total, column=col, value=f"=SUM({p}{first}:{p}{last})")
        ws.cell(row=total, column=col + 1, value=f'=SUMIF({c}{first}:{c}{last},"O",{p}{first}:{p}{last})')

    # table styling
    for row in ws.iter_rows(min_row=4, max_row=total, min_col=1, max_col=last_col):
        for cell in row:
            cell.border = BORDER
            if cell.column != 2 or cell.row <= 5 or cell.row == total:
                cell.alignment = CENTER
    for row in ws.iter_rows(min_row=4, max_row=5, min_col=1, max_col=last_col):
        for cell in row:
            cell.font = BOLD
            cell.fill = HEADER_FILL
    for cell in ws[total][:last_col]:
        cell.font = BOLD
        cell.fill = HEADER_FILL

    # 추가사항
    nr = total + 2
    ws.cell(row=nr, column=1, value="추가사항.").font = BOLD
    for k, note in enumerate(task["notes"], start=nr + 1):
        ws.cell(row=k, column=1, value=note).alignment = LEFT
        ws.merge_cells(start_row=k, start_column=1, end_row=k, end_column=last_col)
        ws.row_dimensions[k].height = 30 if len(note) > 70 else 18

    # sizing / print
    ws.column_dimensions["A"].width = 13
    ws.column_dimensions["B"].width = 60
    for col in range(3, last_col + 1):
        ws.column_dimensions[get_column_letter(col)].width = 6
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True


def main():
    wb = Workbook()
    wb.remove(wb.active)
    build_sheet(wb, TASK_A)
    build_sheet(wb, TASK_B)
    out = Path(__file__).with_name("scoresheet.xlsx")
    wb.save(out)
    print(f"wrote {out}")


if __name__ == "__main__":
    main()
