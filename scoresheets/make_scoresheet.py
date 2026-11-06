"""Generate the TASK A / TASK B scoring sheets (채점표) as one .xlsx file.

Usage: python3 scoresheets/make_scoresheet.py  ->  scoresheets/scoresheet.xlsx

Each task is described by a list of (category, item, [points for attempts 1..6]).
TASK_A_ITEMS is filled in once the Task A rules are confirmed.
"""
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

ATTEMPTS = 6

TASK_B_ITEMS = [
    ("집기", "그리퍼가 상품에 한 번이라도 닿았다", [1] * 6),
    ("집기", "상품을 쥔 채로, 조금이라도 들어올렸다.", [2] * 6),
    ("집기", "상품을 쥔 채로, 상품 전체가 (위에서 볼 때) \n검은 탁자 가장자리 바깥으로 나왔다", [3] * 6),
    ("옮기기", "상품을 쥔 채로, 상품 일부가 바닥의 흰 테두리 선 안쪽으로 들어왔다", [2] * 6),
    ("옮기기", "상품을 쥔 채로, 상품 일부가 목표 판 윗면보다 위로 올라왔다", [2] * 6),
    ("놓기", "목표 층 판 위에 놓였다", [2] * 6),
    ("놓기", "바닥에 닿은 상품 면이 기준 사각형과 겹친다", [3, 2, 2, 2, 2, 2]),
    ("놓기", "라벨 기준 상품 상하가 뒤집히지 않았다", [3, 1, 1, 1, 1, 1]),
    ("놓기", "정면에서 상품 앞면이 조금이라도 보인다", [4] * 6),
    ("놓기", "놓은 뒤 상품이 스스로 구르거나, 미끄러지거나, 넘어지지 않았다", [3, 2, 2, 2, 2, 2]),
    ("놓기", "다른 상품들이 모두 서 있고, \n진열 위치 표시에서 밀린 틈이 보이지 않는다", [3] * 6),
    ("부가점수", "상품이 바닥에 떨어지지 않고 선반 판 위에 놓였다", [2] * 6),
    ("부가점수", "상품을 끝까지 떨어뜨리지 않았는가", [6] * 6),
]

TASK_B_NOTES = (
    "- B-(n)을 모두 통과하지 못하면 B-(n+1,+2,+3)를 획득하지 못한다.   \n"
    "- 항상 놓기 항목 관련 평가를 할 경우, 뒤에 있는 상품을 제거하고 준비된 주황도구로 채점할것\n",
    "- 항상 문제 번호를 제대로 기입할 것\n"
    "- 시도 6번이 끝나면 상품을 바꿀 것\n"
    "- 문제지 3번을 모두 풀고 나면, 작업 공간 층의 상품들의 배열을 바꿀것",
)

# TODO: Task A 채점 항목 — fill in (category, item, points x6) once confirmed.
TASK_A_ITEMS = [
    ("집기", "", [None] * 6),
    ("집기", "", [None] * 6),
    ("집기", "", [None] * 6),
    ("옮기기", "", [None] * 6),
    ("옮기기", "", [None] * 6),
    ("놓기", "", [None] * 6),
    ("놓기", "", [None] * 6),
    ("놓기", "", [None] * 6),
    ("놓기", "", [None] * 6),
    ("놓기", "", [None] * 6),
    ("놓기", "", [None] * 6),
    ("부가점수", "", [None] * 6),
    ("부가점수", "", [None] * 6),
]

TASK_A_NOTES = (
    "- A-(n)을 모두 통과하지 못하면 A-(n+1,+2,+3)를 획득하지 못한다.   \n",
    "- 항상 문제 번호를 제대로 기입할 것\n"
    "- 시도 6번이 끝나면 상품을 바꿀 것",
)

THIN = Side(style="thin")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
TOP_LEFT = Alignment(horizontal="left", vertical="top", wrap_text=True)
HEADER_FILL = PatternFill("solid", fgColor="D9D9D9")
BOLD = Font(bold=True)


def build_sheet(wb, task, items, notes):
    ws = wb.create_sheet(f"TASK {task}")
    last_col = 2 + 2 * ATTEMPTS  # A, B + (배점, 체크) per attempt = 14 (N)

    # Row 1: title + 문제번호
    ws["A1"] = f"TASK {task} 채점표"
    ws["A1"].font = Font(bold=True, size=14)
    ws.merge_cells("A1:F1")
    ws["H1"] = "문제번호:"
    ws["H1"].font = BOLD
    ws.merge_cells(start_row=1, start_column=9, end_row=1, end_column=last_col)

    # Row 2: team / roles (label, then blank cell to write in)
    for col, label in ((1, "팀이름:"), (4, "팀 대표"), (6, "채점자"), (8, "촬영자"), (10, "안전관리자")):
        c = ws.cell(row=2, column=col, value=label)
        c.font = BOLD
    ws.merge_cells("K2:N2")

    # Rows 3-4: header
    ws.cell(row=3, column=1, value="채점 항목")
    ws.cell(row=3, column=2, value="채점 항목")
    ws.merge_cells("A3:A4")
    ws.merge_cells("B3:B4")
    for i in range(ATTEMPTS):
        col = 3 + 2 * i
        ws.cell(row=3, column=col, value=f"시도 # {i + 1}")
        ws.merge_cells(start_row=3, start_column=col, end_row=3, end_column=col + 1)
        ws.cell(row=4, column=col, value="배점")
        ws.cell(row=4, column=col + 1, value="체크 ")

    # Item rows
    first = 5
    check_dv = DataValidation(type="list", formula1='"O,X"', allow_blank=True)
    ws.add_data_validation(check_dv)
    for r, (_, item, pts) in enumerate(items, start=first):
        ws.cell(row=r, column=2, value=item or None).alignment = LEFT
        for i, p in enumerate(pts):
            col = 3 + 2 * i
            ws.cell(row=r, column=col, value=p)
            check_dv.add(ws.cell(row=r, column=col + 1))
    last = first + len(items) - 1

    # Merge category cells in column A
    start = first
    for idx in range(1, len(items) + 1):
        if idx == len(items) or items[idx][0] != items[start - first][0]:
            end = first + idx - 1
            ws.cell(row=start, column=1, value=items[start - first][0])
            if end > start:
                ws.merge_cells(start_row=start, start_column=1, end_row=end, end_column=1)
            start = end + 1

    # 총합 row: "합계" label in 배점 column, sum of checked (O) points in 체크 column
    total = last + 1
    ws.cell(row=total, column=1, value="총합")
    ws.merge_cells(start_row=total, start_column=1, end_row=total, end_column=2)
    for i in range(ATTEMPTS):
        col = 3 + 2 * i
        pts_l, chk_l = get_column_letter(col), get_column_letter(col + 1)
        ws.cell(row=total, column=col, value="합계")
        ws.cell(
            row=total,
            column=col + 1,
            value=f'=SUMIF({chk_l}{first}:{chk_l}{last},"O",{pts_l}{first}:{pts_l}{last})',
        )

    # Styling for the table
    for row in ws.iter_rows(min_row=3, max_row=total, min_col=1, max_col=last_col):
        for c in row:
            c.border = BORDER
            if c.column != 2 or c.row <= 4:
                c.alignment = CENTER
    for row in ws.iter_rows(min_row=3, max_row=4, min_col=1, max_col=last_col):
        for c in row:
            c.font = BOLD
            c.fill = HEADER_FILL
    for c in ws[total]:
        c.font = BOLD

    # Notes block (2 rows below the total)
    nr = total + 2
    ws.cell(row=nr, column=1, value="채점관 유의 사항").font = BOLD
    ws.merge_cells(start_row=nr, start_column=1, end_row=nr, end_column=10)
    ws.cell(row=nr, column=11, value="메모)\n\n").alignment = TOP_LEFT
    ws.merge_cells(start_row=nr, start_column=11, end_row=nr + 2, end_column=last_col)
    ws.cell(row=nr + 2, column=1, value=notes[0]).alignment = TOP_LEFT
    ws.merge_cells(start_row=nr + 2, start_column=1, end_row=nr + 2, end_column=5)
    ws.cell(row=nr + 2, column=6, value=notes[1]).alignment = TOP_LEFT
    ws.merge_cells(start_row=nr + 2, start_column=6, end_row=nr + 2, end_column=10)
    for row in ws.iter_rows(min_row=nr, max_row=nr + 2, min_col=1, max_col=last_col):
        for c in row:
            c.border = BORDER
    ws.row_dimensions[nr + 2].height = 75

    # Sizing / print setup
    ws.column_dimensions["A"].width = 10
    ws.column_dimensions["B"].width = 52
    for col in range(3, last_col + 1):
        ws.column_dimensions[get_column_letter(col)].width = 6
    for r in range(first, last + 1):
        ws.row_dimensions[r].height = 32
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    return ws


def main():
    wb = Workbook()
    wb.remove(wb.active)
    build_sheet(wb, "A", TASK_A_ITEMS, TASK_A_NOTES)
    build_sheet(wb, "B", TASK_B_ITEMS, TASK_B_NOTES)
    out = Path(__file__).with_name("scoresheet.xlsx")
    wb.save(out)
    print(f"wrote {out}")


if __name__ == "__main__":
    main()
