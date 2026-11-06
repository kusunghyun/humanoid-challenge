"""TASK A 채점표 (.xlsx) in the TASK B Google Sheets layout.

Usage: python3 scoresheets/make_scoresheet.py  ->  scoresheets/task_a_scoresheet.xlsx

Sizes are measured from the TASK B sheet in Google Sheets (pixels at 100% zoom):
  columns  A=116, B=481, C..N=100
  rows     1=89, 2=32, 3=37, 4=21, item rows=40, 총합 2x21, 유의 사항 2x21
Fonts/fills follow the TASK B sheet (Arial; title row #D9D9D9, header rows #EFEFEF).
"""
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

TITLE = "TASK A 채점표"
ATTEMPTS = 6

# (대분류, 채점 항목, 배점) — 시도 1~6 같은 배점
ITEMS = [
    ("A-1 집기", "그리퍼가 한 쪽(Left or Right)이라도 바구니에 한 번이라도 닿았다", 1),
    ("A-1 집기", "바구니를 쥔 채로, 바구니 밑면과 탁자 윗면 사이에 틈이 보였다", 2),
    ("A-1 집기", "로봇이 든 바구니가, 한 번이라도 탁자 전체에서 벗어났다.", 3),
    ("A-2 옮기기", "로봇 바퀴나 베이스 일부가 바닥의 도착 구역 선(책상 근처 30cm 반경 다각형) 한 바퀴라도 안쪽으로 들어왔다", 4),
    ("A-2 옮기기", "바로 위 항목을 받았을 당시에 로봇이 바구니를 들고 있었다.", 2),
    ("A-3 놓기", "판 내내 책상 다리가 바닥 표시선에서 5cm 이상 밀리지 않았다. (크게 밀린 경우, A-3 전체를 받지 못함)", 1),
    ("A-3 놓기", "바구니 밑면이 책상 상판에 한번이라도 닿았다", 2),
    ("A-3 놓기", "그리퍼를 열고 손을 뗀 뒤 6초 동안, 바구니가 상판에서 떨어지거나 넘어지지 않았다 (떨어졌을 경우, A-3 모든 점수를 받지 못한다)", 4),
    ("A-4 부가점수", "판 내내 로봇과 바구니가 매장 가구(진열대·냉장고·계산대·스툴)에 닿지 않았다 (A-2 를 모두 통과한 시도에만)", 4),
]

NOTES_LEFT = (
    "- A-(n)을 모두 통과하지 못하면 A-(n+1,+2,+3)를 획득하지 못한다. (A-4 는 A-2 까지 통과하면 획득할 수 있다.)\n"
    "- 가구에 살짝 스쳤지만 가구가 움직이지 않았으면 A-4 의 4점만 잃고 시도는 계속한다.\n"
    "- 6초는 마지막으로 손을 뗀 때부터 센다. 다시 쥐고 고쳐 놓으면 마지막 놓기부터 다시 6초. "
    "1초 미만의 순간 접촉은 다시 쥔 것으로 보지 않는다. 손을 끝까지 떼지 않으면 6초 항목은 0점."
)
NOTES_RIGHT = (
    "- 바닥 표시: 도착 구역 = 책상 상판 중심 기준 반지름 1.0 m 원.\n"
    "- 책상 밀림 허용 = 다리 표시선에서 5 cm.\n"
    "- 6번 시도 후에 좌석 바꾸기"
)

# sizes in px (from TASK B)
COL_PX = {"A": 116, "B": 481, **{get_column_letter(c): 100 for c in range(3, 15)}}
ROW_PX = {1: 89, 2: 32, 3: 37, 4: 21}
ITEM_ROW_PX = 40
SMALL_ROW_PX = 21
NOTES_ROW_PX = 110  # TASK B: 67px for 3 lines; TASK A notes run to ~5 lines

GREY = PatternFill("solid", fgColor="D9D9D9")
LIGHT = PatternFill("solid", fgColor="EFEFEF")
THIN = Side(style="thin", color="000000")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def px_to_width(px):
    return round((px - 5) / 7, 2)


def px_to_pt(px):
    return px * 0.75


def font(sz, b=False, i=False):
    return Font(name="Arial", size=sz, bold=b, italic=i)


def put(ws, ref, value, f, fill=None, h="general", v="bottom"):
    c = ws[ref]
    c.value = value
    c.font = f
    c.alignment = Alignment(horizontal=h, vertical=v, wrap_text=True)
    if fill:
        c.fill = fill


def box(ws, ref, fill=None):
    for row in ws[ref]:
        for c in row:
            c.border = BOX
            if fill:
                c.fill = fill


def main():
    wb = Workbook()
    ws = wb.active
    ws.title = "TASK A"

    for col, px in COL_PX.items():
        ws.column_dimensions[col].width = px_to_width(px)
    for r, px in ROW_PX.items():
        ws.row_dimensions[r].height = px_to_pt(px)

    # 1: title / 문제번호
    box(ws, "A1:N1", GREY)
    put(ws, "A1", TITLE, font(29, b=True), GREY)
    put(ws, "G1", "문제번호:", font(29, b=True), GREY)
    ws.merge_cells("A1:F1")
    ws.merge_cells("G1:N1")

    # 2: team / roles
    box(ws, "A2:N2", LIGHT)
    for ref, text in (("A2", "팀이름:"), ("C2", "팀 대표"), ("E2", "채점자"), ("G2", "촬영자"), ("I2", "안전관리자")):
        put(ws, ref, text, font(14, b=True, i=True), LIGHT)
    ws.merge_cells("A2:B2")

    # 3-4: header
    box(ws, "A3:N4", LIGHT)
    put(ws, "A3", "채점 항목", font(17, b=True), LIGHT)
    put(ws, "B3", "채점 항목", font(17, b=True), LIGHT)
    ws.merge_cells("A3:A4")
    ws.merge_cells("B3:B4")
    for i in range(ATTEMPTS):
        p, c = get_column_letter(3 + 2 * i), get_column_letter(4 + 2 * i)
        put(ws, f"{p}3", f"시도 # {i + 1}", font(18, b=True), LIGHT)
        ws.merge_cells(f"{p}3:{c}3")
        put(ws, f"{p}4", "배점", font(10, b=True), LIGHT)
        put(ws, f"{c}4", "체크", font(10, b=True), LIGHT)

    # items
    first = r = 5
    cat_start = first
    for idx, (cat, text, pts) in enumerate(ITEMS):
        ws.row_dimensions[r].height = px_to_pt(ITEM_ROW_PX)
        if idx == 0 or ITEMS[idx - 1][0] != cat:
            cat_start = r
            put(ws, f"A{r}", cat, font(16))
        put(ws, f"B{r}", text, font(13))
        for i in range(ATTEMPTS):
            put(ws, f"{get_column_letter(3 + 2 * i)}{r}", pts, font(15), h="right")
        if (idx == len(ITEMS) - 1 or ITEMS[idx + 1][0] != cat) and r > cat_start:
            ws.merge_cells(f"A{cat_start}:A{r}")
        r += 1
    box(ws, f"A{first}:N{r - 1}")

    # 총합 (2 rows, merged per column like TASK B)
    t = r
    for k in (0, 1):
        ws.row_dimensions[t + k].height = px_to_pt(SMALL_ROW_PX)
    box(ws, f"A{t}:N{t + 1}")
    put(ws, f"A{t}", "총합", font(15))
    ws.merge_cells(f"A{t}:B{t + 1}")
    for i in range(ATTEMPTS):
        p, c = get_column_letter(3 + 2 * i), get_column_letter(4 + 2 * i)
        put(ws, f"{p}{t}", "합계", font(15))
        ws.merge_cells(f"{p}{t}:{p}{t + 1}")
        ws.merge_cells(f"{c}{t}:{c}{t + 1}")

    # 채점관 유의 사항 + 메모
    n = t + 2
    for k in (0, 1):
        ws.row_dimensions[n + k].height = px_to_pt(SMALL_ROW_PX)
    ws.row_dimensions[n + 2].height = px_to_pt(NOTES_ROW_PX)
    box(ws, f"A{n}:J{n + 1}", GREY)
    box(ws, f"K{n}:N{n + 2}")
    box(ws, f"A{n + 2}:J{n + 2}")
    put(ws, f"A{n}", "채점관 유의 사항", font(17), GREY)
    ws.merge_cells(f"A{n}:J{n + 1}")
    put(ws, f"K{n}", "메모)", font(15), v="top")
    ws.merge_cells(f"K{n}:N{n + 2}")
    put(ws, f"A{n + 2}", NOTES_LEFT, font(12))
    ws.merge_cells(f"A{n + 2}:D{n + 2}")
    put(ws, f"E{n + 2}", NOTES_RIGHT, font(12))
    ws.merge_cells(f"E{n + 2}:J{n + 2}")

    ws.print_area = f"A1:N{n + 2}"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    out = Path(__file__).with_name("task_a_scoresheet.xlsx")
    wb.save(out)
    print(f"wrote {out}")


if __name__ == "__main__":
    main()
