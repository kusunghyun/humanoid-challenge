"""Build the TASK A scoring sheet (.cell, 한셀) from the TASK B .cell file, keeping its formatting.

Usage: python3 scoresheets/make_task_a_cell.py
  template: scoresheets/templates/task_b.cell   (TASK B 채점표, 한셀에서 만든 원본)
  output:   scoresheets/task_a.cell  (+ task_a.xlsx, the same file under an Excel name)

The .cell file is an OOXML zip, so the sheet XML is edited directly: every cell keeps the
style index (s="..") of the matching TASK B cell, and only the values, rows and merges change.
"""
import math
import re
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape

HERE = Path(__file__).parent
TEMPLATE = HERE / "templates" / "task_b.cell"
OUT = HERE / "task_a.cell"

TITLE = "TASK A 채점표"
ATTEMPT_COLS = "CEGIKM"  # 배점 columns; 체크 is the next column

# (대분류, 채점 항목, 배점) — 시도 1~6 모두 같은 배점
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

# 채점관 유의 사항: 왼쪽(A:D) 3줄, 오른쪽(E:J) 3줄 — TASK B와 같은 자리
NOTES_LEFT = [
    "- A-(n)을 모두 통과하지 못하면 A-(n+1,+2,+3)를 획득하지 못한다. (A-4 는 A-2 까지 통과하면 획득할 수 있다.)",
    "- 가구에 살짝 스쳤지만 가구가 움직이지 않았으면 A-4 의 4점만 잃고 시도는 계속한다.",
    "- 6초는 마지막으로 손을 뗀 때부터 센다. 다시 쥐고 고쳐 놓으면 마지막 놓기부터 다시 6초. "
    "1초 미만의 순간 접촉은 다시 쥔 것으로 보지 않는다. 손을 끝까지 떼지 않으면 6초 항목은 0점.",
]
NOTES_RIGHT = [
    "- 바닥 표시: 도착 구역 = 책상 상판 중심 기준 반지름 1.0 m 원.",
    "- 책상 밀림 허용 = 다리 표시선에서 5 cm.",
    "- 6번 시도 후에 좌석 바꾸기",
]

# Row-height model taken from the TASK B sheet: ~4.5 characters per line in column B, 16.5pt per line.
ITEM_CHARS_PER_LINE = 4.5
NOTE_CHARS_PER_LINE = 36  # 4 merged default-width columns
LINE_PT = 16.5


class Strings:
    def __init__(self):
        self.items, self.index = [], {}

    def add(self, text):
        if text not in self.index:
            self.index[text] = len(self.items)
            self.items.append(text)
        return self.index[text]

    def xml(self, head):
        body = "".join(f'<x:si><x:t xml:space="preserve">{escape(t)}</x:t></x:si>' for t in self.items)
        head = re.sub(r'count="\d+"', f'count="{len(self.items)}"', head)
        head = re.sub(r'uniqueCount="\d+"', f'uniqueCount="{len(self.items)}"', head)
        return head + body + "</x:sst>"


def cell_re(col, row):
    return re.compile(rf'<x:c r="{col}{row}"( s="\d+")?( t="\w+")?(?:/>|>.*?</x:c>)', re.S)


def set_cell(row_xml, col, row, value, strings):
    def repl(m):
        style = m.group(1) or ""
        if value is None:
            return f'<x:c r="{col}{row}"{style}/>'
        if isinstance(value, str):
            return f'<x:c r="{col}{row}"{style} t="s"><x:v>{strings.add(value)}</x:v></x:c>'
        return f'<x:c r="{col}{row}"{style}><x:v>{value}</x:v></x:c>'

    new, n = cell_re(col, row).subn(repl, row_xml)
    assert n == 1, (col, row)
    return new


def renumber(row_xml, old, new):
    row_xml = re.sub(rf'<x:row r="{old}"', f'<x:row r="{new}"', row_xml)
    return re.sub(rf'r="([A-Z]+){old}"', rf'r="\g<1>{new}"', row_xml)


def set_height(row_xml, ht):
    row_xml = re.sub(r' ht="[\d.]+"', "", row_xml, count=1)
    return row_xml.replace(' spans="1:15"', f' spans="1:15" ht="{ht}"', 1)


def keep_strings(row_xml, row, old_strings, strings):
    """Re-map shared-string indexes of a row copied as-is to the new string table."""
    def repl(m):
        return f'{m.group(1)}<x:v>{strings.add(old_strings[int(m.group(2))])}</x:v>'
    return re.sub(r'(<x:c r="[A-Z]+\d+"(?: s="\d+")? t="s">)<x:v>(\d+)</x:v>', repl, row_xml)


def main():
    with zipfile.ZipFile(TEMPLATE) as z:
        files = {n: z.read(n) for n in z.namelist()}
    sheet = files["xl/worksheets/sheet1.xml"].decode("utf-8")
    sst = files["xl/sharedStrings.xml"].decode("utf-8")
    old_strings = [
        "".join(re.findall(r"<x:t[^>]*>(.*?)</x:t>", si, re.S))
        for si in re.findall(r"<x:si>(.*?)</x:si>", sst, re.S)
    ]
    sst_head = sst[: sst.index("<x:si>")]

    rows = {int(m.group(1)): m.group(0) for m in re.finditer(r'<x:row r="(\d+)".*?</x:row>', sheet, re.S)}
    strings = Strings()
    out_rows, merges = [], []

    # 1-4: title / team / header — unchanged except the title
    for r in range(1, 5):
        out_rows.append(keep_strings(rows[r], r, old_strings, strings))
    out_rows[0] = set_cell(out_rows[0], "A", 1, TITLE, strings)
    merges += ["A1:F1", "G1:N1", "A2:B2", "A3:A4", "B3:B4"]
    merges += [f"{c}3:{chr(ord(c) + 1)}3" for c in ATTEMPT_COLS]

    # items: first row of a category uses TASK B row 5 (category label), others row 6
    r = 5
    cat_start = r
    for i, (cat, text, pts) in enumerate(ITEMS):
        first_of_cat = i == 0 or ITEMS[i - 1][0] != cat
        tpl_row = 5 if first_of_cat else 6
        x = renumber(rows[tpl_row], tpl_row, r)
        x = keep_strings(x, r, old_strings, strings)
        if first_of_cat:
            x = set_cell(x, "A", r, cat, strings)
            cat_start = r
        x = set_cell(x, "B", r, text, strings)
        for c in ATTEMPT_COLS:
            x = set_cell(x, c, r, pts, strings)
        x = set_height(x, LINE_PT * max(3, math.ceil(len(text) / ITEM_CHARS_PER_LINE)))
        out_rows.append(x)
        last_of_cat = i == len(ITEMS) - 1 or ITEMS[i + 1][0] != cat
        if last_of_cat and r > cat_start:
            merges.append(f"A{cat_start}:A{r}")
        r += 1

    # 총합 (TASK B rows 20-21)
    total = r
    for k, old in enumerate((20, 21)):
        out_rows.append(keep_strings(renumber(rows[old], old, total + k), total + k, old_strings, strings))
    merges.append(f"A{total}:B{total + 1}")
    merges += [f"{chr(ord(c) + d)}{total}:{chr(ord(c) + d)}{total + 1}" for c in ATTEMPT_COLS for d in (0, 1)]

    # 채점관 유의 사항 (TASK B rows 22-26)
    nr = total + 2
    for k, old in enumerate(range(22, 27)):
        x = keep_strings(renumber(rows[old], old, nr + k), nr + k, old_strings, strings)
        if k >= 2:
            left, right = NOTES_LEFT[k - 2], NOTES_RIGHT[k - 2]
            x = set_cell(x, "A", nr + k, left, strings)
            x = set_cell(x, "E", nr + k, right, strings)
            lines = max(math.ceil(len(left) / NOTE_CHARS_PER_LINE), math.ceil(len(right) / NOTE_CHARS_PER_LINE))
            x = set_height(x, max(18, LINE_PT * lines))
        out_rows.append(x)
    merges += [f"A{nr}:J{nr + 1}", f"K{nr}:N{nr + 4}"]
    merges += [f"A{nr + k}:D{nr + k}" for k in (2, 3, 4)] + [f"E{nr + k}:J{nr + k}" for k in (2, 3, 4)]

    last_row = nr + 4
    head = sheet[: sheet.index("<x:sheetData>")]
    head = re.sub(r'<x:dimension ref="[^"]+"/>', f'<x:dimension ref="A1:O{last_row}"/>', head)
    head = head.replace('topLeftCell="A10"', 'topLeftCell="A1"').replace('sqref="A1:O26"', 'sqref="A1"')
    tail = sheet[sheet.index("</x:sheetData>"):]
    merge_xml = f'<x:mergeCells count="{len(merges)}">' + "".join(f'<x:mergeCell ref="{m}"/>' for m in merges) + "</x:mergeCells>"
    tail = re.sub(r"<x:mergeCells.*?</x:mergeCells>", merge_xml, tail, flags=re.S)

    files["xl/worksheets/sheet1.xml"] = (head + "<x:sheetData>" + "".join(out_rows) + tail).encode("utf-8")
    files["xl/sharedStrings.xml"] = strings.xml(sst_head).encode("utf-8")

    for path in (OUT, OUT.with_suffix(".xlsx")):
        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
            for name, data in files.items():
                z.writestr(name, data)
        print(f"wrote {path}")


if __name__ == "__main__":
    main()
