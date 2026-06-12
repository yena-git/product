#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
셀러 자체 엑셀 → 그립 상품 양식 자동 변환 PoC
=================================================
승인된 플랜(distributed-jumping-crescent.md)의 변환 엔진 개념 증명.

흐름:
  1) 셀러 엑셀 파싱
  2) 필드 매핑 적용 (셀러 컬럼 → 그립 필드 / 옵션 축)
  3) 상품명 기준 그룹핑
  4) 옵션 축 합성(wide→long pivot) — 값이 있는 축만 채택해 '옵션 종류명' 생성
  5) 기본 정규화(trim·공백압축) 후 distinct 옵션값 수집(첫 등장 순서 보존)
  6) 데카르트곱 생성 → 누락 조합을 판매 off / 재고 0 으로 자동 보완
  7) 그립 '플래시상품' 양식 xlsx 출력 + 요약 리포트

표준 라이브러리만 사용(openpyxl 불필요). 출력 xlsx는 'off 표시용' 보조 컬럼
(판매여부)를 포함 — 실서비스에서는 등록 API가 옵션별 판매상태 off로 처리.
"""

import sys
import re
import zipfile
import html
import xml.etree.ElementTree as ET
from itertools import product
from collections import OrderedDict, defaultdict

NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"

# 그립 제약
MAX_ROWS_PER_FILE = 10_000
MAX_COMBOS_PER_PRODUCT = 2_000

# ── 매핑 설정 ────────────────────────────────────────────────────────────────
# 실서비스에서는 2단계 매핑 UI가 만들어 주는 값. PoC에서는 셀러 파일 기준 기본값.
#   - FIELD_MAP : 셀러 컬럼명 → 그립 상품 필드
#   - AXIS_COLS : 옵션 축으로 취급할 셀러 컬럼명(여러 개 → 옵션1/2/3로 pivot)
FIELD_MAP = {
    "사이트 상품명": "상품명",
    "라방판매가": "라이브가",
    "잔여재고": "재고",
    "품목코드": "자체코드",  # 플래시 양식엔 미사용. 리포트 참고용
}
AXIS_COLS = ["색상", "사이즈"]  # 등장 순서가 곧 옵션1, 옵션2

# 그립 플래시상품 양식 헤더 (+ PoC 표시용 보조 컬럼)
GRIP_FLASH_HEADER = [
    "카테고리", "상품 정보고시 상품군", "상품명 (필수)", "대표 이미지 (선택)",
    "라이브가 (필수)", "옵션 종류명 (선택)", "옵션1 (선택)", "옵션2 (선택)",
    "옵션3 (선택)", "추가 금액 (선택)", "재고 수량 (선택)", "판매여부(PoC표시용)",
]


# ── 정규화(1층) ──────────────────────────────────────────────────────────────
def normalize(s):
    """trim + 연속공백 1칸 압축. (전/반각 통일은 PoC에서 생략)"""
    if s is None:
        return ""
    return re.sub(r"\s+", " ", str(s)).strip()


# 이모지·장식 특수기호(★ ♡ ♥ 등) 제거 — 프로토타입(index.html)과 동일 범위
_EMOJI_RE = re.compile(
    "[\U0001F000-\U0001FAFF←-⇿⌀-➿"
    "⬀-⯿☀-⛿︀-️‍]",
    flags=re.UNICODE,
)


def strip_emoji(s):
    return normalize(_EMOJI_RE.sub("", str(s or "")))


# ── xlsx 읽기 (표준 라이브러리) ──────────────────────────────────────────────
def _col_to_num(ref):
    n = 0
    for c in ref:
        if c.isalpha():
            n = n * 26 + (ord(c.upper()) - 64)
        else:
            break
    return n


def read_xlsx(path, sheet_index=0):
    """첫 시트를 [[셀, ...], ...] (헤더 포함) 로 반환."""
    z = zipfile.ZipFile(path)
    shared = []
    if "xl/sharedStrings.xml" in z.namelist():
        root = ET.fromstring(z.read("xl/sharedStrings.xml"))
        for si in root.findall(NS + "si"):
            shared.append("".join(t.text or "" for t in si.iter(NS + "t")))
    sheets = sorted(
        n for n in z.namelist() if re.match(r"xl/worksheets/sheet\d+\.xml$", n)
    )
    root = ET.fromstring(z.read(sheets[sheet_index]))
    rows = {}
    maxcol = 0
    for row in root.iter(NS + "row"):
        for c in row.iter(NS + "c"):
            ref = c.get("r")
            t = c.get("t")
            col = _col_to_num("".join(ch for ch in ref if ch.isalpha()))
            rnum = int("".join(ch for ch in ref if ch.isdigit()))
            v = c.find(NS + "v")
            isv = c.find(NS + "is")
            if t == "s" and v is not None:
                val = shared[int(v.text)]
            elif isv is not None:
                val = "".join(tt.text or "" for tt in isv.iter(NS + "t"))
            elif v is not None:
                val = v.text
            else:
                val = ""
            rows.setdefault(rnum, {})[col] = val
            maxcol = max(maxcol, col)
    out = []
    for r in sorted(rows):
        out.append([rows[r].get(c, "") for c in range(1, maxcol + 1)])
    return out


# ── xlsx 쓰기 (표준 라이브러리, inline string) ───────────────────────────────
def _cell_xml(col_idx, row_idx, value):
    ref = ""
    n = col_idx
    while n > 0:
        n, rem = divmod(n - 1, 26)
        ref = chr(65 + rem) + ref
    ref = f"{ref}{row_idx}"
    if value is None or value == "":
        return f'<c r="{ref}"/>'
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return f'<c r="{ref}"><v>{value}</v></c>'
    txt = html.escape(str(value))
    return f'<c r="{ref}" t="inlineStr"><is><t xml:space="preserve">{txt}</t></is></c>'


def write_xlsx(path, rows):
    sheet_rows = []
    for ri, row in enumerate(rows, start=1):
        cells = "".join(_cell_xml(ci, ri, v) for ci, v in enumerate(row, start=1))
        sheet_rows.append(f'<row r="{ri}">{cells}</row>')
    sheet = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        f"<sheetData>{''.join(sheet_rows)}</sheetData></worksheet>"
    )
    content_types = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        "</Types>"
    )
    rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
        "</Relationships>"
    )
    workbook = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<sheets><sheet name="변환결과" sheetId="1" r:id="rId1"/></sheets></workbook>'
    )
    wb_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
        "</Relationships>"
    )
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types)
        z.writestr("_rels/.rels", rels)
        z.writestr("xl/workbook.xml", workbook)
        z.writestr("xl/_rels/workbook.xml.rels", wb_rels)
        z.writestr("xl/worksheets/sheet1.xml", sheet)


# ── 변환 엔진 ────────────────────────────────────────────────────────────────
def convert(seller_path):
    grid = read_xlsx(seller_path)
    header = [normalize(h) for h in grid[0]]
    col_idx = {name: i for i, name in enumerate(header)}

    def cell(row, name):
        i = col_idx.get(name)
        return normalize(row[i]) if i is not None and i < len(row) else ""

    data = grid[1:]

    name_col = next(k for k, v in FIELD_MAP.items() if v == "상품명")
    price_col = next((k for k, v in FIELD_MAP.items() if v == "라이브가"), None)
    stock_col = next((k for k, v in FIELD_MAP.items() if v == "재고"), None)

    # 상품명 기준 그룹핑 (첫 등장 순서 보존). 상품명 이모지는 제거 후 판단/등록.
    groups = OrderedDict()
    report_emoji = 0
    for row in data:
        raw = cell(row, name_col)
        if not raw:
            continue
        nm = strip_emoji(raw)
        if nm != raw:
            report_emoji += 1
        groups.setdefault(nm, []).append(row)

    out_rows = [GRIP_FLASH_HEADER]
    report = {
        "products": len(groups),
        "input_rows": len(data),
        "input_combos": 0,
        "output_combos": 0,
        "auto_off": 0,
        "over_2000": [],
        "missing_price": [],
        "emoji_removed": report_emoji,
    }

    for nm, rows in groups.items():
        # 활성 옵션 축: 그룹 내 값이 하나라도 있는 축만
        active_axes = [ax for ax in AXIS_COLS if any(cell(r, ax) for r in rows)]
        옵션종류명 = ", ".join(active_axes)

        # 라이브가: 그룹 첫 유효값
        price = next((cell(r, price_col) for r in rows if price_col and cell(r, price_col)), "")
        if not price:
            report["missing_price"].append(nm)

        # 옵션 없는 상품
        if not active_axes:
            stock = cell(rows[0], stock_col) if stock_col else ""
            out_rows.append(["", "", nm, "", _num(price), "", "", "", "", "", _num(stock), "판매"])
            report["input_combos"] += 1
            report["output_combos"] += 1
            continue

        # 축별 distinct 값(정규화 후, 첫 등장 순서 보존)
        axis_values = []
        for ax in active_axes:
            seen = OrderedDict()
            for r in rows:
                v = cell(r, ax)
                if v:
                    seen[v] = True
            axis_values.append(list(seen.keys()))

        # 입력 조합 → {옵션튜플: (재고, 추가금액)}
        present = {}
        for r in rows:
            key = tuple(cell(r, ax) for ax in active_axes)
            if "" in key:
                continue  # 일부 축만 채워진 비정상 행은 스킵
            present[key] = (cell(r, stock_col) if stock_col else "", "")
        report["input_combos"] += len(present)

        full = list(product(*axis_values))
        if len(full) > MAX_COMBOS_PER_PRODUCT:
            report["over_2000"].append((nm, len(full)))
            # 초과 상품은 입력분만 등록(자동보완 생략)
            full = list(present.keys())

        first = True
        for combo in full:
            o1 = combo[0] if len(combo) > 0 else ""
            o2 = combo[1] if len(combo) > 1 else ""
            o3 = combo[2] if len(combo) > 2 else ""
            if combo in present:
                stock, extra = present[combo]
                sale = "판매"
            else:
                stock, extra = "0", ""
                sale = "off(자동추가)"
                report["auto_off"] += 1
            # 상품명은 동일 상품 판단 기준 → 모든 옵션 조합 행에 입력.
            # 라이브가·옵션 종류명은 그립 양식대로 그룹 첫 행에만.
            row = [
                "", "", nm, "",
                (_num(price) if first else ""),
                (옵션종류명 if first else ""),
                o1, o2, o3, _num(extra), _num(stock), sale,
            ]
            out_rows.append(row)
            report["output_combos"] += 1
            first = False

    return out_rows, report


def _num(s):
    """숫자 문자열이면 int/float로, 아니면 원문(빈값은 '')."""
    if s == "" or s is None:
        return ""
    t = str(s).replace(",", "")
    try:
        f = float(t)
        return int(f) if f.is_integer() else f
    except ValueError:
        return s


# ── 실행 ─────────────────────────────────────────────────────────────────────
def main():
    seller = sys.argv[1] if len(sys.argv) > 1 else "셀러_상품 리스트.xlsx"
    out_path = sys.argv[2] if len(sys.argv) > 2 else "변환결과_그립_플래시상품.xlsx"

    out_rows, rep = convert(seller)

    total_data_rows = len(out_rows) - 1  # 헤더 제외
    if total_data_rows > MAX_ROWS_PER_FILE:
        print(f"⚠️  변환 결과 {total_data_rows}행 — 파일당 1만 row 초과. 분할 필요.")

    write_xlsx(out_path, out_rows)

    print("=" * 60)
    print("  셀러 엑셀 → 그립 플래시상품 양식 변환 PoC 결과")
    print("=" * 60)
    print(f"  입력 파일        : {seller}")
    print(f"  출력 파일        : {out_path}")
    print(f"  상품 수          : {rep['products']:,}")
    print(f"  셀러 입력 행     : {rep['input_rows']:,}")
    print(f"  입력 조합        : {rep['input_combos']:,}")
    print(f"  출력 조합        : {rep['output_combos']:,}  (그립 양식 데이터 행)")
    print(f"  자동 추가 off 조합 : {rep['auto_off']:,}  ← 누락 조합 자동 보완")
    print(f"  이모지 제거 행   : {rep['emoji_removed']:,}  ← 상품명 이모지/특수기호 제거")
    if rep["over_2000"]:
        print(f"  ⚠️ 2000조합 초과 상품 : {len(rep['over_2000'])}건 (자동보완 생략)")
        for nm, n in rep["over_2000"][:5]:
            print(f"       - {nm}: {n:,}조합")
    if rep["missing_price"]:
        print(f"  ⚠️ 라이브가 누락(필수) : {len(rep['missing_price'])}건")
        for nm in rep["missing_price"][:5]:
            print(f"       - {nm}")
    print("=" * 60)


if __name__ == "__main__":
    main()
