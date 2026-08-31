from __future__ import annotations

import copy
import datetime as dt
import html
import json
import re
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


ROOT = Path(__file__).resolve().parents[1]
TEMPLATE = ROOT / "inputTemplate.xlsx"
OUTPUT_DIR = ROOT / "test_cases"
INPUT_DIR = OUTPUT_DIR / "inputs"
RESULT_FILE = OUTPUT_DIR / "test_results.xlsx"

HEADERS = [
    "이름",
    "학교",
    "학년",
    "리포트시점",
    "종합평가",
    "실기",
    "내신",
    "묘사력",
    "형태력",
    "사고력",
    "완성도",
    "멘탈관리",
    "현재 상황 지원 가능한 학교",
    "발전 후 지원 가능한 학교",
]

REQUIRED_COLUMNS = [
    "이름",
    "학교",
    "학년",
    "리포트시점",
    "종합평가",
    "실기",
    "내신",
    "묘사력",
    "형태력",
    "사고력",
    "완성도",
    "멘탈관리",
    "현재 상황 지원 가능한 학교",
    "발전 후 지원 가능한 학교",
]

DRAWING_METRICS = ["형태력", "사고력", "멘탈관리", "완성도", "묘사력"]
NUMERIC_COLUMNS = ["학년", "실기", "내신", *DRAWING_METRICS]
BASE_DATE = dt.date(1899, 12, 30)
TEMPLATE_DOWNLOAD_HINT = "사용 방법의 '템플릿 다운로드'에서 받을 수 있습니다."

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
NS_X14AC = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"
NS_XR = "http://schemas.microsoft.com/office/spreadsheetml/2014/revision"
NS_XR2 = "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2"
NS_XR3 = "http://schemas.microsoft.com/office/spreadsheetml/2016/revision3"

ET.register_namespace("", NS_MAIN)
ET.register_namespace("mc", NS_MC)
ET.register_namespace("x14ac", NS_X14AC)
ET.register_namespace("xr", NS_XR)
ET.register_namespace("xr2", NS_XR2)
ET.register_namespace("xr3", NS_XR3)


def qname(ns: str, name: str) -> str:
    return f"{{{ns}}}{name}"


def col_name(index: int) -> str:
    name = ""
    while index:
        index, rem = divmod(index - 1, 26)
        name = chr(65 + rem) + name
    return name


def excel_serial(value: dt.date) -> int:
    return (value - BASE_DATE).days


def iso_date(value: str) -> str:
    return value


def default_rows(name: str = "테스트학생") -> list[dict[str, object]]:
    return [
        {
            "이름": name,
            "학교": "아트중",
            "학년": 2,
            "리포트시점": dt.date(2026, 1, 1),
            "종합평가": "기본기와 수업 태도가 안정적이며, 세부 묘사와 발상 확장이 다음 과제입니다.",
            "실기": 72,
            "내신": 84,
            "묘사력": 66,
            "형태력": 74,
            "사고력": 62,
            "완성도": 70,
            "멘탈관리": 68,
            "현재 상황 지원 가능한 학교": "A예중, B예고",
            "발전 후 지원 가능한 학교": "C예고",
        },
        {
            "리포트시점": dt.date(2026, 2, 1),
            "종합평가": "최근 과제 완성도가 좋아졌고, 형태 관찰이 더 정확해졌습니다.",
            "실기": 78,
            "내신": 86,
            "묘사력": 70,
            "형태력": 79,
            "사고력": 66,
            "완성도": 76,
            "멘탈관리": 72,
            "현재 상황 지원 가능한 학교": "B예고",
            "발전 후 지원 가능한 학교": "C예고, D예고",
        },
        {
            "리포트시점": dt.date(2026, 3, 1),
            "종합평가": "실전 시간 안배가 개선되었고, 목표 학교 기준에 근접했습니다.",
            "실기": 82,
            "내신": 88,
            "묘사력": 75,
            "형태력": 83,
            "사고력": 71,
            "완성도": 80,
            "멘탈관리": 76,
            "현재 상황 지원 가능한 학교": "C예고",
            "발전 후 지원 가능한 학교": "D예고, E예고",
        },
    ]


def complete_row(name: str, index: int = -1) -> dict[str, object]:
    rows = default_rows(name)
    return {**rows[0], **rows[index]}


def long_evaluation() -> str:
    sentence = (
        "이 학생은 관찰력과 화면 구성력에서 꾸준한 성장을 보이고 있으나, "
        "시험 상황에서는 아이디어를 정리하는 시간이 길어져 완성도에 영향을 받습니다. "
        "다음 달에는 발상 스케치 시간을 제한하고, 질감별 묘사 루틴을 반복하는 방식으로 보완합니다. "
    )
    return sentence * 18


def scenario_rows() -> list[dict[str, object]]:
    return [
        {"case_id": "TC01", "filename": "tc01_valid_progress.xlsx", "description": "정상 3개월 성장 데이터", "rows": default_rows("정상케이스")},
        {
            "case_id": "TC02",
            "filename": "tc02_long_evaluation.xlsx",
            "description": "종합평가가 매우 긴 경우",
            "rows": [*default_rows("긴평가"), {**complete_row("긴평가"), "리포트시점": dt.date(2026, 4, 1), "종합평가": long_evaluation()}],
        },
        {
            "case_id": "TC03",
            "filename": "tc03_scores_out_of_range.xlsx",
            "description": "점수가 0~100 범위를 벗어나는 경우",
            "rows": [{**complete_row("범위초과"), "실기": 135, "내신": -20, "묘사력": 150, "형태력": -5, "사고력": 101, "완성도": 999, "멘탈관리": -1}],
        },
        {
            "case_id": "TC04",
            "filename": "tc04_missing_student_info.xlsx",
            "description": "학생 기본 정보가 모두 비어 있는 경우",
            "rows": [{**complete_row("정보누락"), "이름": "", "학교": "", "학년": ""}],
        },
        {
            "case_id": "TC05",
            "filename": "tc05_missing_required_column.xlsx",
            "description": "필수 컬럼 '완성도'가 빠진 경우",
            "headers": [h for h in HEADERS if h != "완성도"],
            "rows": default_rows("컬럼누락"),
        },
        {
            "case_id": "TC06",
            "filename": "tc06_invalid_dates.xlsx",
            "description": "리포트시점이 비어 있거나 해석 불가한 경우",
            "rows": [
                {**default_rows("날짜오류")[0], "리포트시점": "not-a-date"},
                {**default_rows("날짜오류")[1], "리포트시점": ""},
            ],
        },
        {
            "case_id": "TC07",
            "filename": "tc07_blank_last_identity_fallback.xlsx",
            "description": "마지막 행 이름/학교/학년이 비어 첫 행 값으로 대체되는 경우",
            "rows": [default_rows("대체확인")[0], {**default_rows("대체확인")[-1], "이름": "", "학교": "", "학년": ""}],
        },
        {
            "case_id": "TC08",
            "filename": "tc08_zero_detail_scores.xlsx",
            "description": "실기 세부항목 점수가 0인 경우",
            "rows": [{**complete_row("제로점수"), "묘사력": 0, "형태력": 0, "사고력": 0, "완성도": 0, "멘탈관리": 0}],
        },
        {
            "case_id": "TC09",
            "filename": "tc09_text_and_comma_numbers.xlsx",
            "description": "문자열 숫자와 콤마 포함 숫자 입력",
            "rows": [{**complete_row("문자숫자"), "학년": "2", "실기": "88", "내신": "1,000", "묘사력": "77", "형태력": " 82 ", "사고력": "69", "완성도": "91", "멘탈관리": "73"}],
        },
        {
            "case_id": "TC10",
            "filename": "tc10_unsorted_dates.xlsx",
            "description": "행 순서상 마지막 날짜가 더 이전인 경우",
            "rows": [
                {**default_rows("날짜역순")[0], "리포트시점": dt.date(2026, 5, 1)},
                {**complete_row("날짜역순", 1), "리포트시점": dt.date(2026, 3, 1)},
            ],
        },
        {
            "case_id": "TC11",
            "filename": "tc11_header_only.xlsx",
            "description": "헤더만 있고 데이터 행이 없는 경우",
            "rows": [],
        },
        {
            "case_id": "TC12",
            "filename": "tc12_wrong_univ_template_marker.xlsx",
            "description": "'성적' 컬럼이 있어 미대 입시용 템플릿으로 판단되는 경우",
            "headers": [*HEADERS, "성적"],
            "rows": [{**complete_row("잘못된템플릿"), "성적": 95}],
        },
        {
            "case_id": "TC13",
            "filename": "tc13_missing_optional_current_school.xlsx",
            "description": "현재 상황 지원 가능한 학교 컬럼이 빠진 경우",
            "headers": [h for h in HEADERS if h != "현재 상황 지원 가능한 학교"],
            "rows": default_rows("선택컬럼누락"),
        },
        {
            "case_id": "TC14",
            "filename": "tc14_score_units.xlsx",
            "description": "점수에 '점' 단위가 붙어 있는 경우",
            "rows": [{**default_rows("단위점수")[0], "실기": "88점", "내신": "85점", "묘사력": "77점", "형태력": "82점", "사고력": "69점", "완성도": "90점", "멘탈관리": "73점"}],
        },
        {
            "case_id": "TC15",
            "filename": "tc15_non_numeric_scores.xlsx",
            "description": "숫자 칸에 숫자가 전혀 없는 문자를 넣은 경우",
            "rows": [{**complete_row("문자점수"), "실기": "상", "내신": "중", "묘사력": "N/A", "형태력": "", "사고력": "칠십", "완성도": "-", "멘탈관리": "없음"}],
        },
    ]


def cell_xml(ref: str, value: object, style: str) -> str:
    attrs = f' r="{ref}" s="{style}"'
    if value is None or value == "":
        return f"<c{attrs}/>"
    if isinstance(value, dt.date):
        return f"<c{attrs}><v>{excel_serial(value)}</v></c>"
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return f"<c{attrs}><v>{value}</v></c>"
    text = html.escape(str(value), quote=False)
    return f'<c{attrs} t="inlineStr"><is><t>{text}</t></is></c>'


def build_sheet_xml(headers: list[str], rows: list[dict[str, object]]) -> bytes:
    row_count = max(1, len(rows) + 1)
    col_count = len(headers)
    dimension = f"A1:{col_name(col_count)}{row_count}"
    xml_rows = []

    header_cells = [cell_xml(f"{col_name(i)}1", header, "1" if i != 4 else "2") for i, header in enumerate(headers, start=1)]
    xml_rows.append(f'<row r="1" spans="1:{col_count}" x14ac:dyDescent="0.3">{"".join(header_cells)}</row>')

    for row_index, row in enumerate(rows, start=2):
        cells = []
        for col_index, header in enumerate(headers, start=1):
            style = "2" if header == "리포트시점" else "1"
            cells.append(cell_xml(f"{col_name(col_index)}{row_index}", row.get(header, ""), style))
        height = ' ht="67.5"' if len(str(row.get("종합평가", ""))) > 140 else ""
        xml_rows.append(f'<row r="{row_index}" spans="1:{col_count}"{height} x14ac:dyDescent="0.3">{"".join(cells)}</row>')

    cols = (
        '<cols>'
        '<col min="1" max="1" width="8.64453125" style="1" customWidth="1"/>'
        '<col min="2" max="2" width="14.3515625" style="1" customWidth="1"/>'
        '<col min="3" max="3" width="8.64453125" style="1" customWidth="1"/>'
        '<col min="4" max="4" width="13" style="2" customWidth="1"/>'
        '<col min="5" max="5" width="64" style="1" customWidth="1"/>'
        '<col min="6" max="14" width="12" style="1" customWidth="1"/>'
        '</cols>'
    )

    xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="{NS_MAIN}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="{NS_MC}" mc:Ignorable="x14ac xr xr2 xr3" xmlns:x14ac="{NS_X14AC}" xmlns:xr="{NS_XR}" xmlns:xr2="{NS_XR2}" xmlns:xr3="{NS_XR3}">
<dimension ref="{dimension}"/>
<sheetViews><sheetView tabSelected="1" zoomScale="70" zoomScaleNormal="70" workbookViewId="0"/></sheetViews>
<sheetFormatPr defaultRowHeight="13.5" x14ac:dyDescent="0.3"/>
{cols}
<sheetData>{"".join(xml_rows)}</sheetData>
<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
</worksheet>'''
    return xml.encode("utf-8")


def write_xlsx_from_template(path: Path, headers: list[str], rows: list[dict[str, object]]) -> None:
    sheet_xml = build_sheet_xml(headers, rows)
    with zipfile.ZipFile(TEMPLATE, "r") as zin, zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            if item.filename == "xl/worksheets/sheet1.xml":
                zout.writestr(item, sheet_xml)
            else:
                zout.writestr(item, zin.read(item.filename))


def read_shared_strings(zf: zipfile.ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in zf.namelist():
        return []
    root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
    strings = []
    for si in root.findall(qname(NS_MAIN, "si")):
        strings.append("".join(t.text or "" for t in si.findall(f".//{qname(NS_MAIN, 't')}")))
    return strings


def cell_value(cell: ET.Element, shared: list[str]) -> object:
    cell_type = cell.attrib.get("t")
    if cell_type == "inlineStr":
        return "".join(t.text or "" for t in cell.findall(f".//{qname(NS_MAIN, 't')}"))
    value = cell.find(qname(NS_MAIN, "v"))
    if value is None or value.text is None:
        return ""
    if cell_type == "s":
        return shared[int(value.text)]
    text = value.text
    try:
        number = float(text)
        return int(number) if number.is_integer() else number
    except ValueError:
        return text


def read_xlsx_rows(path: Path) -> list[dict[str, object]]:
    with zipfile.ZipFile(path) as zf:
        shared = read_shared_strings(zf)
        sheet = ET.fromstring(zf.read("xl/worksheets/sheet1.xml"))
    grid: dict[int, dict[int, object]] = {}
    for row in sheet.findall(f".//{qname(NS_MAIN, 'sheetData')}/{qname(NS_MAIN, 'row')}"):
        row_number = int(row.attrib["r"])
        grid[row_number] = {}
        for cell in row.findall(qname(NS_MAIN, "c")):
            ref = cell.attrib.get("r", "")
            col_letters = re.match(r"[A-Z]+", ref).group(0)
            col_index = 0
            for char in col_letters:
                col_index = col_index * 26 + ord(char) - 64
            grid[row_number][col_index] = cell_value(cell, shared)

    if not grid:
        return []
    headers = [str(grid[1].get(i, "")) for i in range(1, max(grid[1]) + 1)]
    rows = []
    for row_number in sorted(k for k in grid if k != 1):
        row = {header: grid[row_number].get(i, "") for i, header in enumerate(headers, start=1) if header}
        if any(value != "" for value in row.values()):
            rows.append(row)
    return rows


def parse_excel_date(value: object) -> dt.date | None:
    if isinstance(value, dt.date):
        return value
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return BASE_DATE + dt.timedelta(days=int(value))
    text = str(value or "").strip()
    if not text:
        return None
    try:
        return dt.datetime.fromisoformat(text).date()
    except ValueError:
        pass
    try:
        return BASE_DATE + dt.timedelta(days=int(float(text)))
    except ValueError:
        return None


def to_number(value: object) -> float:
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return float(value)
    text = str(value if value is not None else "").strip()
    match = re.search(r"[+-]?(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d+)?", text)
    if not match:
        return 0.0
    try:
        return float(match.group(0).replace(",", ""))
    except ValueError:
        return 0.0


def normalize_row(row: dict[str, object]) -> dict[str, object]:
    normalized = copy.deepcopy(row)
    normalized["리포트시점"] = parse_excel_date(row.get("리포트시점"))
    for key in NUMERIC_COLUMNS:
        normalized[key] = to_number(row.get(key))
    for key in DRAWING_METRICS:
        if normalized[key] == 0:
            normalized[key] = None
    return normalized


def validate_template(rows: list[dict[str, object]]) -> None:
    if not rows:
        raise ValueError(f"엑셀 파일에 데이터가 없습니다. {TEMPLATE_DOWNLOAD_HINT}")
    columns = list(rows[0].keys())
    if "성적" in columns:
        raise ValueError(f"예중/예고 입시용 템플릿이 아닌 미대 입시용 템플릿을 불러들였습니다. {TEMPLATE_DOWNLOAD_HINT}")
    missing = [column for column in REQUIRED_COLUMNS if column not in columns]
    if missing:
        raise ValueError(f"입력 템플릿 형식이 아닙니다. 필수 컬럼이 없습니다: {', '.join(missing)}. {TEMPLATE_DOWNLOAD_HINT}")


def has_value(value: object) -> bool:
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return value > 0
    return str(value or "").strip() != ""


def value_with_fallback(value: object, fallback: object) -> object:
    return value if has_value(value) else fallback


def validate_report_rows(rows: list[dict[str, object]]) -> None:
    first = rows[0]
    last = rows[-1]
    missing = []

    if not has_value(value_with_fallback(last.get("이름"), first.get("이름"))):
        missing.append("이름")
    if not has_value(value_with_fallback(last.get("학교"), first.get("학교"))):
        missing.append("학교")
    if not has_value(value_with_fallback(last.get("학년"), first.get("학년"))):
        missing.append("학년")

    if missing:
        raise ValueError(f"학생 기본 정보가 없습니다: {', '.join(missing)}")


def month_span(start: dt.date, end: dt.date) -> int:
    return (end.year - start.year) * 12 + end.month - start.month + 1


def app_like_result(path: Path) -> dict[str, object]:
    rows = read_xlsx_rows(path)
    validate_template(rows)
    normalized = [normalize_row(row) for row in rows]
    normalized = [row for row in normalized if row["리포트시점"] is not None]
    if not normalized:
        raise ValueError("리포트시점 형식이 올바르지 않습니다. 날짜는 2026-03-01처럼 YYYY-MM-DD 형식 또는 엑셀 날짜 형식으로 입력하세요.")
    validate_report_rows(normalized)
    first = normalized[0]
    last = normalized[-1]
    return {
        "status": "PASS",
        "rendered_rows": len(normalized),
        "student": last.get("이름") or first.get("이름") or "학생",
        "report_month": f"{last['리포트시점'].year}-{last['리포트시점'].month:02d}",
        "study_period": month_span(first["리포트시점"], last["리포트시점"]),
        "normalized_last": {key: last.get(key) for key in ["실기", "내신", *DRAWING_METRICS]},
        "notes": "",
    }


def classify_expected(case: dict[str, object], result: dict[str, object] | None, error: str | None) -> tuple[str, str]:
    case_id = case["case_id"]
    if error:
        return "오류 표시", error
    assert result is not None
    if case_id == "TC02":
        return "생성됨", "종합평가가 렌더링되지만 현재 UI 높이에서는 잘릴 수 있으므로 육안/PDF 확인 필요"
    if case_id == "TC03":
        return "생성됨", "점수 범위 검증은 없고 그래프는 0~100으로 클램프됨"
    if case_id == "TC04":
        return "생성됨", "학생 기본 정보 누락인데 생성됨. 검증 로직 확인 필요"
    if case_id == "TC07":
        return "생성됨", "마지막 행의 이름/학교/학년 공백은 첫 행 값으로 대체됨"
    if case_id == "TC08":
        return "생성됨", "세부항목 0점은 null로 바뀌어 오각형/월별 세부 그래프에서 표시되지 않음"
    if case_id == "TC09":
        return "생성됨", "문자열 숫자와 콤마는 숫자로 변환됨. 1,000은 1000으로 변환 후 그래프에서 100으로 클램프됨"
    if case_id == "TC10":
        return "생성됨", "행 순서를 그대로 사용하여 기간이 음수로 계산될 수 있음"
    if case_id == "TC13":
        return "생성됨", "현재 상황 지원 가능한 학교 컬럼도 기본 템플릿 컬럼이므로 없으면 오류 표시"
    if case_id == "TC14":
        return "생성됨", "'85점', '90점'처럼 단위가 붙은 점수에서 숫자만 추출됨"
    if case_id == "TC15":
        return "생성됨", "숫자가 전혀 없는 값은 0 처리됨. 실기 세부항목의 0은 그래프 선에서 누락됨"
    return "생성됨", "정상 생성"


def write_results_xlsx(results: list[dict[str, object]]) -> None:
    headers = [
        "case_id",
        "file",
        "description",
        "expected_focus",
        "actual_result",
        "status_or_error",
        "rendered_rows",
        "student",
        "report_month",
        "study_period_months",
        "normalized_last_scores",
        "notes",
    ]
    rows = []
    for result in results:
        rows.append({header: result.get(header, "") for header in headers})
    write_xlsx_from_template(RESULT_FILE, headers, rows)


def main() -> None:
    INPUT_DIR.mkdir(parents=True, exist_ok=True)
    cases = scenario_rows()
    expected_files = {case["filename"] for case in cases}
    for stale_file in INPUT_DIR.glob("*.xlsx"):
        if stale_file.name not in expected_files:
            stale_file.unlink()
    results = []

    for case in cases:
        headers = case.get("headers", HEADERS)
        path = INPUT_DIR / case["filename"]
        write_xlsx_from_template(path, headers, case["rows"])

        actual = None
        error = None
        try:
            actual = app_like_result(path)
        except Exception as exc:
            error = str(exc)

        actual_result, notes = classify_expected(case, actual, error)
        results.append(
            {
                "case_id": case["case_id"],
                "file": str(path.relative_to(ROOT)),
                "description": case["description"],
                "expected_focus": case["description"],
                "actual_result": actual_result,
                "status_or_error": error or "리포트 생성 가능",
                "rendered_rows": actual.get("rendered_rows", "") if actual else "",
                "student": actual.get("student", "") if actual else "",
                "report_month": actual.get("report_month", "") if actual else "",
                "study_period_months": actual.get("study_period", "") if actual else "",
                "normalized_last_scores": json.dumps(actual.get("normalized_last", ""), ensure_ascii=False) if actual else "",
                "notes": notes,
            }
        )

    write_results_xlsx(results)
    print(f"created {len(results)} test input files")
    print(f"inputs: {INPUT_DIR}")
    print(f"report: {RESULT_FILE}")


if __name__ == "__main__":
    main()
