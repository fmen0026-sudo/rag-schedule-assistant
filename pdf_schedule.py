import re
from datetime import timedelta
from io import BytesIO

import pdfplumber
from openpyxl import Workbook
from pdfminer.pdfparser import PDFSyntaxError


DAYS = ("周一", "周二", "周三", "周四", "周五", "周六", "周日")

# 匹配一个格子里的单段安排。一个格子可以有多段。
COURSE_PATTERN = re.compile(
    r"(?P<course>[^/]+?)\s*"
    r"\((?P<slot>\d+(?:-\d+)?)节\)\s*"
    r"(?P<weeks>[^/]+)/校\s*区\s*[:：][^/]+"
    r"/场\s*地\s*[:：](?P<room>[^/]+)"
    r"/教\s*师\s*[:：][^\n/]+"
    r"(?:\n(?![^\n]+\n?\s*\(\d+(?:-\d+)?节\))[^\n/]+)?"
)
SLOT_PATTERN = re.compile(r"\(\d+(?:-\d+)?节\)")


def expand_weeks(week_text):
    """把“1-3周(单),6周”展开成第 1、3、6 周。"""
    compact = re.sub(r"\s+", "", week_text).replace("，", ",")
    weeks = set()

    for part in compact.split(","):
        match = re.fullmatch(
            r"(\d+)(?:-(\d+))?周(?:[（(]([单双])[）)])?",
            part,
        )
        if not match:
            raise ValueError(f"无法识别教学周写法：{part}")

        start = int(match[1])
        end = int(match[2] or match[1])
        parity = match[3]

        if not 1 <= start <= end <= 53:
            raise ValueError(f"教学周范围无效：{part}")

        for week in range(start, end + 1):
            if parity is None or week % 2 == (1 if parity == "单" else 0):
                weeks.add(week)

    return sorted(weeks)


def convert_pdf_to_xlsx(pdf_bytes, first_monday, output_path):
    """读取教务系统 PDF，把每次上课写成 Excel 的一行。"""
    if first_monday is None or first_monday.weekday() != 0:
        raise ValueError("请填写学校第 1 教学周的周一日期。")

    records = set()

    try:
        with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
            for page in pdf.pages:
                for table in page.extract_tables():
                    for row in table:
                        # 前两列是时间段和节次，后七列是星期一到星期日。
                        if len(row) != 9 or not (row[1] or "").strip().isdigit():
                            continue

                        for day_index, cell in enumerate(row[2:]):
                            if not cell or not cell.strip():
                                continue

                            matches = list(COURSE_PATTERN.finditer(cell))
                            if not matches or len(matches) != len(SLOT_PATTERN.findall(cell)):
                                raise ValueError("有课程格子无法完整识别，请检查 PDF 格式。")

                            for match in matches:
                                course = match["course"].replace("\n", "").strip()
                                room = re.sub(r"\s+", "", match["room"])
                                slot = f"第{match['slot']}节"

                                for week in expand_weeks(match["weeks"]):
                                    day = first_monday + timedelta(
                                        weeks=week - 1, days=day_index
                                    )
                                    records.add((
                                        day.isoformat(),
                                        DAYS[day_index],
                                        slot,
                                        course,
                                        room,
                                    ))
    except PDFSyntaxError as error:
        raise ValueError("无法打开这个 PDF，请检查文件是否完整。") from error

    if not records:
        raise ValueError("没有识别到课程，请上传同版式的教务系统课表。")

    with_workbook = Workbook()
    try:
        sheet = with_workbook.active
        sheet.title = "课表"
        sheet.append(["星期", "时间", "课程", "教室", "日期"])

        # 按日期和起始节次排序。
        for day, weekday, slot, course, room in sorted(
            records, key=lambda row: (row[0], int(re.search(r"\d+", row[2])[0]))
        ):
            sheet.append([weekday, slot, course, room, day])

        with_workbook.save(output_path)
    finally:
        with_workbook.close()