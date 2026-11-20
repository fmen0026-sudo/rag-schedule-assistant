
print("课表助手启动成功")
import re
from pathlib import Path
from shutil import copyfile

from openpyxl import load_workbook
from datetime import date as CalendarDate, datetime, timedelta, timezone

# 以 main.py 所在的文件夹为准，定位原始课表和导入后的副本。
project_dir = Path(__file__).resolve().parent
source_file = project_dir / "schedule.xlsx"
saved_file = project_dir / "saved_schedule.xlsx"


def import_schedule():
    if not source_file.is_file():
        print("找不到 schedule.xlsx，请把它放到 main.py 所在的文件夹。")
        return

    copyfile(source_file, saved_file)
    print("课表已导入。下次运行可以直接查询。")

DATE_QUERY_HINT = "请指定日期或星期，例如 10月12号、下周一或周六。"


def load_schedule(schedule_path, notice_path):
    """Read course rows and attach any matching classroom-change notice."""
    notices = notice_path.read_text(encoding="utf-8-sig").splitlines() if notice_path.is_file() else []
    workbook = load_workbook(schedule_path, read_only=True)
    records = []
    try:
        sheet = workbook["课表"]
        for weekday, time, course, classroom, day in sheet.iter_rows(
            min_row=2, max_col=5, values_only=True
        ):
            if not course:
                continue
            date_text = str(day or "")[:10]
            slot = str(time or "未注明时段")
            room = str(classroom or "未注明教室")
            period = re.fullmatch(r"(?:第)?\s*(\d+)\s*(?:[-–—至]\s*(\d+))?\s*节", slot)
            prefix = f"{date_text} {course}临时调至 "
            notice = next((line for line in notices if line.startswith(prefix)), "")
            records.append({
                "date": date_text,
                "weekday": str(weekday or ""),
                "time": slot,
                "course": str(course),
                "classroom": notice[len(prefix):].split("，", 1)[0].strip() if notice else room,
                "original_classroom": room if notice else "",
                "notice": notice,
                "first_period": int(period[1]) if period else None,
                "last_period": int(period[2] or period[1]) if period else None,
            })
    finally:
        workbook.close()
    return records


def find_schedule(
    question,
    schedule_path,
    notice_path,
    first_monday=None,
    *,
    date_range=None,
    course_filter="",
):
    """查找课程和调课通知，返回可以显示的文字列表。"""
    if not schedule_path.is_file():
        return ["请先导入课表。"]

    # 1. 把问题转换成具体日期或星期。
    # 公共网页部署后，仍按学校所在的北京时间计算“今天”。
    today = datetime.now(timezone(timedelta(hours=8))).date()

    date_match = re.search(r"(?<!\d)\d{4}-\d{2}-\d{2}(?!\d)", question)
    month_day = re.search(r"(\d{1,2})月\s*(\d{1,2})(?:日|号)", question)
    next_week = re.search(r"下(?:个)?(?:周|星期)([一二三四五六日天])", question)

    weekdays = ("周一", "周二", "周三", "周四", "周五", "周六", "周日")
    matched_days = [day for day in weekdays if day in question]

    if date_range is not None:
        start_date, end_date = date_range
        target = f"{start_date} 至 {end_date}"
    elif date_match:
        target = date_match.group()

    elif month_day:
        month, day = map(int, month_day.groups())
        anchor = first_monday or today
        candidates = []

        for year in (anchor.year - 1, anchor.year, anchor.year + 1):
            try:
                candidates.append(CalendarDate(year, month, day))
            except ValueError:
                continue

        if not candidates:
            return ["这个月日无效，请检查后再试。"]

        # 用离学期起始日最近的年份，兼顾跨年的秋季学期。
        target = min(candidates, key=lambda day: abs(day - anchor)).isoformat()

    elif next_week:
        weekday_index = "一二三四五六日".index(next_week[1].replace("天", "日"))
        next_monday = today + timedelta(days=7 - today.weekday())
        target = (next_monday + timedelta(days=weekday_index)).isoformat()

    elif "明天" in question:
        target = (today + timedelta(days=1)).isoformat()

    elif re.search(r"第\s*[\d一二三四五六七八九十百两零〇]+\s*(?:教学)?周", question):
        return ["按教学周查询请启用 AI，或输入具体日期。"]

    elif len(matched_days) == 1:
        target = matched_days[0]


    else:
        return [DATE_QUERY_HINT]

    results = []
    for record in load_schedule(schedule_path, notice_path):
        day = record["date"]
        if date_range is not None:
            if not start_date <= day <= end_date:
                continue
        elif target not in (day, record["weekday"]):
            continue
        if course_filter and course_filter not in record["course"]:
            continue
        room = record["classroom"]
        if record["notice"]:
            room += f"（原教室 {record['original_classroom']}）"
        results.append(f"{day}  {record['time']}  {record['course']}  {room}")
        if record["notice"]:
            results.append(f"调课依据：{record['notice']}")
    return results or [f"课表里没有找到 {target} 的课程。"]


def query_schedule():
    """从终端接收问题，并在终端显示查询结果。"""
    question = input("请输入问题，例如：2026-10-24 有什么课？").strip()
    results = find_schedule(
        question,
        saved_file,
        project_dir / "notices.txt",
    )

    for line in results:
        print(line)


def main():
    print("1. 导入课表")
    print("2. 查询课表")
    choice = input("请选择 1 或 2：").strip()

    if choice == "1":
        import_schedule()
    elif choice == "2":
        query_schedule()
    else:
        print("请输入 1 或 2。")


if __name__ == "__main__":
    main()