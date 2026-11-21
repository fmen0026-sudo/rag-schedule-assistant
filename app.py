import os
import re
from datetime import date
from hashlib import sha256
from io import BytesIO
from pathlib import Path
from tempfile import TemporaryDirectory
from zipfile import BadZipFile

import pdfplumber
import streamlit as st
from openpyxl.utils.exceptions import InvalidFileException
from streamlit.errors import StreamlitSecretNotFoundError

from main import find_schedule, project_dir, saved_file
from schedule_ai import ask_schedule
from schedule_import import (
    ImportedSchedule, expand_records, recognize_schedule, workbook_bytes,
)


def get_deepseek_key():
    """Read the API key from the environment or Streamlit Secrets."""
    key = os.getenv("DEEPSEEK_API_KEY", "").strip()
    if key:
        return key

    try:
        return st.secrets.get("DEEPSEEK_API_KEY", "").strip()
    except StreamlitSecretNotFoundError:
        return ""


def run_query(question, schedule_path, notice_path, use_ai, first_monday):
    """执行一次课表查询，把结果交给界面渲染。"""
    if not use_ai:
        lines = find_schedule(question, schedule_path, notice_path, first_monday)
        return {"lines": lines}

    api_key = get_deepseek_key()
    if not api_key:
        return {"notice": "尚未配置 DeepSeek 密钥。关闭 AI 后可以按具体日期查询。"}

    with st.spinner("正在理解问题并查询课表…"):
        answer, evidence, notice = ask_schedule(
            question, schedule_path, notice_path, api_key, first_monday
        )
    return {"answer": answer, "evidence": evidence, "notice": notice}


def render_evidence(results, *, expanded):
    """按查询类型展示原始记录，并为统计结果显示可核对的摘要。"""
    with st.expander("查看查询依据", expanded=expanded):
        for result in results:
            start = result.get("start_date")
            end = result.get("end_date")
            scope = f"{start} 至 {end}" if start and end != start else start or "整份课表"
            st.caption(f"查询范围：{scope}")

            records = result.get("records", [])
            days = result.get("days", [])
            if records:
                rows = [
                    {"日期": row["date"], "星期": row["weekday"], "节次": row["time"],
                     "课程": row["course"], "教室": row["classroom"]}
                    for row in records
                ]
                st.dataframe(rows, hide_index=True)
                notices = dict.fromkeys(row["notice"] for row in records if row["notice"])
                for notice in notices:
                    st.text(f"调课依据：{notice}")
            elif days:
                busiest = result.get("busiest_days", [])
                if busiest:
                    st.caption(
                        f"单日最高 {busiest[0]['class_count']} 条上课记录，"
                        f"共 {len(busiest)} 天达到最高值。"
                    )
                st.dataframe(
                    [{"日期": day["date"], "上课记录数": day["class_count"]} for day in days],
                    hide_index=True,
                )
            else:
                st.info("这份课表未找到匹配记录。")

            if result.get("note"):
                st.caption(result["note"])


def render_query_result(result):
    """显示一次已完成查询。重新渲染时无需再次调用模型。"""
    with st.container(border=True, key="schedule-result"):
        st.subheader("查询结果")
        st.text(result["question"])
        for line in result.get("lines", []):
            st.text(line)
        answer = result.get("answer")
        if answer:
            st.markdown(answer)
        if result.get("notice"):
            st.info(result["notice"])
        if result.get("error"):
            st.warning(result["error"])
        if result.get("evidence"):
            render_evidence(result["evidence"], expanded=not answer)


# A verified calendar for the original example. Other schools can supply their own.
FIRST_MONDAYS = {("2026-2027", 1): date(2026, 9, 7)}


def known_pdf_calendar(data):
    """Look up the known calendar only when the PDF contains matching campus markers."""
    try:
        with pdfplumber.open(BytesIO(data)) as pdf:
            cover_text = pdf.pages[0].extract_text() if pdf.pages else ""
    except Exception:
        # This is an optional lookup. Files without extractable text still use AI import.
        return None

    text = re.sub(r"\s+", "", cover_text or "")
    term = re.search(r"(\d{4})[-－–—](\d{4})学年第([12])学期", text)
    if term and "旗山北校区" in text and "博学" in text:
        return FIRST_MONDAYS.get((f"{term[1]}-{term[2]}", int(term[3])))
    return None


def prepare_uploaded_schedule(uploaded_file):
    """Recognize each uploaded file once per session and show inline import status."""
    if uploaded_file is None:
        st.session_state.pop("schedule_import", None)
        return None, None

    data = uploaded_file.getvalue()
    extension = Path(uploaded_file.name).suffix.lower()
    identity = (extension, sha256(data).hexdigest())
    state = st.session_state.get("schedule_import")
    if state is None or state["identity"] != identity:
        state = {"identity": identity}
        st.session_state["schedule_import"] = state

    api_key = get_deepseek_key()
    if not api_key:
        st.warning("AI 导入需要配置 DeepSeek 密钥。移除文件后仍可查询演示课表。")
        return None, None

    retry = bool(state.get("error")) and st.button("重试课表识别")
    if "schedule" not in state and ("error" not in state or retry):
        with st.spinner("DeepSeek 正在识别课表…"):
            try:
                state["schedule"] = recognize_schedule(uploaded_file.name, data, api_key)
                state.pop("error", None)
            except ValueError as error:
                state["error"] = str(error)

    if "error" in state:
        st.warning(state["error"])
        return None, None

    schedule: ImportedSchedule = state["schedule"]
    if schedule.warnings:
        with st.expander("查看识别提示", expanded=True):
            for message in schedule.warnings:
                st.text(message)

    if "calendar_hint" not in state:
        state["calendar_hint"] = known_pdf_calendar(data) if extension == ".pdf" else None
    first_monday = state["calendar_hint"] or schedule.first_monday
    if schedule.requires_calendar:
        first_monday = st.date_input(
            "第 1 教学周的周一",
            value=first_monday,
            key=f"calendar-{identity[1]}",
            help="课表含教学周。请按学校校历核对，日期变化只重新换算，不会再次调用 DeepSeek。",
        )
        if first_monday is None:
            st.info("已识别课程和周次。请补充第 1 教学周的周一日期，即可查询。")
            return None, None

    try:
        if "workbook" not in state or state.get("converted_with") != first_monday:
            records = expand_records(schedule, first_monday)
            state["workbook"] = workbook_bytes(records)
            state["records"] = records
            state["converted_with"] = first_monday
    except (ValueError, OverflowError) as error:
        st.warning(f"课表日期无法换算：{error}")
        return None, None

    st.success(f"已识别 {len(state['records'])} 条上课记录。")
    with st.expander("查看识别结果"):
        st.dataframe(
            [dict(zip(("星期", "时间", "课程", "教室", "日期"), row)) for row in state["records"]],
            hide_index=True,
        )
    return state["workbook"], first_monday


def render_header():
    """显示品牌标题以及应用的主要用途。"""
    st.html("""
    <style>
      .schedule-hero {
        box-sizing: border-box;
        border-radius: 20px;
        padding: clamp(1.4rem, 4vw, 2.65rem);
        color: #fff;
        background: linear-gradient(115deg, #152944 18%, #244d86 100%);
        margin-bottom: .4rem;
      }
      .schedule-hero h1 {
        color: #fff;
        font-size: clamp(1.55rem, 4vw, 2.5rem);
        line-height: 1.25;
        letter-spacing: .01em;
        margin: 0;
        padding: 0;
      }
      .schedule-hero p {
        color: #e1ebff;
        font-size: .97rem;
        line-height: 1.65;
        margin: .65rem 0 0;
      }
      .st-key-schedule-result { background: #fff; }
    </style>
    <header class="schedule-hero">
      <h1>智能课表问答与分析助手</h1>
      <p>上传课表，用一句话查询课程、教室和每天的排课情况。</p>
    </header>
    """)


def run_uploaded_query(question, workbook, first_monday, use_ai):
    """把当前会话的课表交给已有查询层，并在结束时清理临时文件。"""
    with TemporaryDirectory() as temp_dir:
        upload_dir = Path(temp_dir)
        schedule_path = upload_dir / "schedule.xlsx"
        schedule_path.write_bytes(workbook)
        return run_query(
            question, schedule_path, upload_dir / "notices.txt", use_ai, first_monday
        )


st.set_page_config(page_title="智能课表问答与分析助手", page_icon="📅", layout="centered")
render_header()

with st.container(border=True, key="schedule-source"):
    st.subheader("课表来源")
    uploaded_file = st.file_uploader(
        "上传课表 PDF 或 Excel（可选）",
        type=["xlsx", "pdf"],
        help="上传会将文件发送给 DeepSeek 云端识别。请先去除姓名和学号。",
    )
    workbook, first_monday = prepare_uploaded_schedule(uploaded_file)

# Associate the displayed answer with the same file and calendar used to compute it.
import_state = st.session_state.get("schedule_import")
source_id = (import_state["identity"], first_monday) if import_state else "demo"
saved_query = st.session_state.get("last_schedule_query")
if saved_query and (saved_query["source_id"] != source_id or (uploaded_file and workbook is None)):
    st.session_state.pop("last_schedule_query", None)

with st.form("schedule_query"):
    st.subheader("问课表")
    example = "哪天课最多？" if uploaded_file is not None else "2026-10-24 刑法学在哪上？"
    question = st.text_input("你想查什么？", placeholder=example)
    use_ai = st.checkbox(
        "使用 AI 理解问题并生成回答",
        value=True,
        help="关闭后可按具体日期查询。",
    )
    submitted = st.form_submit_button(
        "查询", type="primary", disabled=uploaded_file is not None and workbook is None,
        width="stretch",
    )

if submitted:
    question = question.strip()
    if not question:
        st.warning("请输入你想查的课表问题。")
    else:
        try:
            if uploaded_file is None:
                result = run_query(question, saved_file, project_dir / "notices.txt", use_ai, None)
            else:
                result = run_uploaded_query(question, workbook, first_monday, use_ai)
        except ValueError as error:
            result = {"error": str(error)}
        except KeyError:
            result = {"error": "找不到名为「课表」的工作表，请检查 Excel 格式。"}
        except (BadZipFile, InvalidFileException):
            result = {"error": "无法读取课表数据，请重新上传文件。"}
        st.session_state["last_schedule_query"] = {
            **result, "question": question, "source_id": source_id,
        }

if result := st.session_state.get("last_schedule_query"):
    render_query_result(result)
