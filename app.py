import os
import re
from datetime import date
from io import BytesIO
from pathlib import Path
from tempfile import TemporaryDirectory
from zipfile import BadZipFile

import pdfplumber
import streamlit as st
from openpyxl.utils.exceptions import InvalidFileException
from streamlit.errors import StreamlitSecretNotFoundError

from main import find_schedule, project_dir, saved_file
from pdf_schedule import convert_pdf_to_xlsx
from schedule_ai import ask_schedule


def get_deepseek_key():
    key = os.getenv("DEEPSEEK_API_KEY", "").strip()
    if key:
        return key

    try:
        return st.secrets.get("DEEPSEEK_API_KEY", "").strip()
    except StreamlitSecretNotFoundError:
        return ""


def render_query(question, schedule_path, notice_path, use_ai, first_monday):
    if not use_ai:
        st.subheader("查询结果")
        for line in find_schedule(question, schedule_path, notice_path, first_monday):
            st.text(line)
        return

    api_key = get_deepseek_key()
    if not api_key:
        st.info("尚未配置 DeepSeek 密钥。关闭 AI 后可以按具体日期查询。")
        return
    with st.spinner("正在理解问题并查询课表…"):
        answer, evidence, notice = ask_schedule(
            question, schedule_path, notice_path, api_key, first_monday
        )
    if answer:
        st.subheader("AI 回答")
        st.markdown(answer)
    if notice:
        st.info(notice)
    if evidence:
        with st.expander("查看查询依据", expanded=not answer):
            for result in evidence:
                if result["start_date"]:
                    st.caption(f'{result["start_date"]} 至 {result["end_date"]}')
                if "note" in result:
                    st.caption(result["note"])
                rows = result.get("records", result.get("days", []))
                if rows:
                    st.dataframe(rows, hide_index=True)
                else:
                    st.info("这份课表未找到匹配记录。")


st.title("课表助手")
st.caption("演示课表，调课通知为模拟数据。试试：2026-10-24 刑法学在哪上？")
uploaded_file = st.file_uploader(
    "上传课表（可选）",
    type=["xlsx", "pdf"],
    help="Excel 请使用样例格式。PDF 支持带表格线、可提取文字的同版式教务课表。",
)
first_monday = None
if uploaded_file is not None and uploaded_file.name.lower().endswith(".pdf"):
    # 福建理工大学 2026–2027 学年第一学期校历。
    FIRST_MONDAYS = {
        ("2026-2027", 1): date(2026, 9, 7),
    }

    first_monday = None

    if uploaded_file is not None and uploaded_file.name.lower().endswith(".pdf"):
        try:
            with pdfplumber.open(BytesIO(uploaded_file.getvalue())) as pdf:
                cover_text = pdf.pages[0].extract_text() if pdf.pages else ""

            term = re.search(
                r"(\d{4})\s*[-－–—]\s*(\d{4})\s*学年第\s*([12])\s*学期",
                cover_text or "",
            )

            campus_text = re.sub(r"\s+", "", cover_text or "")
            if term and "旗山北校区" in campus_text and "博学" in campus_text:
                school_year = f"{term[1]}-{term[2]}"
                first_monday = FIRST_MONDAYS.get((school_year, int(term[3])))
        except Exception:
            # 上传内容无法读取时，让提交查询时的 PDF 解析给出提示。
            first_monday = None

        if first_monday is None:
            first_monday = st.date_input(
                "学校第 1 教学周的周一",
                value=None,
                help="这份 PDF 未匹配到已配置的校历，请按学校校历填写。",
            )
with st.form("schedule_query"):
    question = st.text_input("你想查什么？")
    use_ai = st.checkbox(
        "使用 AI 理解问题并生成回答",
        help=(
            "DeepSeek 会理解问题并选择查询方式；查到的相关课表记录和通知会用于生成回答。"
        ),
    )
    submitted = st.form_submit_button("查询")

if submitted:
    question = question.strip()
    if not question:
        st.warning("请输入你想查的课表问题。")
    else:
        try:
            if uploaded_file is None:
                render_query(question, saved_file, project_dir / "notices.txt", use_ai, first_monday)
            else:
                with TemporaryDirectory() as temp_dir:
                    upload_dir = Path(temp_dir)
                    schedule_path = upload_dir / "schedule.xlsx"
                    if uploaded_file.name.lower().endswith(".pdf"):
                        convert_pdf_to_xlsx(uploaded_file.getvalue(), first_monday, schedule_path)
                    else:
                        schedule_path.write_bytes(uploaded_file.getvalue())
                    render_query(
                        question, schedule_path, upload_dir / "notices.txt", use_ai, first_monday
                    )
        except ValueError as error:
            st.warning(str(error))
        except KeyError:
            st.warning("找不到名为「课表」的工作表，请检查 Excel 格式。")
        except (BadZipFile, InvalidFileException):
            st.warning("无法读取 Excel 文件，请上传有效的 XLSX 课表。")
