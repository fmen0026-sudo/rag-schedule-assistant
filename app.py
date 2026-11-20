import streamlit as st
from pathlib import Path
from tempfile import TemporaryDirectory
from main import find_schedule, project_dir, saved_file


st.title("课表助手")
st.caption("试试：2026-10-24 刑法学在哪上？")
uploaded_file = st.file_uploader(
    "上传课表（可选）",
    type=["xlsx"],
    help="请使用样例格式：工作表名为「课表」，前五列为星期、时间、课程、教室、日期。",
)
with st.form("schedule_query"):
    question = st.text_input("你想查什么？")
    submitted = st.form_submit_button("查询")

if submitted:
    question = question.strip()

    if not question:
        st.warning("请输入日期或星期，例如：周六有什么课？")
    else:
        if uploaded_file is None:
            results = find_schedule(
                question, saved_file, project_dir / "notices.txt"
            )
        else:
            with TemporaryDirectory() as temp_dir:
                upload_dir = Path(temp_dir)
                schedule_path = upload_dir / "schedule.xlsx"
                schedule_path.write_bytes(uploaded_file.getvalue())

                try:
                    results = find_schedule(
                        question, schedule_path, upload_dir / "notices.txt"
                    )
                except KeyError:
                    results = ["找不到名为「课表」的工作表，请检查 Excel 格式。"]

        st.subheader("查询结果")
        for line in results:
            st.text(line)