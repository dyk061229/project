import streamlit as st
import sqlite3
import pandas as pd
from datetime import date
from io import BytesIO

st.set_page_config(
    page_title="SQA Issue Management System",
    page_icon="📋",
    layout="wide"
)

DB_NAME = "sqa_issue.db"

FAIL_TYPES = [
    "Ghost touch",
    "Jitter",
    "Line broken",
    "Linecross",
    "터치 미인식",
    "기타",
    "참고사항"
]

def get_connection():
    return sqlite3.connect(DB_NAME, check_same_thread=False)

def create_table():
    conn = get_connection()
    cur = conn.cursor()
    cur.execute("""
        CREATE TABLE IF NOT EXISTS issues (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            date TEXT NOT NULL,
            panel_id TEXT NOT NULL,
            ic TEXT,
            model TEXT,
            build TEXT NOT NULL,
            test_item TEXT,
            fail_type TEXT NOT NULL,
            issue_detail TEXT
        )
    """)
    conn.commit()
    conn.close()

def insert_issue(panel_id, ic, model, build, test_item, fail_type, issue_detail):
    conn = get_connection()
    cur = conn.cursor()
    cur.execute("""
        INSERT INTO issues (
            date, panel_id, ic, model, build,
            test_item, fail_type, issue_detail
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        str(date.today()),
        panel_id,
        ic,
        model,
        build,
        test_item,
        fail_type,
        issue_detail
    ))
    conn.commit()
    conn.close()

def load_issues():
    conn = get_connection()
    df = pd.read_sql_query("SELECT * FROM issues ORDER BY id DESC", conn)
    conn.close()
    return df

def update_issue(issue_id, panel_id, ic, model, build, test_item, fail_type, issue_detail):
    conn = get_connection()
    cur = conn.cursor()

    cur.execute("""
        UPDATE issues
        SET panel_id = ?,
            ic = ?,
            model = ?,
            build = ?,
            test_item = ?,
            fail_type = ?,
            issue_detail = ?
        WHERE id = ?
    """, (
        panel_id,
        ic,
        model,
        build,
        test_item,
        fail_type,
        issue_detail,
        issue_id
    ))

    conn.commit()
    conn.close()


def delete_issue(issue_id):
    conn = get_connection()
    cur = conn.cursor()

    cur.execute("DELETE FROM issues WHERE id = ?", (issue_id,))

    conn.commit()
    conn.close()


def to_excel(df):
    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Issue List")

    return output.getvalue()

create_table()

st.title("📋 SQA Issue Management System")
st.caption("Panel ID / FW Version 기반 테스트 이력 및 Fail Issue 등록·조회 시스템")

menu = st.radio(
    "메뉴 선택",
    ["이슈 등록", "이슈 조회"],
    horizontal=True
)

if menu == "이슈 등록":
    st.subheader("이슈 등록")
        
        col1, col2, col3 = st.columns([1,1,1])
        
        with col1:
            panel_id = st.text_input("Panel ID / Code *")
        
        with col2:
            ic = st.text_input("IC")
        
        with col3:
            model = st.text_input("Model")
        
        col4, col5 = st.columns([1,1])
        
        with col4:
            build = st.text_input("FW Version *")
        
        with col5:
            test_item = st.selectbox(
                "Test Item",
                ["Drawing", "Ghost", "Jitter", "Palm", "Line Broken"]
            )
        
        fail_type = st.multiselect(
            "Fail Type *",
            FAIL_TYPES
        )
        
        issue_detail = st.text_area(
            "상세 이슈 내용",
            height=150
        )

        submitted = st.form_submit_button("이슈 등록")

        if submitted:
            if not panel_id or not build or not fail_type_list:
                st.warning("Panel ID / FW Version / Fail Type은 필수 입력 항목입니다.")
            else:
                fail_type = ", ".join(fail_type_list)
                insert_issue(panel_id, ic, model, build, test_item, fail_type, issue_detail)
                st.success("이슈 등록이 완료되었습니다.")


if menu == "이슈 조회":
    st.subheader("이슈 조회")

    df = load_issues()

    col1, col2, col3 = st.columns(3)

    with col1:
        panel_search = st.text_input("Panel ID / Code")

    with col2:
        ic_search = st.text_input("IC")

    with col3:
        model_search = st.text_input("Model")

    col4, col5 = st.columns(2)

    with col4:
        fw_search = st.text_input("FW Version")

    with col5:
        test_item_search = st.selectbox(
            "Test Item",
            ["전체", "Drawing", "Jitter", "Ghost", "Line Broken", "Palm", "Edge", "Multi", "WHLK", "CS", "ODM", "ETC"]
        )

    fail_filter = st.multiselect(
        "Fail Type",
        FAIL_TYPES,
        placeholder="Fail Type을 선택해주세요. 중복 선택 가능"
    )

    if st.button("조회"):
        st.session_state["search_clicked"] = True

    search_clicked = st.session_state.get("search_clicked", False)



    if df.empty:
        st.info("등록된 이슈가 없습니다.")

    elif search_clicked:
        filtered_df = df.copy()

        if panel_search:
            filtered_df = filtered_df[
                filtered_df["panel_id"].str.contains(panel_search, case=False, na=False)
            ]

        if ic_search:
            filtered_df = filtered_df[
                filtered_df["ic"].str.contains(ic_search, case=False, na=False)
            ]

        if model_search:
            filtered_df = filtered_df[
                filtered_df["model"].str.contains(model_search, case=False, na=False)
            ]

        if fw_search:
            filtered_df = filtered_df[
                filtered_df["build"].str.contains(fw_search, case=False, na=False)
            ]

        if test_item_search != "전체":
            filtered_df = filtered_df[
                filtered_df["test_item"] == test_item_search
            ]

        if fail_filter:
            pattern = "|".join(fail_filter)
            filtered_df = filtered_df[
                filtered_df["fail_type"].str.contains(pattern, case=False, na=False)
            ]


        st.metric("조회 결과", len(filtered_df))

        display_df = filtered_df.rename(columns={
            "id": "ID",
            "date": "Date",
            "panel_id": "Panel ID",
            "ic": "IC",
            "model": "Model",
            "build": "FW Version",
            "test_item": "Test Item",
            "fail_type": "Fail Type",
            "issue_detail": "Issue Detail"
        })

        display_df.insert(0, "No", range(1, len(display_df) + 1))

        display_df["Issue Preview"] = (
            display_df["Issue Detail"]
            .astype(str)
            .str.slice(0, 30)
            + "..."
        )

        editable_df = display_df.copy()
        editable_df.insert(0, "선택", False)

        edited_df = st.data_editor(
            editable_df,
            use_container_width=True,
            hide_index=True,
            num_rows="fixed",
            disabled=["ID", "No", "Date"],
            column_config={
                "ID": None,
                "Issue Detail": None,
                "Issue Preview": st.column_config.TextColumn("Issue Detail", width="large")
            },
            key="issue_editor"
        )

        col_save, col_delete = st.columns(2)

        with col_save:
            if st.button("수정 내용 저장"):
                save_df = edited_df.drop(columns=["선택"])

                for _, row in save_df.iterrows():
                    update_issue(
                        row["ID"],
                        row["Panel ID"],
                        "" if pd.isna(row["IC"]) else row["IC"],
                        "" if pd.isna(row["Model"]) else row["Model"],
                        row["FW Version"],
                        row["Test Item"],
                        row["Fail Type"],
                        "" if pd.isna(row["Issue Detail"]) else row["Issue Detail"]
                    )

                st.success("수정 내용이 저장되었습니다.")
                st.rerun()

            with col_delete:
                selected_rows = edited_df[edited_df["선택"] == True]

            if st.button("선택한 이슈 삭제"):
                if selected_rows.empty:
                    st.warning("삭제할 이슈를 체크해주세요.")
                else:
                    st.session_state["delete_confirm"] = True

            if st.session_state.get("delete_confirm", False):
                st.warning("정말 선택한 이슈를 삭제하시겠습니까? 삭제 후에는 되돌릴 수 없습니다.")

                confirm_col, cancel_col = st.columns(2)

                with confirm_col:
                    if st.button("예, 삭제합니다"):
                        for _, row in selected_rows.iterrows():
                            delete_issue(row["ID"])

                        st.session_state["delete_confirm"] = False
                        st.success("선택한 이슈가 삭제되었습니다.")
                        st.rerun()

                with cancel_col:
                    if st.button("취소"):
                        st.session_state["delete_confirm"] = False
                        st.info("삭제를 취소했습니다.")
                        st.rerun()

        
        excel_df = filtered_df.rename(columns={
            "date": "Date",
            "panel_id": "Panel ID",
            "ic": "IC",
            "model": "Model",
            "build": "FW Version",
            "test_item": "Test Item",
            "fail_type": "Fail Type",
            "issue_detail": "Issue Detail"
        })

        excel_df.drop(columns=["id"], inplace=True, errors="ignore")
        excel_df.insert(0, "No", range(1, len(excel_df) + 1))

        excel_data = to_excel(excel_df)

        st.download_button(
            "전체 이슈 Excel 다운로드",
            excel_data,
            file_name=f"Issue_Report_{date.today()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    else:
        st.caption("조회 조건을 입력한 뒤 [조회] 버튼을 눌러주세요.")
