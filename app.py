# -*- coding: utf-8 -*-
import hashlib
from datetime import datetime
from typing import List

import pandas as pd
import pytz
import streamlit as st

import gspread
from google.oauth2.service_account import Credentials
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Alignment, Font, Border, Side
from openpyxl.utils import get_column_letter

# =============== 設定 ===============
JST = pytz.timezone("Asia/Tokyo")

APP_SECRETS = st.secrets.get("app", {})
ALLOWED_DATES: List[str] = list(APP_SECRETS.get("allowed_dates", ["2025-10-25"]))
ALLOWED_PLACES: List[str] = list(APP_SECRETS.get("allowed_places", ["メインステージ"]))
DAY_START = APP_SECRETS.get("day_start", "09:00")
DAY_END = APP_SECRETS.get("day_end", "18:00")
GSHEET_ID = APP_SECRETS.get("gsheet_id", "")
ADMIN_PASSWORD = APP_SECRETS.get("admin_password", "")

# =============== Google Sheets 接続 ===============
@st.cache_resource(show_spinner=False)
def get_worksheet():
    creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"])
    scoped = creds.with_scopes([
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ])
    gc = gspread.authorize(scoped)

    try:
        sh = gc.open_by_key(GSHEET_ID)
    except Exception as ex:
        st.error(f"スプレッドシートを開けませんでした。GSHEET_ID={GSHEET_ID}, Error={ex}")
        raise

    try:
        ws = sh.worksheet("data")
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title="data", rows=1000, cols=12)
        ws.append_row([
            "timestamp", "group_name", "rep_name", "faculty", "email", "phone",
            "date", "place", "start", "end", "priority", "remarks"
        ])
    return ws

# =============== ユーティリティ ===============
def time_slots(day_start: str, day_end: str, step_min: int = 15) -> List[str]:
    base = pd.to_datetime(f"2000-01-01 {day_start}")
    end = pd.to_datetime(f"2000-01-01 {day_end}")
    return pd.date_range(base, end, freq=f"{step_min}min").strftime("%H:%M").tolist()

SLOTS = time_slots(DAY_START, DAY_END, 15)

def validate_range(start: str, end: str) -> bool:
    return pd.to_datetime(start) < pd.to_datetime(end)

def name_to_color(name: str) -> str:
    palette = [
        "#CFE8FF", "#FFD6A5", "#B9FBC0", "#FFADAD", "#FDFFB6", "#A0C4FF",
        "#CAFFBF", "#9BF6FF", "#F1C0E8", "#BDB2FF", "#FFC6FF", "#E0F7FA",
    ]
    h = int(hashlib.md5(name.encode("utf-8")).hexdigest(), 16)
    return palette[h % len(palette)]

def append_rows(ws, rows: list[list[str]]):
    ws.append_rows(rows, value_input_option="USER_ENTERED")

@st.cache_data(ttl=30)
def load_df() -> pd.DataFrame:
    """シートからデータを読み込む。空でも安全。"""
    ws = get_worksheet()
    values = ws.get_all_values()
    if not values:
        return pd.DataFrame(columns=[
            "timestamp", "group_name", "rep_name", "faculty", "email", "phone",
            "date", "place", "start", "end", "priority", "remarks"
        ])
    header, rows = values[0], values[1:]
    df = pd.DataFrame(rows, columns=header)
    # 型調整
    for c in ["date", "start", "end"]:
        if c in df.columns:
            df[c] = df[c].astype(str)
    if "priority" in df.columns:
        df["priority"] = pd.to_numeric(df["priority"], errors="coerce").astype("Int64")
    return df

def make_excel_by_date(df: pd.DataFrame, date_str: str) -> str:
    df_day = df[df["date"] == date_str].copy()
    if df_day.empty:
        raise ValueError("この日付のデータがありません。")

    wb = Workbook()
    wb.remove(wb.active)

    border_thin = Border(left=Side(style='thin'), right=Side(style='thin'),
                         top=Side(style='thin'), bottom=Side(style='thin'))

    for place in ALLOWED_PLACES:
        df_p = df_day[df_day["place"] == place].copy()
        ws = wb.create_sheet(title=place[:31])

        header = ["団体名"] + SLOTS
        ws.append(header)
        for col_idx in range(1, len(header) + 1):
            cell = ws.cell(row=1, column=col_idx)
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal="center", vertical="center")

        users = df_p["group_name"].dropna().unique().tolist()

        for r, user in enumerate(users, start=2):
            ws.cell(row=r, column=1, value=user)
            ws.cell(row=r, column=1).alignment = Alignment(vertical="center")

            sub = df_p[df_p["group_name"] == user]
            user_color = name_to_color(user)
            fill = PatternFill(start_color=user_color.replace('#',''),
                               end_color=user_color.replace('#',''),
                               fill_type="solid")

            for _, rec in sub.iterrows():
                start, end, pr = str(rec["start"]), str(rec["end"]), int(rec["priority"])
                try:
                    start = pd.to_datetime(start).strftime("%H:%M")
                    end = pd.to_datetime(end).strftime("%H:%M")
                except Exception:
                    continue
                if not validate_range(start, end):
                    continue
                try:
                    s_idx = SLOTS.index(start)
                    e_idx = SLOTS.index(end)
                except ValueError:
                    continue
                start_col = 2 + s_idx
                end_col_exclusive = 2 + e_idx
                for c in range(start_col, end_col_exclusive):
                    cell = ws.cell(row=r, column=c)
                    cell.fill = fill
                    label = f"第{pr}希望"
                    cell.value = f"{cell.value},{label}" if cell.value else label
                    cell.alignment = Alignment(horizontal="center", vertical="center")

        for row in ws.iter_rows(min_row=1, max_row=ws.max_row,
                                min_col=1, max_col=ws.max_column):
            for cell in row:
                cell.border = border_thin
        ws.column_dimensions['A'].width = 18
        for col in range(2, len(SLOTS) + 2):
            ws.column_dimensions[get_column_letter(col)].width = 4.2

    out_name = f"{date_str.replace('-', '')}.xlsx"
    wb.save(out_name)
    return out_name

# =============== UI ===============
st.set_page_config(page_title="施設利用希望フォーム", layout="wide")
st.title("大学祭発表団体募集フォーム")

ws = get_worksheet()

user_tab, admin_tab = st.tabs(["📝 利用者フォーム", "🛠 管理（一覧・Excel出力）"])

with user_tab:
    st.caption("※ 第1〜第3希望はすべて必須です。時間は15分刻みで選択してください。\n"
               "時間は準備・撤収も含めて設定してください。")

    group_name = st.text_input("団体名（必須）", key="group_name_input")

    st.markdown("#### 代表者情報")
    rep_name = st.text_input("代表者氏名（必須）", key="rep_name_input")
    faculty  = st.text_input("学部（必須）", key="faculty_input")
    email    = st.text_input("メールアドレス（必須）", key="email_input")
    phone    = st.text_input("電話番号（必須）", key="phone_input")

    remarks = st.text_area(
        "希望理由・備考（任意）",
        placeholder="希望理由や備考があれば入力してください",
        height=120,
        key="remarks_input"
    )

    def hope_block(title: str):
        st.subheader(title)
        c1, c2, c3, c4 = st.columns([1.2, 1.2, 1, 1])
        with c1:
            sel_date = st.selectbox("日付", ALLOWED_DATES, key=f"date_{title}")
        with c2:
            sel_place = st.selectbox("場所", ALLOWED_PLACES, key=f"place_{title}")
        with c3:
            start = st.selectbox("開始", SLOTS[:-1], key=f"start_{title}")
        with c4:
            end = st.selectbox("終了", SLOTS[1:], key=f"end_{title}")
        return sel_date, sel_place, start, end

    d1, p1, s1, e1 = hope_block("第1希望")
    d2, p2, s2, e2 = hope_block("第2希望")
    d3, p3, s3, e3 = hope_block("第3希望")

    if st.button("送信する", type="primary", key="submit_button"):
        errors = []
        if not group_name.strip():
            errors.append("団体名は必須です。")
        if not rep_name.strip():
            errors.append("代表者氏名は必須です。")
        if not faculty.strip():
            errors.append("学部は必須です。")
        if not email.strip():
            errors.append("メールアドレスは必須です。")
        if not phone.strip():
            errors.append("電話番号は必須です。")
        for idx, (s, e) in enumerate([(s1, e1), (s2, e2), (s3, e3)], start=1):
            if not validate_range(s, e):
                errors.append(f"第{idx}希望の時間範囲が不正です（開始 < 終了）。")

        if errors:
            st.error("\n".join(errors))
        else:
            ts = datetime.now(JST).strftime("%Y-%m-%d %H:%M:%S")
            rows = [
                [ts, group_name, rep_name, faculty, email, phone, d1, p1, s1, e1, 1, remarks],
                [ts, group_name, rep_name, faculty, email, phone, d2, p2, s2, e2, 2, remarks],
                [ts, group_name, rep_name, faculty, email, phone, d3, p3, s3, e3, 3, remarks],
            ]
            try:
                append_rows(ws, rows)
                st.success("送信しました。ご協力ありがとうございます！")
                load_df.clear()
            except Exception as ex:
                st.error(f"送信に失敗しました: {ex}")

# --- 管理タブ（パスワード保護） ---
with admin_tab:
    st.subheader("管理（一覧・Excel出力）")
    if "admin_auth" not in st.session_state:
        st.session_state["admin_auth"] = False
    if "admin_msg" not in st.session_state:
        st.session_state["admin_msg"] = ""

    if st.session_state["admin_auth"]:
        col_l, col_r = st.columns([1, 6])
        with col_l:
            if st.button("ログアウト", key="logout_button"):
                st.session_state["admin_auth"] = False
                st.session_state["admin_msg"] = "ログアウトしました。"
        with col_r:
            if st.session_state.get("admin_msg"):
                st.info(st.session_state["admin_msg"])

        df = load_df()
        st.subheader("データ一覧（最新）")
        st.dataframe(df, use_container_width=True)

        st.divider()
        st.subheader("Excel 出力（ガントチャート風）")
        selectable_dates = sorted(df["date"].dropna().unique().tolist()) if not df.empty else []
        target_dates = st.multiselect("作成する日付を選択", options=selectable_dates, default=selectable_dates, key="excel_dates")

        if st.button("選択した日付のExcelを作成", key="excel_button"):
            if not target_dates:
                st.warning("対象日付がありません。")
            else:
                for d in target_dates:
                    try:
                        path = make_excel_by_date(df, d)
                        with open(path, "rb") as f:
                            st.download_button(
                                label=f"{d} のExcelをダウンロード",
                                file_name=path,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                data=f.read(),
                                key=f"download_{d}"
                            )
                    except Exception as ex:
                        st.error(f"{d} の生成に失敗: {ex}")
    else:
        st.info("管理画面を表示するにはパスワードが必要です。")
        pwd = st.text_input("管理パスワード", type="password", key="admin_pwd_input")
        col1, col2 = st.columns([1, 1])
        with col1:
            if st.button("ログイン", key="login_button"):
                if not ADMIN_PASSWORD:
                    st.error("管理パスワードが未設定です。Secrets に app.admin_password を設定してください。")
                else:
                    if pwd == ADMIN_PASSWORD:
                        st.session_state["admin_auth"] = True
                        st.session_state["admin_msg"] = "認証に成功しました。"
                        st.rerun()
                    else:
                        st.error("パスワードが違います。")
        with col2:
            if st.button("キャンセル", key="cancel_button"):
                st.session_state["admin_msg"] = ""
                st.rerun()
