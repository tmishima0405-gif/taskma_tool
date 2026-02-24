import re
import zipfile
import secrets
from datetime import date, timedelta
from datetime import datetime, timezone
from io import BytesIO
from pathlib import Path
from xml.etree import ElementTree as ET

import pandas as pd
import streamlit as st
import qrcode


COLUMNS = ["DueDate", "Schedule", "Section", "Project", "Tag", "TaskName", "Estimated"]
NS_MAIN = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
NS_RELS = {"r": "http://schemas.openxmlformats.org/package/2006/relationships"}
FALLBACK_SECTIONS = [
    "00:00",
    "05:00",
    "07:00",
    "08:00",
    "10:00",
    "11:00",
    "13:00",
    "14:00",
    "16:00",
    "17:00",
    "19:00",
    "21:00",
    "22:00",
]
FALLBACK_PROJECTS = [
    ".重要",
    ".情報",
    ".雑務",
    ".家族",
    ".家のこと",
    ".整理",
    ".勉強",
    ".他",
    ".娯楽",
    ".オペ-90 臨時",
    ".月次作業",
]
FALLBACK_TAGS = ["", "スキップ", "急ぎ", "自分", "会議中", "ミーティング"]
DEFAULT_PUBLIC_APP_URL = "https://streamlitcsv.streamlit.app"
TEMP_DOWNLOADS = {}


def resolve_xlsm_path() -> Path | None:
    base = Path(__file__).parent
    preferred = base / "たすくま.xlsm"
    if preferred.exists():
        return preferred
    candidates = sorted(base.glob("*.xlsm"))
    return candidates[0] if candidates else None


def excel_time_to_hhmm(value: str) -> str:
    minutes = int(round(float(value) * 24 * 60))
    hours, mins = divmod(minutes, 60)
    return f"{hours % 24:02d}:{mins:02d}"


def cell_text(cell, shared_strings):
    value = cell.find("m:v", NS_MAIN)
    if value is None or value.text is None:
        return None
    if cell.attrib.get("t") == "s":
        try:
            return shared_strings[int(value.text)]
        except (ValueError, IndexError):
            return None
    return value.text


def add_unique(items: list[str], value):
    if value is None:
        return
    text = str(value).strip()
    if text and text not in items:
        items.append(text)


def read_dropdown_options(xlsm_path: Path | None):
    sections, projects, tags = [], [], []
    if xlsm_path is None or not xlsm_path.exists():
        return sections, projects, tags

    with zipfile.ZipFile(xlsm_path) as archive:
        shared_strings = []
        if "xl/sharedStrings.xml" in archive.namelist():
            ss_root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
            for si in ss_root.findall("m:si", NS_MAIN):
                text = "".join((t.text or "") for t in si.findall(".//m:t", NS_MAIN))
                shared_strings.append(text)

        workbook = ET.fromstring(archive.read("xl/workbook.xml"))
        rels = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        relationship_map = {
            rel.attrib["Id"]: rel.attrib["Target"]
            for rel in rels.findall("r:Relationship", NS_RELS)
        }

        list_sheet_target = None
        for sheet in workbook.findall("m:sheets/m:sheet", NS_MAIN):
            name = sheet.attrib.get("name", "")
            rid = sheet.attrib.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
            target = relationship_map.get(rid, "")
            if name == "リスト":
                list_sheet_target = f"xl/{target}"
                break

        if list_sheet_target is None:
            fallback = "xl/worksheets/sheet3.xml"
            if fallback in archive.namelist():
                list_sheet_target = fallback
            else:
                return sections, projects, tags

        list_sheet = ET.fromstring(archive.read(list_sheet_target))
        ref_pattern = re.compile(r"^([A-Z]+)(\d+)$")

        for cell in list_sheet.findall(".//m:sheetData//m:c", NS_MAIN):
            ref = cell.attrib.get("r", "")
            match = ref_pattern.match(ref)
            if not match:
                continue
            col, row_text = match.groups()
            row = int(row_text)
            if row < 4:
                continue

            raw = cell_text(cell, shared_strings)
            if col == "C" and raw is not None:
                try:
                    add_unique(sections, excel_time_to_hhmm(raw))
                except ValueError:
                    add_unique(sections, raw)
            elif col == "D":
                add_unique(projects, raw)
            elif col == "E":
                add_unique(tags, raw)

    return sections, projects, tags


def merge_with_current(base_options: list[str], current_series: pd.Series):
    merged = [""]
    for option in base_options:
        add_unique(merged, option)
    for value in current_series.dropna().tolist():
        add_unique(merged, value)
    return merged


def fill_blank_with_default(df: pd.DataFrame, defaults: dict):
    out = df.copy()

    out["DueDate"] = pd.to_datetime(out["DueDate"], errors="coerce")
    out["DueDate"] = out["DueDate"].fillna(pd.Timestamp(defaults["due_date"]))

    out["Schedule"] = out["Schedule"].astype("string").fillna("").str.strip()
    if defaults["Schedule"] != "":
        out.loc[out["Schedule"] == "", "Schedule"] = defaults["Schedule"]

    for col in ["Section", "Project", "Tag", "TaskName"]:
        out[col] = out[col].astype("string").fillna("").str.strip()
        if defaults[col] != "":
            out.loc[out[col] == "", col] = defaults[col]

    out["Estimated"] = pd.to_numeric(out["Estimated"], errors="coerce")
    out["Estimated"] = out["Estimated"].fillna(defaults["Estimated"])

    return out


def normalize_schedule_to_hhmm(series: pd.Series) -> pd.Series:
    normalized = []
    for value in series.tolist():
        if pd.isna(value):
            normalized.append("")
            continue
        text = str(value).strip()
        if text == "" or text.lower() == "none":
            normalized.append("")
            continue
        if re.match(r"^\d{3,4}$", text):
            if len(text) == 3:
                h = int(text[0])
                m = int(text[1:])
            else:
                h = int(text[:2])
                m = int(text[2:])
            if 0 <= h <= 23 and 0 <= m <= 59:
                normalized.append(f"{h:02d}:{m:02d}")
                continue
        parsed = pd.to_datetime(text, errors="coerce")
        if pd.notna(parsed):
            normalized.append(parsed.strftime("%H:%M"))
        else:
            normalized.append(text)
    return pd.Series(normalized, index=series.index, dtype="string")


def empty_rows(count: int):
    return pd.DataFrame(
        [
            {
                "DueDate": pd.NaT,
                "Schedule": "",
                "Section": "",
                "Project": "",
                "Tag": "",
                "TaskName": "",
                "Estimated": None,
            }
            for _ in range(count)
        ]
    )


def rows_from_task_lines(text: str, defaults: dict):
    tasks = [line.strip() for line in text.splitlines() if line.strip()]
    if not tasks:
        return empty_rows(0)

    due_date = defaults.get("due_date")
    due_value = pd.Timestamp(due_date) if due_date is not None else pd.NaT

    return pd.DataFrame(
        [
            {
                "DueDate": due_value,
                "Schedule": defaults.get("Schedule", ""),
                "Section": defaults.get("Section", ""),
                "Project": defaults.get("Project", ""),
                "Tag": defaults.get("Tag", ""),
                "TaskName": task,
                "Estimated": defaults.get("Estimated", None),
            }
            for task in tasks
        ]
    )


def cleanup_temp_downloads():
    now = datetime.now(timezone.utc)
    expired = [token for token, item in TEMP_DOWNLOADS.items() if item["expires_at"] <= now]
    for token in expired:
        del TEMP_DOWNLOADS[token]


def create_temp_download(data: bytes, file_name: str, ttl_minutes: int) -> tuple[str, datetime]:
    cleanup_temp_downloads()
    token = secrets.token_urlsafe(18)
    expires_at = datetime.now(timezone.utc) + timedelta(minutes=ttl_minutes)
    TEMP_DOWNLOADS[token] = {"data": data, "file_name": file_name, "expires_at": expires_at}
    return token, expires_at


def get_temp_download(token: str):
    cleanup_temp_downloads()
    item = TEMP_DOWNLOADS.get(token)
    if item is None:
        return None
    if item["expires_at"] <= datetime.now(timezone.utc):
        del TEMP_DOWNLOADS[token]
        return None
    return item


def make_qr_png(url: str) -> bytes:
    img = qrcode.make(url)
    buf = BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue()


def normalize_public_url(url: str) -> str:
    text = (url or "").strip()
    if text == "":
        return ""
    text = text.replace("https:://", "https://").replace("http:://", "http://")
    if not text.startswith("http://") and not text.startswith("https://"):
        text = "https://" + text
    return text.rstrip("/")


st.set_page_config(page_title="Task CSV Builder", layout="wide")
query_token = st.query_params.get("token")
if isinstance(query_token, list):
    query_token = query_token[0] if query_token else ""
if query_token:
    item = get_temp_download(str(query_token))
    st.title("一時ダウンロード")
    if item is None:
        st.error("このURLは無効か、有効期限切れです。PC側で再作成してください。")
    else:
        remaining = int((item["expires_at"] - datetime.now(timezone.utc)).total_seconds())
        st.success(f"ダウンロード可能です（残り約 {max(0, remaining)} 秒）")
        st.download_button(
            "CSVをダウンロード",
            data=item["data"],
            file_name=item["file_name"],
            mime="text/csv",
            use_container_width=True,
        )
    st.stop()

st.title("Task CSV Builder")
st.caption("スプレッドシート形式で入力してCSVを出力します。Schedule列は時間入力（空欄可）です。")

xlsm_path = resolve_xlsm_path()
section_options, project_options, tag_options = read_dropdown_options(xlsm_path)
if not section_options:
    section_options = FALLBACK_SECTIONS.copy()
if not project_options:
    project_options = FALLBACK_PROJECTS.copy()
if not tag_options:
    tag_options = FALLBACK_TAGS.copy()

default_due_date = date.today() + timedelta(days=1)
default_section = "10:00"
default_project = ".雑務"
default_estimated = 5

if "table_df" not in st.session_state:
    st.session_state.table_df = pd.DataFrame(columns=COLUMNS)
if "show_preview" not in st.session_state:
    st.session_state.show_preview = False
if "share_token" not in st.session_state:
    st.session_state.share_token = ""
if "share_expires_at" not in st.session_state:
    st.session_state.share_expires_at = None
if "share_url" not in st.session_state:
    st.session_state.share_url = ""

current_df = st.session_state.table_df.copy()
for col in ["Section", "Project", "Tag", "TaskName"]:
    if col not in current_df.columns:
        current_df[col] = ""
    current_df[col] = current_df[col].astype("string").fillna("")
if "DueDate" not in current_df.columns:
    current_df["DueDate"] = pd.NaT
if "Schedule" not in current_df.columns:
    current_df["Schedule"] = ""
current_df["Schedule"] = normalize_schedule_to_hhmm(current_df["Schedule"])
for col in ["Estimated"]:
    if col not in current_df.columns:
        current_df[col] = None
    current_df[col] = pd.to_numeric(current_df[col], errors="coerce")
current_df = current_df.reindex(columns=COLUMNS)

section_options = merge_with_current(section_options + [default_section], current_df["Section"])
project_options = merge_with_current(project_options + [default_project], current_df["Project"])
tag_options = merge_with_current(tag_options, current_df["Tag"])

st.subheader("空欄時のデフォルト値")
d1, d2, d3, d4, d5, d6, d7 = st.columns(7)
with d1:
    st.caption("DueDate")
with d2:
    st.caption("Schedule(HHMM)")
with d3:
    st.caption("Section")
with d4:
    st.caption("Project")
with d5:
    st.caption("Tag")
with d6:
    st.caption("TaskName")
with d7:
    st.caption("Estimated")

col1, col2, col3, col4, col5, col6, col7 = st.columns(7)
with col1:
    input_default_due_date = st.date_input("DueDate", value=default_due_date, label_visibility="collapsed")
with col2:
    input_default_schedule = st.text_input("Schedule", value="", placeholder="1725", label_visibility="collapsed")
with col3:
    input_default_section = st.selectbox(
        "Section",
        options=section_options,
        index=section_options.index(default_section) if default_section in section_options else 0,
        label_visibility="collapsed",
    )
with col4:
    input_default_project = st.selectbox(
        "Project",
        options=project_options,
        index=project_options.index(default_project) if default_project in project_options else 0,
        label_visibility="collapsed",
    )
with col5:
    input_default_tag = st.selectbox("Tag", options=tag_options, index=0, label_visibility="collapsed")
with col6:
    input_default_task_name = st.text_input("TaskName", value="", label_visibility="collapsed")
with col7:
    input_default_estimated = st.number_input("Estimated", min_value=0, value=default_estimated, step=1, label_visibility="collapsed")

defaults = {
    "due_date": input_default_due_date,
    "Schedule": input_default_schedule.strip(),
    "Section": input_default_section,
    "Project": input_default_project,
    "Tag": input_default_tag,
    "TaskName": input_default_task_name.strip(),
    "Estimated": int(input_default_estimated),
}

if st.session_state.get("bulk_clear_requested"):
    st.session_state["bulk_seed_text"] = ""
    st.session_state["bulk_signature"] = ""
    st.session_state["bulk_clear_requested"] = False
    for key in list(st.session_state.keys()):
        if key.startswith("bulk_due_") or key.startswith("bulk_schedule_"):
            del st.session_state[key]
        elif key.startswith("bulk_section_") or key.startswith("bulk_project_") or key.startswith("bulk_tag_"):
            del st.session_state[key]
        elif key.startswith("bulk_task_") or key.startswith("bulk_est_"):
            del st.session_state[key]

st.subheader("ダウンロード")
dl_col1, dl_col2 = st.columns([1, 2])
with dl_col1:
    encoding = st.selectbox(
        "CSV文字コード",
        options=["utf-8-sig", "cp932"],
        index=0,
        help="Excelで文字化けする場合は cp932 を選択してください。",
    )
    public_app_url_raw = st.text_input(
        "公開URL（QR生成用）",
        value=st.session_state.get("public_app_url", DEFAULT_PUBLIC_APP_URL),
        placeholder="https://xxxx.streamlit.app",
        help="Streamlit Cloud のアプリURLを入力してください。",
    ).strip()
    if public_app_url_raw == "":
        public_app_url_raw = DEFAULT_PUBLIC_APP_URL
    public_app_url = normalize_public_url(public_app_url_raw)
    st.session_state["public_app_url"] = public_app_url
    share_ttl = st.number_input("一時保存（分）", min_value=1, max_value=30, value=5, step=1)

st.subheader("タスク入力シート")
edited_df = st.data_editor(
    current_df,
    num_rows="fixed",
    use_container_width=True,
    key="task_sheet_editor",
    column_config={
        "DueDate": st.column_config.DateColumn("DueDate", format="YYYY/MM/DD"),
        "Schedule": st.column_config.TextColumn("Schedule", help="HHMM または HH:MM 形式"),
        "Section": st.column_config.SelectboxColumn("Section", options=section_options),
        "Project": st.column_config.SelectboxColumn("Project", options=project_options),
        "Tag": st.column_config.SelectboxColumn("Tag", options=tag_options),
        "TaskName": st.column_config.TextColumn("TaskName"),
        "Estimated": st.column_config.NumberColumn("Estimated", min_value=0, step=5),
    },
)
if not edited_df.equals(current_df):
    edited_df["Schedule"] = normalize_schedule_to_hhmm(edited_df["Schedule"])
    st.session_state.table_df = edited_df.copy()
    st.rerun()
else:
    edited_df["Schedule"] = normalize_schedule_to_hhmm(edited_df["Schedule"])
    st.session_state.table_df = edited_df.copy()

st.subheader("テキスト行数ぶんの入力ブロック")
bulk_text = st.text_area(
    "1行につき1タスク（行数ぶん入力ブロックを作成）",
    key="bulk_seed_text",
    placeholder="買い物\n資料作成\nメール返信",
    height=120,
)
bulk_lines = [line.strip() for line in bulk_text.splitlines() if line.strip()]
bulk_signature = "\n".join(bulk_lines)

if st.session_state.get("bulk_signature") != bulk_signature:
    for i, line in enumerate(bulk_lines):
        st.session_state[f"bulk_due_{i}"] = defaults["due_date"]
        st.session_state[f"bulk_schedule_{i}"] = defaults["Schedule"]
        st.session_state[f"bulk_section_{i}"] = defaults["Section"]
        st.session_state[f"bulk_project_{i}"] = defaults["Project"]
        st.session_state[f"bulk_tag_{i}"] = defaults["Tag"]
        st.session_state[f"bulk_task_{i}"] = line
        st.session_state[f"bulk_est_{i}"] = int(defaults["Estimated"])
    st.session_state["bulk_signature"] = bulk_signature

if not bulk_lines:
    st.info("下のテキストに1行以上入力すると、行数ぶん入力ブロックを表示します。")
else:
    h1, h2, h3, h4, h5, h6, h7 = st.columns(7)
    with h1:
        st.caption("DueDate")
    with h2:
        st.caption("Schedule(HHMM)")
    with h3:
        st.caption("Section")
    with h4:
        st.caption("Project")
    with h5:
        st.caption("Tag")
    with h6:
        st.caption("TaskName")
    with h7:
        st.caption("Estimated")

    for i, _line in enumerate(bulk_lines):
        b1, b2, b3, b4, b5, b6, b7 = st.columns(7)
        with b1:
            st.date_input("DueDate", key=f"bulk_due_{i}", label_visibility="collapsed")
        with b2:
            st.text_input("Schedule", key=f"bulk_schedule_{i}", label_visibility="collapsed", placeholder="1725")
        with b3:
            st.selectbox("Section", options=section_options, key=f"bulk_section_{i}", label_visibility="collapsed")
        with b4:
            st.selectbox("Project", options=project_options, key=f"bulk_project_{i}", label_visibility="collapsed")
        with b5:
            st.selectbox("Tag", options=tag_options, key=f"bulk_tag_{i}", label_visibility="collapsed")
        with b6:
            st.text_input("TaskName", key=f"bulk_task_{i}", label_visibility="collapsed")
        with b7:
            st.number_input("Estimated", min_value=0, step=1, key=f"bulk_est_{i}", label_visibility="collapsed")

    if st.button("入力ブロックをシートに追加"):
        rows = []
        for i in range(len(bulk_lines)):
            task_name = str(st.session_state.get(f"bulk_task_{i}", "")).strip()
            if task_name == "":
                continue
            rows.append(
                {
                    "DueDate": pd.Timestamp(st.session_state.get(f"bulk_due_{i}")),
                    "Schedule": str(st.session_state.get(f"bulk_schedule_{i}", "")).strip(),
                    "Section": st.session_state.get(f"bulk_section_{i}", ""),
                    "Project": st.session_state.get(f"bulk_project_{i}", ""),
                    "Tag": st.session_state.get(f"bulk_tag_{i}", ""),
                    "TaskName": task_name,
                    "Estimated": int(st.session_state.get(f"bulk_est_{i}", defaults["Estimated"])),
                }
            )
        if rows:
            st.session_state.table_df = pd.concat(
                [st.session_state.table_df, pd.DataFrame(rows)],
                ignore_index=True,
            )
            st.session_state["bulk_clear_requested"] = True
            st.rerun()

output_df = fill_blank_with_default(st.session_state.table_df, defaults).reindex(columns=COLUMNS)
output_df["TaskName"] = output_df["TaskName"].astype("string").fillna("").str.strip()
output_df = output_df[output_df["TaskName"] != ""].copy()

output_df["DueDate"] = pd.to_datetime(output_df["DueDate"], errors="coerce").dt.strftime("%Y/%m/%d")
output_df["DueDate"] = output_df["DueDate"].fillna("")
output_df["Schedule"] = normalize_schedule_to_hhmm(output_df["Schedule"]).fillna("")

csv_bytes = output_df.to_csv(index=False).encode(encoding, errors="replace")
with dl_col1:
    st.download_button(
        label="CSVをダウンロード",
        data=csv_bytes,
        file_name="tasks.csv",
        mime="text/csv",
    )
    if st.button("iPhone取り込み用 一時URLを作成", use_container_width=True):
        if not st.session_state.get("public_app_url", "").strip():
            st.error("先に「公開URL（QR生成用）」を入力してください。")
        else:
            token, expires_at = create_temp_download(csv_bytes, "tasks.csv", int(share_ttl))
            base = st.session_state["public_app_url"]
            share_url = f"{base}/?token={token}"
            st.session_state.share_token = token
            st.session_state.share_expires_at = expires_at
            st.session_state.share_url = share_url

with dl_col2:
    if st.session_state.share_url:
        remain = int((st.session_state.share_expires_at - datetime.now(timezone.utc)).total_seconds())
        if remain > 0:
            st.caption(f"一時URL（残り約 {remain} 秒）")
            st.code(st.session_state.share_url)
            st.image(make_qr_png(st.session_state.share_url), caption="iPhoneでQRを読み取ってダウンロード")
        else:
            st.warning("一時URLの有効期限が切れました。再作成してください。")
            st.session_state.share_url = ""
            st.session_state.share_token = ""
            st.session_state.share_expires_at = None
    if st.button("プレビュー表示"):
        st.session_state.show_preview = True
    if st.session_state.show_preview:
        st.dataframe(output_df, use_container_width=True, height=220)
