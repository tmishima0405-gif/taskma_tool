import re
import secrets
import base64
import zlib
import json
from datetime import date, timedelta
from datetime import datetime, timezone
from io import BytesIO
from pathlib import Path
from urllib.parse import urlsplit
from uuid import uuid4

import pandas as pd
import streamlit as st
import qrcode
import requests


COLUMNS = ["DueDate", "Schedule", "Section", "Project", "Tag", "TaskName", "Estimated"]
UI_COLUMN_ORDER = ["DueDate", "Schedule", "Section", "Project", "Tag", "Estimated", "TaskName"]
INPUT_COL_RATIOS = [0.7, 0.6, 0.6, 0.7, 0.7, 0.6, 2.2]
ESTIMATED_OPTIONS = [i for i in range(5, 121, 5)]
ESTIMATED_OTHER = "その他"
LIST_OPTIONS_PATH = Path(__file__).parent / "list_options.json"
DEFAULT_SECTIONS = ["05:00", "07:00", "08:00", "10:00", "11:00", "13:00", "14:00", "16:00", "19:00", "21:00", "22:00", "00:00"]
DEFAULT_PROJECTS = [
    ". プロジェクト化",
    ". 重要",
    ". 家のこと",
    ". 改善",
    ". 他",
    ". 情報",
    ". 家族",
    ". 整理",
    ". 勉強",
    ". 娯楽",
    ". オペ-90 臨時",
    ". 雑務",
]
DEFAULT_TAGS = ["", "電車", "ミーティング", "スキップ", "自分", "会議中", "プログラミング"]
DEFAULT_PUBLIC_APP_URL = "https://streamlitcsv.streamlit.app"
FIXED_CSV_ENCODING = "utf-8-sig"
FIXED_SHARE_TTL_MINUTES = 10
TEMP_DOWNLOADS = {}


def add_unique(items: list[str], value):
    if value is None:
        return
    text = str(value).strip()
    if text and text not in items:
        items.append(text)


def load_list_options():
    if LIST_OPTIONS_PATH.exists():
        try:
            data = json.loads(LIST_OPTIONS_PATH.read_text(encoding="utf-8"))
            sections = data.get("sections", [])
            projects = data.get("projects", [])
            tags = data.get("tags", [])
            if isinstance(sections, list) and isinstance(projects, list) and isinstance(tags, list):
                return sections, projects, tags
        except Exception:
            pass
    return DEFAULT_SECTIONS.copy(), DEFAULT_PROJECTS.copy(), DEFAULT_TAGS.copy()


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


def render_field_label(text: str):
    st.markdown(
        f"<div style='margin:0;padding:0;font-size:0.9rem;color:#111111;font-weight:600;line-height:1.05'>{text}</div>",
        unsafe_allow_html=True,
    )


def estimated_option_labels():
    return [str(v) for v in ESTIMATED_OPTIONS] + [ESTIMATED_OTHER]


def estimated_from_choice(choice: str, other_value: int) -> int:
    if choice == ESTIMATED_OTHER:
        return int(other_value)
    try:
        return int(choice)
    except Exception:
        return 5


def apply_app_style():
    st.markdown(
        """
<style>
@import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@500;600;700&family=Noto+Sans+JP:wght@400;500;700&display=swap');

:root {
  --ink: #111111;
  --muted: #4f5668;
  --line: #d9deea;
  --card: rgba(255,255,255,0.82);
  --brand: #1e66f5;
}

.stApp {
  font-family: 'Noto Sans JP', sans-serif;
  background:
    radial-gradient(1000px 380px at 0% 0%, #e8eefc 0%, rgba(232,238,252,0) 70%),
    radial-gradient(900px 300px at 100% 10%, #e7f6ee 0%, rgba(231,246,238,0) 70%),
    #eef2f8;
}

h1, h2, h3 {
  font-family: 'Space Grotesk', 'Noto Sans JP', sans-serif !important;
  color: var(--ink) !important;
  letter-spacing: 0.01em;
}

h1 {
  font-size: 2rem !important;
  font-weight: 700 !important;
}

h3 {
  border-left: 5px solid var(--brand);
  padding-left: 10px;
}

[data-testid="stTextInput"] input,
[data-testid="stDateInput"] input,
[data-testid="stNumberInput"] input {
  background: #ffffff !important;
  border-radius: 10px !important;
  border: 1.5px solid #9aa8c2 !important;
  box-shadow: 0 1px 0 rgba(0,0,0,0.04) !important;
}

[data-testid="stSelectbox"] > div > div {
  background: #ffffff !important;
  border-radius: 10px !important;
  border: 1.5px solid #9aa8c2 !important;
  box-shadow: 0 1px 0 rgba(0,0,0,0.04) !important;
}

[data-testid="stTextArea"] textarea {
  background: #ffffff !important;
  border: 1.5px solid #9aa8c2 !important;
  border-radius: 10px !important;
}

[data-testid="stDataEditor"] {
  border: 1px solid #c7d1e3;
  border-radius: 12px;
  background: #ffffff;
  padding: 6px;
}

.stButton button, .stDownloadButton button {
  border-radius: 10px !important;
  border: 1px solid #173f99 !important;
  font-weight: 600 !important;
}

.stButton button:hover, .stDownloadButton button:hover {
  border-color: var(--brand) !important;
  color: var(--brand) !important;
}

[data-testid="stCodeBlock"] {
  border-radius: 10px;
}
</style>
        """,
        unsafe_allow_html=True,
    )


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
    # If multiple URLs were concatenated, keep the last one.
    if text.count("http://") + text.count("https://") >= 2:
        last_https = text.rfind("https://")
        last_http = text.rfind("http://")
        start = max(last_https, last_http)
        if start >= 0:
            text = text[start:]
    if not text.startswith("http://") and not text.startswith("https://"):
        text = "https://" + text
    return text.rstrip("/")


def current_app_base_url() -> str:
    try:
        current = str(st.context.url)
    except Exception:
        current = ""
    if not current:
        return DEFAULT_PUBLIC_APP_URL
    parsed = urlsplit(current)
    if not parsed.scheme or not parsed.netloc:
        return DEFAULT_PUBLIC_APP_URL
    return f"{parsed.scheme}://{parsed.netloc}"


def get_supabase_config() -> tuple[str, str, str]:
    try:
        url = str(st.secrets.get("SUPABASE_URL", "")).strip().rstrip("/")
        key = str(st.secrets.get("SUPABASE_SERVICE_ROLE_KEY", "")).strip()
        bucket = str(st.secrets.get("SUPABASE_BUCKET", "")).strip()
        return url, key, bucket
    except Exception:
        return "", "", ""


def create_supabase_signed_csv_url(data: bytes, ttl_minutes: int) -> tuple[str, str]:
    supabase_url, supabase_key, supabase_bucket = get_supabase_config()
    if not supabase_url or not supabase_key or not supabase_bucket:
        return "", "Supabase Secretsが未設定です。"

    object_path = f"tmp/tasks_{datetime.now(timezone.utc).strftime('%Y%m%d_%H%M%S')}_{uuid4().hex[:8]}.csv"
    upload_headers = {
        "Authorization": f"Bearer {supabase_key}",
        "apikey": supabase_key,
        "Content-Type": "text/csv",
        "x-upsert": "true",
    }
    sign_headers = {
        "Authorization": f"Bearer {supabase_key}",
        "apikey": supabase_key,
        "Content-Type": "application/json",
    }
    upload_url = f"{supabase_url}/storage/v1/object/{supabase_bucket}/{object_path}"

    try:
        upload_res = requests.post(upload_url, headers=upload_headers, data=data, timeout=20)
        if upload_res.status_code not in (200, 201):
            return "", f"アップロード失敗: {upload_res.status_code}"
    except Exception as exc:
        return "", f"アップロード通信エラー: {exc}"

    sign_attempts = [
        (
            f"{supabase_url}/storage/v1/object/sign/{supabase_bucket}/{object_path}",
            {"expiresIn": int(ttl_minutes * 60)},
        ),
        (
            f"{supabase_url}/storage/v1/object/sign/{supabase_bucket}",
            {"path": object_path, "expiresIn": int(ttl_minutes * 60)},
        ),
        (
            f"{supabase_url}/storage/v1/object/sign/{supabase_bucket}",
            {"paths": [object_path], "expiresIn": int(ttl_minutes * 60)},
        ),
    ]

    last_error = ""
    for sign_url, sign_payload in sign_attempts:
        try:
            sign_res = requests.post(sign_url, headers=sign_headers, json=sign_payload, timeout=20)
            if sign_res.status_code != 200:
                body = sign_res.text[:200]
                last_error = f"署名URL作成失敗: {sign_res.status_code} {body}"
                continue

            body = sign_res.json()
            signed_url = body.get("signedURL", "")
            if not signed_url and isinstance(body.get("signedUrls"), list) and body["signedUrls"]:
                first = body["signedUrls"][0]
                if isinstance(first, dict):
                    signed_url = first.get("signedURL", "") or first.get("signedUrl", "")
                elif isinstance(first, str):
                    signed_url = first
            if not signed_url:
                last_error = "署名URLの取得に失敗しました。"
                continue

            if signed_url.startswith("http://") or signed_url.startswith("https://"):
                return signed_url, ""
            if signed_url.startswith("/storage/v1/"):
                return f"{supabase_url}{signed_url}", ""
            if signed_url.startswith("/object/"):
                return f"{supabase_url}/storage/v1{signed_url}", ""
            if signed_url.startswith("object/"):
                return f"{supabase_url}/storage/v1/{signed_url}", ""
            return f"{supabase_url}/{signed_url.lstrip('/')}", ""
        except Exception as exc:
            last_error = f"署名URL通信エラー: {exc}"

    return "", last_error or "署名URL作成に失敗しました。"


def encode_csv_payload(data: bytes, ttl_minutes: int) -> str:
    expires_at = int((datetime.now(timezone.utc) + timedelta(minutes=ttl_minutes)).timestamp())
    envelope = {
        "exp": expires_at,
        "b64": base64.b64encode(data).decode("ascii"),
    }
    raw = json.dumps(envelope, separators=(",", ":")).encode("utf-8")
    compressed = zlib.compress(raw, level=9)
    return base64.urlsafe_b64encode(compressed).decode("ascii").rstrip("=")


def decode_csv_payload(payload: str) -> tuple[bytes | None, int | None]:
    try:
        pad = "=" * ((4 - len(payload) % 4) % 4)
        compressed = base64.urlsafe_b64decode(payload + pad)
        raw = zlib.decompress(compressed)
        envelope = json.loads(raw.decode("utf-8"))
        exp = int(envelope.get("exp", 0))
        if exp <= int(datetime.now(timezone.utc).timestamp()):
            return None, exp
        data = base64.b64decode(envelope["b64"])
        return data, exp
    except Exception:
        return None, None


st.set_page_config(page_title="Task CSV Builder", layout="wide")
apply_app_style()
query_payload = st.query_params.get("payload")
if isinstance(query_payload, list):
    query_payload = query_payload[0] if query_payload else ""
if query_payload:
    payload_data, payload_exp = decode_csv_payload(str(query_payload))
    st.title("CSVダウンロード")
    if payload_data is None:
        if payload_exp is not None:
            st.error("このURLは有効期限切れです。PC側で再作成してください。")
        else:
            st.error("共有URLが壊れているか、非対応のデータです。")
    else:
        remain = max(0, payload_exp - int(datetime.now(timezone.utc).timestamp()))
        st.success(f"ダウンロード可能です（残り約 {remain} 秒）")
        st.download_button(
            "CSVをダウンロード",
            data=payload_data,
            file_name="tasks.csv",
            mime="text/csv",
            use_container_width=True,
        )
    st.stop()

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

section_options, project_options, tag_options = load_list_options()

default_due_date = date.today() + timedelta(days=1)
default_section = "10:00"
default_project = ". 雑務"
default_estimated = 5
default_tag = "自分"

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
d1, d2, d3, d4, d5, d6, d7 = st.columns(INPUT_COL_RATIOS)
with d1:
    render_field_label("DueDate")
with d2:
    render_field_label("Schedule(HHMM)")
with d3:
    render_field_label("Section")
with d4:
    render_field_label("Project")
with d5:
    render_field_label("Tag")
with d6:
    render_field_label("Estimated")
with d7:
    render_field_label("TaskName")

col1, col2, col3, col4, col5, col6, col7 = st.columns(INPUT_COL_RATIOS)
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
    input_default_tag = st.selectbox(
        "Tag",
        options=tag_options,
        index=tag_options.index(default_tag) if default_tag in tag_options else 0,
        label_visibility="collapsed",
    )
with col6:
    est_labels = estimated_option_labels()
    input_default_est_choice = st.selectbox(
        "Estimated選択",
        options=est_labels,
        index=est_labels.index(str(default_estimated)) if str(default_estimated) in est_labels else 0,
        key="default_est_choice",
        label_visibility="collapsed",
    )
    input_default_est_other = default_estimated
    if input_default_est_choice == ESTIMATED_OTHER:
        input_default_est_other = st.number_input(
            "Estimatedその他",
            min_value=0,
            value=default_estimated,
            step=1,
            key="default_est_other",
            label_visibility="collapsed",
        )
with col7:
    input_default_task_name = st.text_input("TaskName", value="", label_visibility="collapsed")

defaults = {
    "due_date": input_default_due_date,
    "Schedule": input_default_schedule.strip(),
    "Section": input_default_section,
    "Project": input_default_project,
    "Tag": input_default_tag,
    "TaskName": input_default_task_name.strip(),
    "Estimated": estimated_from_choice(input_default_est_choice, input_default_est_other),
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

st.subheader("ダウンロード用一時URL生成（10分間有効）")
dl_col1, dl_col2 = st.columns([1, 2])

st.subheader("テキスト行数ぶんの入力ブロック")
bulk_text = st.text_area(
    "下のテキストに1行以上入力すると、行数分入力ブロックを表示します（複数タスクは改行して入力）",
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
        if int(defaults["Estimated"]) in ESTIMATED_OPTIONS:
            st.session_state[f"bulk_est_choice_{i}"] = str(int(defaults["Estimated"]))
            st.session_state[f"bulk_est_other_{i}"] = 5
        else:
            st.session_state[f"bulk_est_choice_{i}"] = ESTIMATED_OTHER
            st.session_state[f"bulk_est_other_{i}"] = int(defaults["Estimated"])
    st.session_state["bulk_signature"] = bulk_signature

if bulk_lines:
    h1, h2, h3, h4, h5, h6, h7 = st.columns(INPUT_COL_RATIOS)
    with h1:
        render_field_label("DueDate")
    with h2:
        render_field_label("Schedule(HHMM)")
    with h3:
        render_field_label("Section")
    with h4:
        render_field_label("Project")
    with h5:
        render_field_label("Tag")
    with h6:
        render_field_label("Estimated")
    with h7:
        render_field_label("TaskName")

    for i, _line in enumerate(bulk_lines):
        b1, b2, b3, b4, b5, b6, b7 = st.columns(INPUT_COL_RATIOS)
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
            st.selectbox(
                "Estimated選択",
                options=estimated_option_labels(),
                key=f"bulk_est_choice_{i}",
                label_visibility="collapsed",
            )
            if st.session_state.get(f"bulk_est_choice_{i}") == ESTIMATED_OTHER:
                st.number_input(
                    "Estimatedその他",
                    min_value=0,
                    step=1,
                    key=f"bulk_est_other_{i}",
                    label_visibility="collapsed",
                )
        with b7:
            st.text_input("TaskName", key=f"bulk_task_{i}", label_visibility="collapsed")

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
                    "Estimated": estimated_from_choice(
                        st.session_state.get(f"bulk_est_choice_{i}", str(defaults["Estimated"])),
                        int(st.session_state.get(f"bulk_est_other_{i}", defaults["Estimated"])),
                    ),
                }
            )
        if rows:
            st.session_state.table_df = pd.concat(
                [st.session_state.table_df, pd.DataFrame(rows)],
                ignore_index=True,
            )
            st.session_state["bulk_clear_requested"] = True
            st.rerun()

st.subheader("タスク入力シート")
edited_df = st.data_editor(
    current_df,
    num_rows="fixed",
    use_container_width=True,
    key="task_sheet_editor",
    column_order=UI_COLUMN_ORDER,
    column_config={
        "DueDate": st.column_config.DateColumn("DueDate", format="YYYY/MM/DD", width="small"),
        "Schedule": st.column_config.TextColumn("Schedule", help="HHMM または HH:MM 形式", width="small"),
        "Section": st.column_config.SelectboxColumn("Section", options=section_options, width="small"),
        "Project": st.column_config.SelectboxColumn("Project", options=project_options, width="small"),
        "Tag": st.column_config.SelectboxColumn("Tag", options=tag_options, width="small"),
        "TaskName": st.column_config.TextColumn("TaskName", width="large"),
        "Estimated": st.column_config.NumberColumn("Estimated", min_value=0, step=5, width="small"),
    },
)
if not edited_df.equals(current_df):
    edited_df["Schedule"] = normalize_schedule_to_hhmm(edited_df["Schedule"])
    st.session_state.table_df = edited_df.copy()
    st.rerun()
else:
    edited_df["Schedule"] = normalize_schedule_to_hhmm(edited_df["Schedule"])
    st.session_state.table_df = edited_df.copy()

output_df = fill_blank_with_default(st.session_state.table_df, defaults).reindex(columns=COLUMNS)
output_df["TaskName"] = output_df["TaskName"].astype("string").fillna("").str.strip()
output_df = output_df[output_df["TaskName"] != ""].copy()

output_df["DueDate"] = pd.to_datetime(output_df["DueDate"], errors="coerce").dt.strftime("%Y/%m/%d")
output_df["DueDate"] = output_df["DueDate"].fillna("")
output_df["Schedule"] = normalize_schedule_to_hhmm(output_df["Schedule"]).fillna("")

csv_bytes = output_df.to_csv(index=False).encode(FIXED_CSV_ENCODING, errors="replace")
with dl_col1:
    btn1, btn2 = st.columns(2)
    btn1.download_button(
        label="CSVをダウンロード",
        data=csv_bytes,
        file_name="tasks.csv",
        mime="text/csv",
        use_container_width=True,
    )
    if btn2.button("iPhone取り込み用一時URLを作成", use_container_width=True):
        share_url, err = create_supabase_signed_csv_url(csv_bytes, FIXED_SHARE_TTL_MINUTES)
        if not share_url:
            base = normalize_public_url(current_app_base_url())
            payload = encode_csv_payload(csv_bytes, FIXED_SHARE_TTL_MINUTES)
            share_url = f"{base}/?payload={payload}"
            st.warning(f"Supabase URL作成に失敗したため、アプリ内一時URLで作成しました: {err}")
        st.session_state.share_token = ""
        st.session_state.share_expires_at = datetime.now(timezone.utc) + timedelta(minutes=FIXED_SHARE_TTL_MINUTES)
        st.session_state.share_url = share_url

with dl_col2:
    if st.session_state.share_url:
        remain = int((st.session_state.share_expires_at - datetime.now(timezone.utc)).total_seconds())
        if remain > 0:
            st.caption(f"一時URL（残り約 {remain} 秒）")
            st.code(st.session_state.share_url)
            st.image(
                make_qr_png(st.session_state.share_url),
                caption="iPhoneでQRを読み取ってダウンロード",
                width=540,
            )
        else:
            st.warning("一時URLの有効期限が切れました。再作成してください。")
            st.session_state.share_url = ""
            st.session_state.share_token = ""
            st.session_state.share_expires_at = None
    if st.button("プレビュー表示"):
        st.session_state.show_preview = True
    if st.session_state.show_preview:
        st.dataframe(output_df, use_container_width=True, height=220)
