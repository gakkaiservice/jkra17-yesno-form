import os
import re
import sqlite3
import hashlib
import hmac
import secrets
import smtplib
import ssl
from email.message import EmailMessage
from io import BytesIO
from pathlib import Path
from datetime import datetime, date

import streamlit as st
from python_calamine import CalamineWorkbook
import xlsxwriter

APP_NAME = "登壇諾否マイページ｜共通管理版 v3.6"
DB_PATH = os.getenv("YESNO_DB_PATH", "yesno_common.db")
ATTACH_DIR = Path(os.getenv("YESNO_ATTACH_DIR", "attachments"))


def get_secret(name, default=""):
    try:
        return st.secrets.get(name, os.getenv(name, default))
    except Exception:
        return os.getenv(name, default)


ADMIN_PASSWORD = get_secret("ADMIN_PASSWORD", "")
TOKEN_SECRET = get_secret("TOKEN_SECRET", "local-development-secret").encode("utf-8")

# 共通メール送信設定（Streamlit Secretsで1回だけ設定）
SMTP_HOST = get_secret("SMTP_HOST", "").strip()
SMTP_PORT = int(get_secret("SMTP_PORT", "587") or 587)
SMTP_USERNAME = get_secret("SMTP_USERNAME", "").strip()
SMTP_PASSWORD = get_secret("SMTP_PASSWORD", "")
SMTP_SECURITY = get_secret("SMTP_SECURITY", "starttls").strip().lower()  # starttls / ssl / none
SMTP_FROM_EMAIL = get_secret("SMTP_FROM_EMAIL", SMTP_USERNAME).strip()

def get_base_url():
    """Return the actual app URL automatically.

    YESNO_BASE_URL can still override it, but placeholder/example values are ignored.
    Streamlit 1.46+ exposes the current app URL via st.context.url.
    """
    configured = get_secret("YESNO_BASE_URL", "").strip().rstrip("/")
    if configured and "あなたのURL" not in configured and "your-app" not in configured.lower():
        return configured
    try:
        current = str(st.context.url).strip().rstrip("/")
        if current:
            return current
    except Exception:
        pass
    return "http://localhost:8501"


def person_url(conf_code, token):
    return f"{get_base_url()}/?c={conf_code}&token={token}"

st.set_page_config(page_title=APP_NAME, page_icon="✅", layout="wide")

st.markdown("""
<style>
.block-container {max-width: 1180px; padding-top: 1.4rem; padding-bottom: 4rem;}
.hero {padding: 1.2rem 1.4rem; border-radius: 18px; background: linear-gradient(135deg,#15334b,#246681); color:#fff; margin-bottom: 1rem;}
.hero h1 {font-size:1.55rem; margin:0 0 .2rem 0; color:#fff;}
.hero p {margin:0; opacity:.92;}
.person {font-size:1.25rem; font-weight:750; margin:.7rem 0 .2rem;}
.subtle {color:#667788; font-size:.92rem;}
.card {border:1px solid #dde5eb; border-radius:16px; padding:1rem 1.15rem; margin:.75rem 0; background:#fff; box-shadow:0 2px 12px rgba(25,50,70,.05);}
.card.pending {border-left:6px solid #b97813;}
.card.done {border-left:6px solid #2b7a50; background:#fbfdfb;}
.badge {display:inline-block; border-radius:999px; padding:.18rem .65rem; font-size:.83rem; font-weight:750; margin-bottom:.55rem;}
.badge.pending {background:#fff2d9; color:#80520b;}
.badge.yes {background:#e7f6ed; color:#17683b;}
.badge.no {background:#fdeaea; color:#9b2c2c;}
.meta {display:grid; grid-template-columns:110px 1fr; gap:.2rem .6rem; font-size:.94rem; margin-top:.4rem;}
.meta .k {color:#718090;}
.notice {padding:.8rem 1rem; border-radius:12px; background:#f4f7fa; color:#3d5263; margin:.7rem 0;}
.warn {padding:.8rem 1rem; border-radius:12px; background:#fff7e8; color:#704b10; margin:.7rem 0;}
</style>
""", unsafe_allow_html=True)


def clean(v):
    if v is None:
        return ""
    if isinstance(v, bool):
        return "〇" if v else ""
    if isinstance(v, datetime):
        return v.strftime("%Y/%m/%d %H:%M")
    if isinstance(v, date):
        return v.strftime("%Y/%m/%d")
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()


def normalize_answer(v):
    s = clean(v)
    if s in {"諾", "承諾", "ご承諾"} or "ご承諾" in s:
        return "諾"
    if s in {"否", "辞退", "ご辞退"} or "ご辞退" in s:
        return "否"
    return ""


def normalize_invitation(v):
    s = clean(v)
    if "必要" in s and "不要" not in s:
        return "必要"
    if "不要" in s:
        return "不要"
    return ""


def bool_int(v):
    return 1 if v else 0


def connect():
    con = sqlite3.connect(DB_PATH)
    con.row_factory = sqlite3.Row
    return con


def init_db():
    with connect() as con:
        con.executescript("""
        CREATE TABLE IF NOT EXISTS conferences (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            code TEXT UNIQUE NOT NULL,
            name TEXT NOT NULL,
            dates TEXT,
            venue TEXT,
            reply_deadline TEXT,
            active INTEGER NOT NULL DEFAULT 1,
            ask_furigana INTEGER NOT NULL DEFAULT 1,
            ask_membership INTEGER NOT NULL DEFAULT 0,
            membership_label TEXT,
            ask_mobile INTEGER NOT NULL DEFAULT 0,
            ask_correction INTEGER NOT NULL DEFAULT 1,
            invitation_enabled INTEGER NOT NULL DEFAULT 0,
            invitation_label TEXT,
            auto_reply_enabled INTEGER NOT NULL DEFAULT 1,
            office_notify_enabled INTEGER NOT NULL DEFAULT 1,
            office_email TEXT,
            sender_name TEXT,
            reply_to_email TEXT,
            auto_reply_subject TEXT,
            auto_reply_body TEXT,
            office_subject TEXT,
            office_body TEXT,
            created_at TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS people (
            conference_id INTEGER NOT NULL,
            token TEXT NOT NULL,
            name TEXT,
            email TEXT,
            affiliation TEXT,
            furigana TEXT,
            membership TEXT,
            mobile TEXT,
            correction TEXT,
            invitation TEXT,
            leader_org TEXT,
            leader_title TEXT,
            leader_name TEXT,
            special_request TEXT,
            upload_name TEXT,
            upload_path TEXT,
            registered_at TEXT,
            source TEXT,
            PRIMARY KEY(conference_id, token),
            FOREIGN KEY(conference_id) REFERENCES conferences(id)
        );
        CREATE TABLE IF NOT EXISTS requests (
            conference_id INTEGER NOT NULL,
            request_id TEXT PRIMARY KEY,
            source_sheet TEXT,
            source_row INTEGER,
            source_seq TEXT,
            token TEXT NOT NULL,
            name TEXT NOT NULL,
            affiliation TEXT,
            email TEXT,
            session_name TEXT,
            theme TEXT,
            role TEXT,
            schedule TEXT,
            source_answer TEXT,
            deadline TEXT,
            request_sent TEXT,
            imported_at TEXT NOT NULL,
            FOREIGN KEY(conference_id) REFERENCES conferences(id)
        );
        CREATE TABLE IF NOT EXISTS responses (
            request_id TEXT PRIMARY KEY,
            conference_id INTEGER NOT NULL,
            answer TEXT NOT NULL,
            decline_reason TEXT,
            note TEXT,
            responded_at TEXT NOT NULL,
            FOREIGN KEY(request_id) REFERENCES requests(request_id),
            FOREIGN KEY(conference_id) REFERENCES conferences(id)
        );
        """)
        # 既存DBをv3.5へ自動移行
        cols = {r[1] for r in con.execute("PRAGMA table_info(conferences)").fetchall()}
        additions = {
            "auto_reply_enabled": "INTEGER NOT NULL DEFAULT 1",
            "office_notify_enabled": "INTEGER NOT NULL DEFAULT 1",
            "office_email": "TEXT",
            "sender_name": "TEXT",
            "reply_to_email": "TEXT",
            "auto_reply_subject": "TEXT",
            "auto_reply_body": "TEXT",
            "office_subject": "TEXT",
            "office_body": "TEXT",
        }
        for col, ddl in additions.items():
            if col not in cols:
                con.execute(f"ALTER TABLE conferences ADD COLUMN {col} {ddl}")


def hero(title, text=""):
    st.markdown(f'<div class="hero"><h1>{title}</h1><p>{text}</p></div>', unsafe_allow_html=True)


def conference_token(conf_code, email, name):
    key = f"{conf_code}|{clean(email).lower() or clean(name)}".encode("utf-8")
    return hmac.new(TOKEN_SECRET, key, hashlib.sha256).hexdigest()[:32]


def request_id(conf_code, row_no, seq_no, session, role, name):
    stable = clean(seq_no) or f"row-{row_no}"
    raw = f"{conf_code}|{stable}|{clean(session)}|{clean(role)}|{clean(name)}"
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()[:24]


def list_conferences(active_only=False):
    q = "SELECT * FROM conferences"
    if active_only:
        q += " WHERE active=1"
    q += " ORDER BY id DESC"
    with connect() as con:
        return con.execute(q).fetchall()


def get_conference_by_code(code):
    with connect() as con:
        return con.execute("SELECT * FROM conferences WHERE code=?", (code,)).fetchone()


def get_conference(cid):
    with connect() as con:
        return con.execute("SELECT * FROM conferences WHERE id=?", (cid,)).fetchone()


def create_conference(values):
    now = datetime.now().isoformat(timespec="seconds")
    with connect() as con:
        con.execute("""
        INSERT INTO conferences(code,name,dates,venue,reply_deadline,active,ask_furigana,ask_membership,membership_label,
                                ask_mobile,ask_correction,invitation_enabled,invitation_label,
                                auto_reply_enabled,office_notify_enabled,office_email,sender_name,reply_to_email,
                                auto_reply_subject,auto_reply_body,office_subject,office_body,created_at)
        VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
        """, (
            values["code"], values["name"], values.get("dates", ""), values.get("venue", ""), values.get("reply_deadline", ""),
            1, bool_int(values.get("ask_furigana")), bool_int(values.get("ask_membership")), values.get("membership_label", ""),
            bool_int(values.get("ask_mobile")), bool_int(values.get("ask_correction")), bool_int(values.get("invitation_enabled")),
            values.get("invitation_label", ""), bool_int(values.get("auto_reply_enabled", True)),
            bool_int(values.get("office_notify_enabled", True)), values.get("office_email", ""), values.get("sender_name", ""),
            values.get("reply_to_email", ""), values.get("auto_reply_subject", ""), values.get("auto_reply_body", ""),
            values.get("office_subject", ""), values.get("office_body", ""), now
        ))


def update_conference(cid, values):
    with connect() as con:
        con.execute("""
        UPDATE conferences SET name=?,dates=?,venue=?,reply_deadline=?,active=?,ask_furigana=?,ask_membership=?,membership_label=?,
            ask_mobile=?,ask_correction=?,invitation_enabled=?,invitation_label=?,
            auto_reply_enabled=?,office_notify_enabled=?,office_email=?,sender_name=?,reply_to_email=?,
            auto_reply_subject=?,auto_reply_body=?,office_subject=?,office_body=? WHERE id=?
        """, (
            values["name"], values.get("dates", ""), values.get("venue", ""), values.get("reply_deadline", ""), bool_int(values.get("active")),
            bool_int(values.get("ask_furigana")), bool_int(values.get("ask_membership")), values.get("membership_label", ""),
            bool_int(values.get("ask_mobile")), bool_int(values.get("ask_correction")), bool_int(values.get("invitation_enabled")),
            values.get("invitation_label", ""), bool_int(values.get("auto_reply_enabled", True)),
            bool_int(values.get("office_notify_enabled", True)), values.get("office_email", ""), values.get("sender_name", ""),
            values.get("reply_to_email", ""), values.get("auto_reply_subject", ""), values.get("auto_reply_body", ""),
            values.get("office_subject", ""), values.get("office_body", ""), cid
        ))



def delete_conference(cid):
    """Delete one conference and all related local data."""
    with connect() as con:
        # Delete children first because foreign-key cascading is not assumed.
        con.execute("DELETE FROM responses WHERE conference_id=?", (cid,))
        con.execute("DELETE FROM requests WHERE conference_id=?", (cid,))
        con.execute("DELETE FROM people WHERE conference_id=?", (cid,))
        con.execute("DELETE FROM conferences WHERE id=?", (cid,))
        con.commit()

def excel_col(n):
    s = ""
    n += 1
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def read_workbook_sheet(file_bytes, sheet_name=None):
    wb = CalamineWorkbook.from_filelike(BytesIO(file_bytes))
    names = wb.sheet_names
    if not names:
        raise ValueError("シートが見つかりません。")
    if not sheet_name:
        sheet_name = "指定演題" if "指定演題" in names else names[0]
    sheet = wb.get_sheet_by_name(sheet_name)
    rows = sheet.to_python(skip_empty_area=False)
    if not rows:
        raise ValueError(f"「{sheet_name}」シートにデータがありません。")
    headers = [clean(x).replace("\n", "") for x in rows[0]]
    return names, sheet_name, headers, rows


def header_indices(headers):
    out = {}
    for i, h in enumerate(headers):
        if h:
            out.setdefault(h, []).append(i)
    return out


def first_index(indices, candidates, occurrence="first"):
    for h in candidates:
        locs = indices.get(h, [])
        if locs:
            return locs[-1] if occurrence == "last" else locs[0]
    return None


def auto_mapping(headers):
    idx = header_indices(headers)
    mapping = {
        "name": first_index(idx, ["元）氏名", "氏名", "名前"]),
        "session": first_index(idx, ["セッション名", "セッション", "企画名"]),
        "role": first_index(idx, ["役割", "担当", "役割名"]),
        "affiliation": first_index(idx, ["所属", "元）所属", "所属機関"]),
        "email": first_index(idx, ["メールアドレス", "E-mail", "Email"], "first"),
        "theme": first_index(idx, ["元）セッションテーマ（キャッチーなテーマ名）", "セッションテーマ", "テーマ", "演題名"]),
        "schedule": first_index(idx, ["候補日時", "日時", "登壇日時"]),
        "answer": first_index(idx, ["諾否", "回答", "諾否回答"], "first"),
        "deadline": first_index(idx, ["諾否締切", "回答期限", "締切"]),
        "sent": first_index(idx, ["依頼送信", "依頼状況", "送信状況"]),
        "seq": first_index(idx, ["スプシセッション順No.", "依頼ID", "No.", "番号"]),
    }
    legacy = {
        "receipt": first_index(idx, ["受付番号"]),
        "sent_at": first_index(idx, ["送信日時"]),
        "form_name": first_index(idx, ["氏名"], "last"),
        "form_email": first_index(idx, ["メールアドレス", "E-mail", "Email"], "last"),
        "furigana": first_index(idx, ["ふりがな", "フリガナ"]),
        "membership": first_index(idx, ["会員・非会員 （日本腎臓リハビリテーション学会）", "会員・非会員", "会員区分"]),
        "mobile": first_index(idx, ["携帯電話番号", "緊急連絡先", "携帯電話"]),
        "correction": first_index(idx, ["修正内容のご入力", "氏名・ご所属等の修正依頼", "修正依頼"]),
        "invitation": first_index(idx, ["派遣依頼状（招聘状）", "招聘状", "派遣依頼状"]),
        "leader_org": first_index(idx, ["所属長の所属機関名"]),
        "leader_title": first_index(idx, ["所属長の役職"]),
        "leader_name": first_index(idx, ["所属長の氏名"]),
        "special_request": first_index(idx, ["指定様式での発行、web申請のご希望", "指定様式・web申請"]),
        "upload_name": first_index(idx, ["派遣依頼状　指定様式アップロード", "派遣依頼状 指定様式アップロード"]),
    }
    return mapping, legacy


def cell(row, index):
    if index is None or index >= len(row):
        return ""
    return row[index]


def mapping_display(headers, mapping):
    rows = []
    labels = {
        "name":"氏名", "session":"セッション名", "role":"役割", "affiliation":"所属", "email":"メール",
        "theme":"テーマ", "schedule":"日時", "answer":"既存諾否", "deadline":"回答期限", "sent":"依頼状況", "seq":"依頼ID/順番"
    }
    for k, label in labels.items():
        i = mapping.get(k)
        rows.append({"項目":label, "検出列": "未検出" if i is None else f"{excel_col(i)}列：{headers[i]}"})
    return rows


def import_master(conf, file_bytes, sheet_name=None):
    _, sheet_name, headers, rows = read_workbook_sheet(file_bytes, sheet_name)
    mapping, legacy = auto_mapping(headers)
    missing = [x for x in ("name", "session", "role") if mapping[x] is None]
    if missing:
        raise ValueError("氏名・セッション名・役割の列を自動判定できませんでした。Excelの見出し名をご確認ください。")

    records, profiles = [], {}
    for excel_row, row in enumerate(rows[1:], start=2):
        name = clean(cell(row, mapping["name"]))
        session = clean(cell(row, mapping["session"]))
        role = clean(cell(row, mapping["role"]))
        if not name or not session or not role:
            continue
        affiliation = clean(cell(row, mapping["affiliation"]))
        email = clean(cell(row, mapping["email"]))
        theme = clean(cell(row, mapping["theme"]))
        schedule = clean(cell(row, mapping["schedule"]))
        source_answer = normalize_answer(cell(row, mapping["answer"]))
        deadline = clean(cell(row, mapping["deadline"])) or clean(conf["reply_deadline"])
        sent = clean(cell(row, mapping["sent"]))
        seq = clean(cell(row, mapping["seq"]))
        token = conference_token(conf["code"], email, name)
        rid = request_id(conf["code"], excel_row, seq, session, role, name)
        records.append({
            "conference_id": conf["id"], "request_id": rid, "source_sheet": sheet_name, "source_row": excel_row,
            "source_seq": seq, "token": token, "name": name, "affiliation": affiliation, "email": email,
            "session_name": session, "theme": theme, "role": role, "schedule": schedule,
            "source_answer": source_answer, "deadline": deadline, "request_sent": sent,
        })

        legacy_signal = clean(cell(row, legacy["receipt"])) or clean(cell(row, legacy["sent_at"])) or clean(cell(row, legacy["form_name"]))
        if legacy_signal:
            p = {
                "conference_id": conf["id"], "token": token,
                "name": clean(cell(row, legacy["form_name"])) or name,
                "email": clean(cell(row, legacy["form_email"])) or email,
                "affiliation": affiliation,
                "furigana": clean(cell(row, legacy["furigana"])),
                "membership": clean(cell(row, legacy["membership"])),
                "mobile": clean(cell(row, legacy["mobile"])),
                "correction": clean(cell(row, legacy["correction"])),
                "invitation": normalize_invitation(cell(row, legacy["invitation"])),
                "leader_org": clean(cell(row, legacy["leader_org"])),
                "leader_title": clean(cell(row, legacy["leader_title"])),
                "leader_name": clean(cell(row, legacy["leader_name"])),
                "special_request": clean(cell(row, legacy["special_request"])),
                "upload_name": clean(cell(row, legacy["upload_name"])), "upload_path": "",
                "registered_at": clean(cell(row, legacy["sent_at"])) or datetime.now().isoformat(timespec="seconds"),
                "source": "legacy_import",
            }
            old = profiles.get(token)
            if not old or p["registered_at"] >= old["registered_at"]:
                profiles[token] = p

    now = datetime.now().isoformat(timespec="seconds")
    ids = {r["request_id"] for r in records}
    with connect() as con:
        existing = [x[0] for x in con.execute("SELECT request_id FROM requests WHERE conference_id=?", (conf["id"],)).fetchall()]
        for rid in existing:
            if rid not in ids:
                con.execute("DELETE FROM responses WHERE request_id=?", (rid,))
                con.execute("DELETE FROM requests WHERE request_id=?", (rid,))
        for r in records:
            con.execute("""
            INSERT INTO requests(conference_id,request_id,source_sheet,source_row,source_seq,token,name,affiliation,email,session_name,theme,role,schedule,source_answer,deadline,request_sent,imported_at)
            VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
            ON CONFLICT(request_id) DO UPDATE SET
                source_sheet=excluded.source_sheet,source_row=excluded.source_row,source_seq=excluded.source_seq,token=excluded.token,
                name=excluded.name,affiliation=excluded.affiliation,email=excluded.email,session_name=excluded.session_name,
                theme=excluded.theme,role=excluded.role,schedule=excluded.schedule,source_answer=excluded.source_answer,
                deadline=excluded.deadline,request_sent=excluded.request_sent,imported_at=excluded.imported_at
            """, (
                r["conference_id"],r["request_id"],r["source_sheet"],r["source_row"],r["source_seq"],r["token"],r["name"],r["affiliation"],
                r["email"],r["session_name"],r["theme"],r["role"],r["schedule"],r["source_answer"],r["deadline"],r["request_sent"],now
            ))
    for p in profiles.values():
        upsert_profile(p, overwrite=False)
    return records, list(profiles.values()), headers, mapping


def upsert_profile(p, overwrite=False):
    with connect() as con:
        current = con.execute("SELECT * FROM people WHERE conference_id=? AND token=?", (p["conference_id"], p["token"])).fetchone()
        if current and not overwrite and current["source"] == "system":
            return
        con.execute("""
        INSERT INTO people(conference_id,token,name,email,affiliation,furigana,membership,mobile,correction,invitation,leader_org,leader_title,leader_name,
                           special_request,upload_name,upload_path,registered_at,source)
        VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
        ON CONFLICT(conference_id,token) DO UPDATE SET
            name=excluded.name,email=excluded.email,affiliation=excluded.affiliation,furigana=excluded.furigana,membership=excluded.membership,
            mobile=excluded.mobile,correction=excluded.correction,invitation=excluded.invitation,leader_org=excluded.leader_org,
            leader_title=excluded.leader_title,leader_name=excluded.leader_name,special_request=excluded.special_request,
            upload_name=excluded.upload_name,upload_path=CASE WHEN excluded.upload_path<>'' THEN excluded.upload_path ELSE people.upload_path END,
            registered_at=excluded.registered_at,source=excluded.source
        """, (
            p["conference_id"],p["token"],p.get("name", ""),p.get("email", ""),p.get("affiliation", ""),p.get("furigana", ""),
            p.get("membership", ""),p.get("mobile", ""),p.get("correction", ""),p.get("invitation", ""),p.get("leader_org", ""),
            p.get("leader_title", ""),p.get("leader_name", ""),p.get("special_request", ""),p.get("upload_name", ""),p.get("upload_path", ""),
            p.get("registered_at", ""),p.get("source", "system")
        ))


def get_profile(conf_id, token):
    with connect() as con:
        return con.execute("SELECT * FROM people WHERE conference_id=? AND token=?", (conf_id, token)).fetchone()


def get_person_requests(conf_id, token):
    with connect() as con:
        return con.execute("""
        SELECT r.*, s.answer AS new_answer, s.decline_reason, s.note, s.responded_at
        FROM requests r LEFT JOIN responses s ON s.request_id=r.request_id
        WHERE r.conference_id=? AND r.token=? ORDER BY r.source_row
        """, (conf_id, token)).fetchall()


def all_rows(conf_id):
    with connect() as con:
        return con.execute("""
        SELECT r.*, s.answer AS new_answer, s.decline_reason, s.note, s.responded_at
        FROM requests r LEFT JOIN responses s ON s.request_id=r.request_id
        WHERE r.conference_id=? ORDER BY r.source_row
        """, (conf_id,)).fetchall()


def all_people(conf_id):
    with connect() as con:
        return con.execute("SELECT * FROM people WHERE conference_id=? ORDER BY name", (conf_id,)).fetchall()


def effective_answer(r):
    return r["source_answer"] or r["new_answer"] or ""


def save_response(conf_id, request_id_value, answer, decline_reason, note):
    with connect() as con:
        src = con.execute("SELECT source_answer FROM requests WHERE conference_id=? AND request_id=?", (conf_id, request_id_value)).fetchone()
        if not src:
            raise ValueError("依頼が見つかりません。")
        if clean(src[0]):
            raise ValueError("この依頼は既に回答済みのため変更できません。")
        con.execute("""
        INSERT INTO responses(request_id,conference_id,answer,decline_reason,note,responded_at)
        VALUES(?,?,?,?,?,?)
        ON CONFLICT(request_id) DO UPDATE SET answer=excluded.answer,decline_reason=excluded.decline_reason,note=excluded.note,responded_at=excluded.responded_at
        """, (request_id_value, conf_id, answer, clean(decline_reason), clean(note), datetime.now().isoformat(timespec="seconds")))


DEFAULT_AUTO_SUBJECT = "【回答受付】登壇諾否のご回答ありがとうございます"
DEFAULT_AUTO_INTRO = """※本メールはご登録いただいたメールアドレスへ自動送信しております。

この度は、指定演題ご依頼の諾否につきまして、ご回答を賜り誠にありがとうございます。
下記の通り、ご回答を受け付けいたしました。

＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
ご登録内容の変更について
＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
本メールの返信にてご連絡ください。
"""


def smtp_ready():
    return bool(SMTP_HOST and SMTP_PORT and SMTP_FROM_EMAIL)


def render_mail_template(template, values):
    class SafeDict(dict):
        def __missing__(self, key):
            return "{" + key + "}"
    return (template or "").format_map(SafeDict(values))


def send_email_message(to_email, subject, body, sender_name="", reply_to="", bcc_email=""):
    if not smtp_ready():
        raise RuntimeError("SMTP設定が未完了です。管理画面の『運用設定』をご確認ください。")
    if not to_email:
        raise RuntimeError("送信先メールアドレスが空です。")
    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = f"{sender_name} <{SMTP_FROM_EMAIL}>" if sender_name else SMTP_FROM_EMAIL
    msg["To"] = to_email
    if reply_to:
        msg["Reply-To"] = reply_to
    if bcc_email:
        msg["Bcc"] = bcc_email
    msg.set_content(body)
    if SMTP_SECURITY == "ssl":
        context = ssl.create_default_context()
        with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT, context=context, timeout=20) as server:
            if SMTP_USERNAME:
                server.login(SMTP_USERNAME, SMTP_PASSWORD)
            server.send_message(msg)
    else:
        with smtplib.SMTP(SMTP_HOST, SMTP_PORT, timeout=20) as server:
            server.ehlo()
            if SMTP_SECURITY == "starttls":
                server.starttls(context=ssl.create_default_context())
                server.ehlo()
            if SMTP_USERNAME:
                server.login(SMTP_USERNAME, SMTP_PASSWORD)
            server.send_message(msg)


def send_response_notifications(conf, token, submitted_rows, answers_map, note, profile, first_registration=False, decline_map=None):
    """Send one automatic receipt email to the respondent and BCC the office.

    The administrator edits only the plain subject and introductory text.  All
    response/profile details below are generated automatically, so no merge
    tags are required.
    """
    now = datetime.now().strftime("%Y/%m/%d %H:%M:%S")
    decline_map = decline_map or {}
    name = (profile["name"] if profile else submitted_rows[0]["name"]) or ""
    email = (profile["email"] if profile else submitted_rows[0]["email"]) or ""
    affiliation = (profile["affiliation"] if profile else submitted_rows[0]["affiliation"]) or ""

    answer_lines = []
    for r in submitted_rows:
        ans = answers_map.get(r["request_id"], "")
        label = "承諾" if ans == "諾" else "辞退" if ans == "否" else ans
        answer_lines.extend([
            f"セッション：{r['session_name'] or ''}",
            f"役割：{r['role'] or ''}",
            f"回答：{label}",
        ])
        if r["theme"]:
            answer_lines.append(f"テーマ：{r['theme']}")
        if r["schedule"]:
            answer_lines.append(f"日時：{r['schedule']}")
        reason = clean(decline_map.get(r["request_id"], ""))
        if ans == "否" and reason:
            answer_lines.append(f"辞退理由：{reason}")
        answer_lines.append("")

    registration_lines = []
    if first_registration and profile:
        registration_lines.extend([
            "＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝",
            "ご登録情報",
            "＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝",
            f"氏名：{profile['name'] or ''}",
            f"所属：{profile['affiliation'] or ''}",
            f"メールアドレス：{profile['email'] or ''}",
        ])
        if conf["ask_furigana"]:
            registration_lines.append(f"ふりがな：{profile['furigana'] or ''}")
        if conf["ask_membership"]:
            registration_lines.append(f"{conf['membership_label'] or '会員区分'}：{profile['membership'] or ''}")
        if conf["ask_mobile"]:
            registration_lines.append(f"緊急連絡先：{profile['mobile'] or ''}")
        if conf["ask_correction"]:
            registration_lines.append(f"氏名・所属等の修正依頼：{profile['correction'] or 'なし'}")
        if conf["invitation_enabled"]:
            registration_lines.append(f"{conf['invitation_label'] or '招聘状・派遣依頼状'}：{profile['invitation'] or ''}")
            if profile["invitation"] == "必要":
                registration_lines.extend([
                    f"所属長の所属機関名：{profile['leader_org'] or ''}",
                    f"所属長の役職：{profile['leader_title'] or ''}",
                    f"所属長の氏名：{profile['leader_name'] or ''}",
                    f"指定様式・web申請等：{profile['special_request'] or 'なし'}",
                ])
        registration_lines.append("")

    intro = (conf["auto_reply_body"] or DEFAULT_AUTO_INTRO).strip()
    subject = (conf["auto_reply_subject"] or DEFAULT_AUTO_SUBJECT).strip()
    # No merge syntax is required. Conference/name are added automatically.
    body_parts = [
        "※本メールはご登録いただいたメールアドレスへ自動送信しております。",
        "",
        f"{name} 先生",
        "",
        intro.replace("※本メールはご登録いただいたメールアドレスへ自動送信しております。", "").strip(),
        "",
        f"運営事務局：{conf['name']}",
        f"E-mail：{conf['reply_to_email'] or conf['office_email'] or ''}",
        "",
        "＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝",
        "今回のご回答内容",
        "＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝",
        *answer_lines,
        f"備考：{note or 'なし'}",
        "",
        *registration_lines,
        "回答日時：" + now,
        "",
        "ご登録内容・これまでの回答は、下記マイページよりご確認いただけます。",
        person_url(conf["code"], token),
    ]
    body = "\n".join(body_parts).strip() + "\n"

    sender_name = conf["sender_name"] or f"{conf['name']} 運営事務局"
    reply_to = conf["reply_to_email"] or conf["office_email"] or ""
    office = reply_to
    results = []
    if not email:
        results.append(("warning", "回答は保存されましたが、回答者のメールアドレスがないため自動返信メールを送信できませんでした。"))
        return results
    try:
        send_email_message(email, subject, body, sender_name, reply_to, bcc_email=office)
        if office:
            results.append(("success", f"自動返信メールを {email} に送信し、同じ内容を事務局（{office}）へBCC送信しました。"))
        else:
            results.append(("success", f"自動返信メールを {email} に送信しました。"))
    except Exception as e:
        results.append(("warning", f"回答は保存されましたが、メールを送信できませんでした：{e}"))
    return results


def set_flash(messages):
    st.session_state["response_flash"] = messages


def show_flash():
    messages = st.session_state.pop("response_flash", None)
    if not messages:
        return
    for kind, text in messages:
        if kind == "success":
            st.success(text)
        elif kind == "warning":
            st.warning(text)
        else:
            st.info(text)


def save_upload(conf_code, token, upload):
    if not upload:
        return "", ""
    ATTACH_DIR.mkdir(parents=True, exist_ok=True)
    safe = re.sub(r"[^0-9A-Za-z._\-ぁ-んァ-ヶ一-龠]", "_", upload.name)
    path = ATTACH_DIR / f"{conf_code}_{token[:8]}_{datetime.now().strftime('%Y%m%d%H%M%S')}_{safe}"
    path.write_bytes(upload.getvalue())
    return upload.name, str(path)


def request_summary_card(r):
    ans = effective_answer(r)
    badge_cls = "yes" if ans == "諾" else "no" if ans == "否" else "pending"
    badge_text = "承諾済み" if ans == "諾" else "辞退済み" if ans == "否" else "今回ご回答ください"
    card_cls = "done" if ans else "pending"
    st.markdown(f"""
    <div class="card {card_cls}">
      <span class="badge {badge_cls}">{badge_text}</span>
      <div style="font-size:1.08rem;font-weight:750">{r['session_name']}　｜　{r['role']}</div>
      <div style="margin:.28rem 0 .65rem;color:#344b5c">{r['theme'] or ''}</div>
      <div class="meta">
        <div class="k">日時</div><div>{r['schedule'] or '―'}</div>
        <div class="k">回答期限</div><div>{r['deadline'] or '―'}</div>
        <div class="k">依頼状況</div><div>{r['request_sent'] or '―'}</div>
      </div>
    </div>
    """, unsafe_allow_html=True)


def locked_profile(conf, p, row):
    st.markdown("### ご登録情報")
    st.caption("初回回答時に登録済みのため、今回は変更できません。変更が必要な場合は備考欄にご記入ください。")
    st.text_input("氏名", value=p["name"] or row["name"], disabled=True)
    if conf["ask_furigana"]:
        st.text_input("ふりがな", value=p["furigana"] or "", disabled=True)
    st.text_input("メールアドレス", value=p["email"] or row["email"], disabled=True)
    if conf["ask_membership"]:
        st.text_input(conf["membership_label"] or "会員・非会員", value=p["membership"] or "", disabled=True)
    if conf["ask_mobile"]:
        st.text_input("当日の緊急連絡先（携帯電話番号）", value=p["mobile"] or "", disabled=True)
    if conf["ask_correction"]:
        st.text_area("氏名・ご所属等の修正依頼", value=p["correction"] or "", disabled=True)
    if conf["invitation_enabled"]:
        st.text_input(conf["invitation_label"] or "派遣依頼状（招聘状）", value=p["invitation"] or "未登録", disabled=True)
        if p["invitation"] == "必要":
            st.text_input("所属長の所属機関名", value=p["leader_org"] or "", disabled=True)
            st.text_input("所属長の役職", value=p["leader_title"] or "", disabled=True)
            st.text_input("所属長の氏名", value=p["leader_name"] or "", disabled=True)
            st.text_area("指定様式での発行、web申請のご希望", value=p["special_request"] or "", disabled=True)
            st.text_input("指定様式アップロード", value=p["upload_name"] or "", disabled=True)


def initial_form(conf, token, rows, pending):
    first = rows[0]
    st.markdown("### 今回ご回答いただくご依頼")
    st.markdown('<div class="notice">まず、今回のご依頼について諾否をご回答ください。初回のみ、その下で基本情報をご登録いただきます。</div>', unsafe_allow_html=True)
    for r in pending:
        request_summary_card(r)

    st.markdown("#### 諾否のご回答")
    answers, declines = {}, {}
    for r in pending:
        label = f"{r['session_name']}｜{r['role']}"
        answers[r["request_id"]] = st.radio(label, ["選択してください", "承諾する", "辞退する"], horizontal=True, key=f"ans_{r['request_id']}")
        declines[r["request_id"]] = st.text_area(f"辞退理由（{label}）※辞退の場合のみ", key=f"dec_{r['request_id']}")

    note = st.text_area("備考", placeholder="ご連絡事項がございましたら、ご入力ください。", key="first_note")

    st.markdown("### ご登録情報")
    st.caption("初回のみご登録ください。次回以降は表示のみとなり、変更できません。変更が必要な場合は備考欄にご記入ください。")
    name = st.text_input("氏名", value=first["name"] or "", key="first_name")
    furigana = st.text_input("ふりがな", key="first_furigana") if conf["ask_furigana"] else ""
    email = st.text_input("メールアドレス", value=first["email"] or "", key="first_email")
    membership = ""
    if conf["ask_membership"]:
        membership = st.radio(conf["membership_label"] or "会員・非会員", ["選択してください", "会員", "非会員", "入会申請中、入会予定"], key="first_membership")
    mobile = st.text_input("当日の緊急連絡先（携帯電話番号）", key="first_mobile") if conf["ask_mobile"] else ""

    correction = ""
    if conf["ask_correction"]:
        correction = st.text_area("氏名・ご所属等の修正依頼", placeholder="氏名・所属等に修正がある場合のみご入力ください。", key="first_correction")

    any_yes = any(v == "承諾する" for v in answers.values())
    invitation = leader_org = leader_title = leader_name = special_request = ""
    upload = None
    if conf["invitation_enabled"] and any_yes:
        st.markdown("#### " + (conf["invitation_label"] or "派遣依頼状（招聘状）の発行について"))
        invitation = st.radio("発行の要否", ["選択してください", "必要", "不要"], horizontal=True, key="first_invitation")
        if invitation == "必要":
            leader_org = st.text_input("所属長の所属機関名", key="first_leader_org")
            leader_title = st.text_input("所属長の役職", key="first_leader_title")
            leader_name = st.text_input("所属長の氏名", key="first_leader_name")
            special_request = st.text_area("指定様式での発行、web申請のご希望", key="first_special")
            upload = st.file_uploader("指定様式アップロード", type=["pdf","doc","docx","xls","xlsx","xlsm","jpg","jpeg","png"], key="first_upload")

    submitted = st.button("回答を登録する", type="primary", use_container_width=True, key="first_submit")
    if not submitted:
        return

    errors = []
    if not name.strip(): errors.append("氏名を入力してください。")
    if conf["ask_furigana"] and not furigana.strip(): errors.append("ふりがなを入力してください。")
    if not email.strip() or "@" not in email: errors.append("メールアドレスを入力してください。")
    if conf["ask_membership"] and membership == "選択してください": errors.append("会員区分を選択してください。")
    if conf["ask_mobile"] and not mobile.strip(): errors.append("携帯電話番号を入力してください。")
    for r in pending:
        if answers[r["request_id"]] == "選択してください": errors.append(f"{r['session_name']}（{r['role']}）の諾否を選択してください。")
    if conf["invitation_enabled"] and any_yes:
        if invitation == "選択してください": errors.append("招聘状等の要否を選択してください。")
        if invitation == "必要" and not (leader_org.strip() and leader_title.strip() and leader_name.strip()):
            errors.append("招聘状等が必要な場合は、所属長の所属機関名・役職・氏名を入力してください。")
    if errors:
        for e in errors: st.error(e)
        return

    upload_name, upload_path = save_upload(conf["code"], token, upload) if conf["invitation_enabled"] and invitation == "必要" else ("", "")
    upsert_profile({
        "conference_id": conf["id"], "token": token, "name": name.strip(), "email": email.strip(), "affiliation": first["affiliation"] or "",
        "furigana": furigana.strip(), "membership": membership if conf["ask_membership"] else "", "mobile": mobile.strip(), "correction": correction.strip(),
        "invitation": invitation if conf["invitation_enabled"] and any_yes else "", "leader_org": leader_org.strip(), "leader_title": leader_title.strip(),
        "leader_name": leader_name.strip(), "special_request": special_request.strip(), "upload_name": upload_name, "upload_path": upload_path,
        "registered_at": datetime.now().isoformat(timespec="seconds"), "source": "system"
    }, overwrite=True)
    saved_answers = {}
    for r in pending:
        ans = "諾" if answers[r["request_id"]] == "承諾する" else "否"
        save_response(conf["id"], r["request_id"], ans, declines[r["request_id"]], note)
        saved_answers[r["request_id"]] = ans
    profile = get_profile(conf["id"], token)
    mails = send_response_notifications(conf, token, pending, saved_answers, note, profile, first_registration=True, decline_map=declines)
    set_flash([("success", "回答を登録しました。次回以降は、諾否と備考のみご回答いただけます。"), *mails])
    st.rerun()


def subsequent_form(conf, p, rows, pending):
    if pending:
        st.markdown("### 今回ご回答いただくご依頼")
        st.markdown('<div class="notice">今回入力できるのは「諾否」と「備考」のみです。</div>', unsafe_allow_html=True)
        for r in pending:
            request_summary_card(r)
        with st.form("subsequent_response"):
            answers, declines = {}, {}
            for r in pending:
                label = f"{r['session_name']}｜{r['role']}"
                answers[r["request_id"]] = st.radio(label, ["選択してください", "承諾する", "辞退する"], horizontal=True, key=f"ans2_{r['request_id']}")
                declines[r["request_id"]] = st.text_area(f"辞退理由（{label}）※辞退の場合のみ", key=f"dec2_{r['request_id']}")
            note = st.text_area("備考", placeholder="ご連絡事項がございましたら、ご入力ください。")
            submitted = st.form_submit_button("今回の回答を登録する", type="primary", use_container_width=True)
            if submitted:
                if any(answers[r["request_id"]] == "選択してください" for r in pending):
                    st.error("すべてのご依頼について、承諾または辞退を選択してください。")
                    return
                saved_answers = {}
                for r in pending:
                    ans = "諾" if answers[r["request_id"]] == "承諾する" else "否"
                    save_response(conf["id"], r["request_id"], ans, declines[r["request_id"]], note)
                    saved_answers[r["request_id"]] = ans
                mails = send_response_notifications(conf, p["token"], pending, saved_answers, note, p, first_registration=False, decline_map=declines)
                set_flash([("success", "回答を登録しました。"), *mails])
                st.rerun()
    else:
        st.success("現在、新たにご回答いただく依頼はありません。")

    locked_profile(conf, p, rows[0])


def user_page(conf_code, token):
    conf = get_conference_by_code(conf_code)
    if not conf or not conf["active"]:
        hero(APP_NAME, "ご依頼内容の確認・諾否回答")
        st.error("この学会の回答ページは現在利用できません。")
        return
    rows = get_person_requests(conf["id"], token)
    if not rows:
        hero(conf["name"], "ご依頼内容の確認・諾否回答")
        st.error("このURLに該当するご依頼が見つかりません。")
        return
    hero(conf["name"], f"{conf['dates'] or ''}　{conf['venue'] or ''}")
    st.markdown(f'<div class="person">{rows[0]["name"]} 先生</div><div class="subtle">{rows[0]["affiliation"]}</div>', unsafe_allow_html=True)
    show_flash()
    pending = [r for r in rows if not effective_answer(r)]
    done = [r for r in rows if effective_answer(r)]
    p = get_profile(conf["id"], token)
    if p:
        subsequent_form(conf, p, rows, pending)
    elif pending:
        initial_form(conf, token, rows, pending)
    else:
        st.info("既に回答済みですが、初回登録情報が見つかりません。事務局へお問い合わせください。")
    if done:
        st.markdown("### これまでに回答済みのご依頼")
        st.caption("これまでにご回答いただいた内容です。")
        for r in done:
            request_summary_card(r)


def make_export(conf):
    rows = all_rows(conf["id"])
    out = BytesIO(); book = xlsxwriter.Workbook(out, {"in_memory":True})
    ws = book.add_worksheet("回答一覧"); pp = book.add_worksheet("先生情報")
    head = book.add_format({"bold":True,"bg_color":"#1E526D","font_color":"#FFFFFF","border":1})
    headers = ["元Excel行","依頼ID","氏名","所属","メール","セッション名","テーマ","役割","日時","依頼状況","既存諾否","新回答","現在の諾否","回答日時","辞退理由","備考","マイページURL"]
    for c,h in enumerate(headers): ws.write(0,c,h,head)
    for rn,r in enumerate(rows,1):
        url=person_url(conf["code"], r["token"])
        vals=[r['source_row'],r['request_id'],r['name'],r['affiliation'],r['email'],r['session_name'],r['theme'],r['role'],r['schedule'],r['request_sent'],r['source_answer'],r['new_answer'] or '',effective_answer(r) or '未回答',r['responded_at'] or '',r['decline_reason'] or '',r['note'] or '',url]
        for c,v in enumerate(vals): ws.write(rn,c,v)
    pheaders=["氏名","ふりがな","メール","所属","会員区分","携帯電話","修正依頼","招聘状等","所属長所属機関","所属長役職","所属長氏名","指定様式等","登録日時","登録元"]
    for c,h in enumerate(pheaders): pp.write(0,c,h,head)
    for rn,p in enumerate(all_people(conf['id']),1):
        vals=[p['name'],p['furigana'],p['email'],p['affiliation'],p['membership'],p['mobile'],p['correction'],p['invitation'],p['leader_org'],p['leader_title'],p['leader_name'],p['special_request'],p['registered_at'],p['source']]
        for c,v in enumerate(vals): pp.write(rn,c,v or '')
    ws.set_column(0,16,20); ws.set_column(16,16,56); pp.set_column(0,13,24)
    ws.freeze_panes(1,0); pp.freeze_panes(1,0); book.close(); out.seek(0)
    return out.getvalue()


def make_url_export(conf, mode="all"):
    rows = all_rows(conf["id"])
    people = {p["token"]: p for p in all_people(conf["id"])}
    grouped = {}
    for r in rows:
        g = grouped.setdefault(r["token"], {
            "name": r["name"], "affiliation": r["affiliation"], "email": r["email"],
            "total": 0, "answered": 0, "pending": 0, "additional": False,
        })
        g["total"] += 1
        if effective_answer(r):
            g["answered"] += 1
        else:
            g["pending"] += 1
    for token, g in grouped.items():
        p = people.get(token)
        # A profile means the person has already completed initial registration.
        # Pending requests after that are treated as additional/current requests.
        g["additional"] = bool(p and p["registered_at"] and g["pending"] > 0)

    selected = []
    for token, g in grouped.items():
        if mode == "pending" and g["pending"] == 0:
            continue
        if mode == "additional" and not g["additional"]:
            continue
        selected.append((token, g))

    out = BytesIO()
    book = xlsxwriter.Workbook(out, {"in_memory": True})
    ws = book.add_worksheet("先生別URL")
    head = book.add_format({"bold": True, "bg_color": "#1E526D", "font_color": "#FFFFFF", "border": 1})
    headers = ["氏名", "所属", "メールアドレス", "学会コード", "専用URL", "依頼件数", "回答済み件数", "未回答件数", "追加依頼あり"]
    for c, h in enumerate(headers):
        ws.write(0, c, h, head)
    for rn, (token, g) in enumerate(selected, 1):
        vals = [
            g["name"], g["affiliation"], g["email"], conf["code"], person_url(conf["code"], token),
            g["total"], g["answered"], g["pending"], "○" if g["additional"] else "",
        ]
        for c, v in enumerate(vals):
            ws.write(rn, c, v)
    ws.set_column(0, 0, 18)
    ws.set_column(1, 2, 30)
    ws.set_column(3, 3, 14)
    ws.set_column(4, 4, 72)
    ws.set_column(5, 8, 14)
    ws.freeze_panes(1, 0)
    book.close()
    out.seek(0)
    return out.getvalue(), len(selected)


def make_reminder_export(conf):
    rows = all_rows(conf["id"])
    grouped = {}
    for r in rows:
        if effective_answer(r):
            continue
        g = grouped.setdefault(r["token"], {"name":r["name"],"affiliation":r["affiliation"],"email":r["email"],"items":[]})
        g["items"].append(f"{r['session_name']}｜{r['role']}")
    out = BytesIO(); book = xlsxwriter.Workbook(out, {"in_memory": True}); ws = book.add_worksheet("催促メール用")
    head = book.add_format({"bold":True,"bg_color":"#1E526D","font_color":"#FFFFFF","border":1})
    headers=["氏名","所属","メールアドレス","未回答依頼","専用URL","件名","メール本文"]
    for c,h in enumerate(headers): ws.write(0,c,h,head)
    for rn,(token,g) in enumerate(grouped.items(),1):
        items="／".join(g["items"])
        subject=f"【ご回答のお願い】{conf['name']} 登壇諾否について"
        body=f"{g['name']} 先生\n\nお世話になっております。\n{conf['name']} 運営事務局でございます。\n\n下記ご依頼につきまして、現在ご回答を確認できておりません。\n{items}\n\n恐れ入りますが、下記専用ページよりご回答くださいますようお願い申し上げます。\n{person_url(conf['code'],token)}\n\n何卒よろしくお願い申し上げます。"
        vals=[g["name"],g["affiliation"],g["email"],items,person_url(conf["code"],token),subject,body]
        for c,v in enumerate(vals): ws.write(rn,c,v)
    ws.set_column(0,2,28); ws.set_column(3,3,48); ws.set_column(4,4,58); ws.set_column(5,5,42); ws.set_column(6,6,80)
    ws.freeze_panes(1,0); book.close(); out.seek(0)
    return out.getvalue(), len(grouped)


def conference_settings_form(conf=None, key_prefix="conf"):
    is_new = conf is None
    code = st.text_input("学会コード（英数字・ハイフン）", value="" if is_new else conf["code"], disabled=not is_new, placeholder="例：JKRA17", key=f"{key_prefix}_code")
    name = st.text_input("学会名", value="" if is_new else conf["name"], key=f"{key_prefix}_name")
    dates = st.text_input("会期", value="" if is_new else conf["dates"] or "", placeholder="例：2027年3月20日（土）・21日（日）", key=f"{key_prefix}_dates")
    venue = st.text_input("会場", value="" if is_new else conf["venue"] or "", key=f"{key_prefix}_venue")
    deadline = st.text_input("標準回答期限", value="" if is_new else conf["reply_deadline"] or "", key=f"{key_prefix}_deadline")
    active = st.checkbox("回答ページを有効にする", value=True if is_new else bool(conf["active"]), key=f"{key_prefix}_active")
    st.markdown("#### 初回だけ聞く項目")
    c1,c2,c3 = st.columns(3)
    ask_furigana = c1.checkbox("ふりがな", value=True if is_new else bool(conf["ask_furigana"]), key=f"{key_prefix}_furigana")
    ask_membership = c2.checkbox("会員区分", value=False if is_new else bool(conf["ask_membership"]), key=f"{key_prefix}_membership")
    ask_mobile = c3.checkbox("緊急連絡先", value=False if is_new else bool(conf["ask_mobile"]), key=f"{key_prefix}_mobile")
    ask_correction = st.checkbox("氏名・所属等の修正依頼", value=True if is_new else bool(conf["ask_correction"]), key=f"{key_prefix}_correction")
    membership_label = st.text_input("会員区分の表示名", value="会員・非会員" if is_new else conf["membership_label"] or "会員・非会員", disabled=not ask_membership, key=f"{key_prefix}_membership_label")
    st.markdown("#### 学会固有の追加項目")
    invitation_enabled = st.checkbox("招聘状／派遣依頼状セットを使う", value=False if is_new else bool(conf["invitation_enabled"]), key=f"{key_prefix}_invitation_enabled")
    invitation_label = st.text_input("表示名", value="派遣依頼状（招聘状）の発行について" if is_new else conf["invitation_label"] or "派遣依頼状（招聘状）の発行について", disabled=not invitation_enabled, key=f"{key_prefix}_invitation_label")

    st.markdown("#### 回答受付メール")
    st.caption("回答登録時は、回答者へ自動返信し、同じメールを運営事務局へBCC送信します。ON/OFF設定はありません。")
    reply_to_email = st.text_input("運営事務局メールアドレス（Reply-To・BCC先）", value="" if is_new else (conf["reply_to_email"] or conf["office_email"] or ""), placeholder="例：jsrr17@gakkai.co.jp", key=f"{key_prefix}_reply_to")
    sender_name = st.text_input("メール差出人表示名", value=(name + " 運営事務局") if is_new and name else (conf["sender_name"] or (name + " 運営事務局")), key=f"{key_prefix}_sender_name")
    default_subject = f"【{name}】ご回答を受け付けました" if name else DEFAULT_AUTO_SUBJECT
    auto_reply_subject = st.text_input("自動返信メール 件名", value=default_subject if is_new else conf["auto_reply_subject"] or default_subject, key=f"{key_prefix}_auto_subject")
    st.caption("氏名・学会名・諾否・辞退理由・備考・初回登録情報・回答日時・マイページURLはシステムが自動でメールに追加します。差込タグは不要です。")
    auto_reply_body = st.text_area("回答内容より前に表示する案内文", value=DEFAULT_AUTO_INTRO if is_new else conf["auto_reply_body"] or DEFAULT_AUTO_INTRO, height=260, key=f"{key_prefix}_auto_body")
    values = {"code":code.strip(),"name":name.strip(),"dates":dates.strip(),"venue":venue.strip(),"reply_deadline":deadline.strip(),"active":active,
              "ask_furigana":ask_furigana,"ask_membership":ask_membership,"membership_label":membership_label.strip(),"ask_mobile":ask_mobile,
              "ask_correction":ask_correction,"invitation_enabled":invitation_enabled,"invitation_label":invitation_label.strip(),
              "auto_reply_enabled":True,"office_notify_enabled":True,"office_email":reply_to_email.strip(),
              "sender_name":sender_name.strip(),"reply_to_email":reply_to_email.strip(),"auto_reply_subject":auto_reply_subject,
              "auto_reply_body":auto_reply_body,"office_subject":"","office_body":""}
    return values


def admin_page():
    hero(APP_NAME, "1つの管理画面で複数の学会を作成・管理できます")
    if not ADMIN_PASSWORD:
        st.error("管理者パスワードが未設定です。Streamlit Secrets に ADMIN_PASSWORD を設定してください。")
        st.code('ADMIN_PASSWORD = "十分に長いパスワード"\nTOKEN_SECRET = "ランダムな長い文字列"', language="toml")
        return
    if not st.session_state.get("admin_ok"):
        pw = st.text_input("管理者パスワード", type="password")
        if st.button("ログイン", type="primary"):
            if hmac.compare_digest(pw.encode("utf-8"), ADMIN_PASSWORD.encode("utf-8")):
                st.session_state.admin_ok=True; st.rerun()
            else: st.error("パスワードが違います。")
        return

    confs = list_conferences()
    tabs = st.tabs(["学会一覧", "＋ 新しい学会", "学会管理", "運用設定"])
    with tabs[0]:
        if not confs:
            st.info("まだ学会がありません。『＋ 新しい学会』から作成してください。")
        else:
            data=[]
            for c in confs:
                rows=all_rows(c['id']); yes=sum(effective_answer(r)=='諾' for r in rows); no=sum(effective_answer(r)=='否' for r in rows); pending=sum(not effective_answer(r) for r in rows)
                data.append({"コード":c['code'],"学会名":c['name'],"会期":c['dates'],"依頼件数":len(rows),"承諾":yes,"辞退":no,"未回答":pending,"招聘状セット":"ON" if c['invitation_enabled'] else "OFF","状態":"有効" if c['active'] else "停止"})
            st.dataframe(data,use_container_width=True,hide_index=True)

    with tabs[1]:
        values=conference_settings_form(key_prefix="new_conf")
        if st.button("この学会を作成",type="primary",use_container_width=True):
            if not values['code'] or not re.fullmatch(r"[A-Za-z0-9_-]+",values['code']): st.error("学会コードは半角英数字・_・-で入力してください。")
            elif not values['name']: st.error("学会名を入力してください。")
            elif get_conference_by_code(values['code']): st.error("同じ学会コードが既にあります。")
            else:
                create_conference(values); st.success("学会を作成しました。『学会管理』からExcelを取り込んでください。"); st.rerun()

    with tabs[2]:
        confs = list_conferences()
        if not confs:
            st.info("先に学会を作成してください。")
        else:
            labels={f"{c['code']}｜{c['name']}":c for c in confs}
            selected=st.selectbox("管理する学会",list(labels.keys()))
            conf=labels[selected]
            subtabs=st.tabs(["回答状況","Excel取込","先生別URL","学会設定","先生情報"])
            with subtabs[0]:
                rows=all_rows(conf['id']); data=[]
                for r in rows:
                    data.append({"元Excel行":r['source_row'],"氏名":r['name'],"セッション":r['session_name'],"役割":r['role'],"既存":r['source_answer'] or '',"新回答":r['new_answer'] or '',"現在":effective_answer(r) or '未回答',"依頼状況":r['request_sent'],"回答日時":r['responded_at'] or ''})
                st.dataframe(data,use_container_width=True,hide_index=True)
                if rows:
                    st.download_button("回答一覧をExcelでダウンロード",make_export(conf),f"{conf['code']}_諾否回答一覧.xlsx","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True)
            with subtabs[1]:
                st.write("この学会の指定演題・登壇者Excelを取り込みます。氏名／セッション／役割などは見出し名から自動判定します。")
                up=st.file_uploader("Excel（.xlsm / .xlsx）",type=["xlsm","xlsx"],key=f"up_{conf['id']}")
                if up:
                    try:
                        _, sname, headers, _ = read_workbook_sheet(up.getvalue())
                        mapping,_=auto_mapping(headers)
                        st.caption(f"読み取り対象シート：{sname}")
                        st.dataframe(mapping_display(headers,mapping),use_container_width=True,hide_index=True)
                        if st.button("このExcelを取り込む",type="primary",key=f"import_{conf['id']}"):
                            records,profiles,_,_=import_master(conf,up.getvalue(),sname)
                            st.success(f"依頼 {len(records)}件、既存の初回登録情報 {len(profiles)}名分を取り込みました。")
                            st.rerun()
                    except Exception as e:
                        st.error(f"Excelの読み取りに失敗しました：{e}")
            with subtabs[2]:
                rows=all_rows(conf['id']); seen={}; urls=[]
                for r in rows:
                    if r['token'] not in seen:
                        seen[r['token']]=1
                        person_rows=[x for x in rows if x['token']==r['token']]
                        pending_count=sum(not effective_answer(x) for x in person_rows)
                        answered_count=sum(bool(effective_answer(x)) for x in person_rows)
                        p=get_profile(conf['id'], r['token'])
                        additional=bool(p and p['registered_at'] and pending_count>0)
                        urls.append({
                            "氏名":r['name'],"所属":r['affiliation'],"メール":r['email'],
                            "未回答":pending_count,"回答済み":answered_count,
                            "追加依頼":"○" if additional else "",
                            "専用URL":person_url(conf["code"], r["token"])
                        })
                st.caption(f"専用URLの本体部分は自動取得しています：{get_base_url()}")
                st.dataframe(urls,use_container_width=True,hide_index=True)
                if rows:
                    st.markdown("#### URL一覧をExcelでダウンロード")
                    c_all,c_pending,c_add=st.columns(3)
                    data_all,n_all=make_url_export(conf,"all")
                    data_pending,n_pending=make_url_export(conf,"pending")
                    data_add,n_add=make_url_export(conf,"additional")
                    c_all.download_button(
                        f"全員（{n_all}名）", data_all, f"{conf['code']}_先生別URL_全員.xlsx",
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True, key=f"url_all_{conf['id']}"
                    )
                    c_pending.download_button(
                        f"未回答あり（{n_pending}名）", data_pending, f"{conf['code']}_先生別URL_未回答.xlsx",
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True, key=f"url_pending_{conf['id']}"
                    )
                    c_add.download_button(
                        f"追加依頼あり（{n_add}名）", data_add, f"{conf['code']}_先生別URL_追加依頼.xlsx",
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True, key=f"url_add_{conf['id']}"
                    )
                    reminder_data, reminder_count = make_reminder_export(conf)
                    st.download_button(
                        f"未回答者 催促メール用Excel（{reminder_count}名）", reminder_data, f"{conf['code']}_未回答者_催促メール用.xlsx",
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True, key=f"reminder_{conf['id']}"
                    )
            with subtabs[3]:
                values=conference_settings_form(conf, key_prefix=f"edit_{conf['id']}")
                if st.button("設定を保存",type="primary",key=f"save_conf_{conf['id']}"):
                    update_conference(conf['id'],values); st.success("保存しました。"); st.rerun()

                st.divider()
                with st.expander("⚠ この学会を削除", expanded=False):
                    st.warning("この操作は取り消せません。この学会に紐づく先生情報・依頼・回答もすべて削除します。")
                    confirm_code = st.text_input(
                        f"確認のため学会コード「{conf['code']}」を入力してください",
                        key=f"delete_code_{conf['id']}"
                    )
                    delete_ok = confirm_code.strip() == conf['code']
                    if st.button(
                        "この学会を完全に削除",
                        type="secondary",
                        disabled=not delete_ok,
                        key=f"delete_conf_{conf['id']}",
                        use_container_width=True,
                    ):
                        deleted_name = conf['name']
                        deleted_code = conf['code']
                        delete_conference(conf['id'])
                        st.success(f"{deleted_code}｜{deleted_name} を削除しました。")
                        st.rerun()
            with subtabs[4]:
                ps=[]
                for p in all_people(conf['id']):
                    ps.append({"氏名":p['name'],"ふりがな":p['furigana'],"メール":p['email'],"会員区分":p['membership'],"携帯電話":p['mobile'],"招聘状等":p['invitation'],"登録元":p['source'],"登録日時":p['registered_at']})
                st.dataframe(ps,use_container_width=True,hide_index=True)

    with tabs[3]:
        st.markdown("#### 本番運用について")
        st.warning("現在の保存先はSQLiteなので、Streamlit Community Cloudでは再起動・再デプロイ時にデータが失われる可能性があります。本番回答を開始する前に永続保存へ切り替えてください。")
        st.markdown("新しい学会は、この管理画面で『＋ 新しい学会』→設定→Excel取込だけで追加できます。")
        st.markdown("招聘状セットなどの質問項目、回答者への自動返信、事務局への回答通知は学会ごとにON/OFFできます。")
        st.markdown("#### 共通SMTP設定")
        if smtp_ready():
            st.success(f"SMTP設定済み：{SMTP_HOST}:{SMTP_PORT} / {SMTP_SECURITY}")
        else:
            st.error("SMTP設定が未完了のため、メールは送信されません。Streamlitの Settings → Secrets に下記を設定してください。")
        st.code('''SMTP_HOST = "mail.example.jp"
SMTP_PORT = "587"
SMTP_USERNAME = "your-account@example.jp"
SMTP_PASSWORD = "メール送信用パスワード"
SMTP_SECURITY = "starttls"
SMTP_FROM_EMAIL = "your-account@example.jp"''', language="toml")
        st.caption("SMTP_SECURITY は starttls / ssl / none のいずれか。メールアカウントの仕様に合わせて設定します。")
        test_to = st.text_input("テスト送信先メールアドレス", key="smtp_test_to")
        if st.button("SMTPテストメールを送信", disabled=not smtp_ready(), key="smtp_test_btn"):
            try:
                send_email_message(test_to.strip(), "【テスト】登壇諾否マイページ メール送信確認", "このメールが届けば、SMTP設定は正常です。", "登壇諾否マイページ")
                st.success(f"テストメールを {test_to.strip()} に送信しました。")
            except Exception as e:
                st.error(f"テストメールを送信できませんでした：{e}")


init_db()
params=st.query_params
if params.get("admin") == "1":
    admin_page()
elif params.get("c") and params.get("token"):
    user_page(params.get("c"), params.get("token"))
else:
    hero(APP_NAME, "メールに記載された先生専用URLからアクセスしてください")
    st.info("管理画面は ?admin=1 です。")
