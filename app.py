import os
import re
import sqlite3
import hashlib
import hmac
from io import BytesIO
from pathlib import Path
from datetime import datetime, date

import streamlit as st
from python_calamine import CalamineWorkbook
import xlsxwriter

APP_TITLE = "第17回日本腎臓リハビリテーション学会学術集会｜座長・演者マイページ"
DB_PATH = os.getenv("JKRA17_DB_PATH", "jkra17_yesno.db")
DEFAULT_MASTER = Path(__file__).with_name("17腎リハ　指定演題（260909）.xlsm")
ADMIN_PASSWORD = os.getenv("JKRA17_ADMIN_PASSWORD", "jkra17-demo")
TOKEN_SECRET = os.getenv("JKRA17_TOKEN_SECRET", "change-this-secret-before-production").encode("utf-8")
BASE_URL = os.getenv("JKRA17_BASE_URL", "http://localhost:8501")
ATTACH_DIR = Path(os.getenv("JKRA17_ATTACH_DIR", "attachments"))

st.set_page_config(page_title="諾否マイページ", page_icon="✅", layout="centered")

st.markdown("""
<style>
.block-container {max-width: 980px; padding-top: 1.7rem; padding-bottom: 4rem;}
.hero {padding: 1.35rem 1.5rem; border-radius: 18px; background: linear-gradient(135deg,#15334b,#246681); color:#fff; margin-bottom: 1.2rem;}
.hero h1 {font-size:1.55rem; margin:0 0 .25rem 0; color:#fff;}
.hero p {margin:0; opacity:.90;}
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
.locked-title {margin-top:1.2rem; font-weight:750; color:#344b5c;}
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


def person_token(email, name):
    key = (clean(email).lower() or clean(name)).encode("utf-8")
    return hmac.new(TOKEN_SECRET, key, hashlib.sha256).hexdigest()[:32]


def request_id(row_no, seq_no, session, role, name):
    stable = clean(seq_no) or f"row-{row_no}"
    raw = f"JKRA17|{stable}|{clean(session)}|{clean(role)}|{clean(name)}"
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()[:24]


def connect():
    con = sqlite3.connect(DB_PATH)
    con.row_factory = sqlite3.Row
    return con


def init_db():
    with connect() as con:
        con.execute("""
        CREATE TABLE IF NOT EXISTS requests (
            request_id TEXT PRIMARY KEY,
            source_row INTEGER NOT NULL,
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
            imported_at TEXT NOT NULL
        )""")
        con.execute("""
        CREATE TABLE IF NOT EXISTS responses (
            request_id TEXT PRIMARY KEY,
            answer TEXT NOT NULL,
            decline_reason TEXT,
            note TEXT,
            responded_at TEXT NOT NULL,
            FOREIGN KEY(request_id) REFERENCES requests(request_id)
        )""")
        con.execute("""
        CREATE TABLE IF NOT EXISTS people (
            token TEXT PRIMARY KEY,
            name TEXT,
            email TEXT,
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
            source TEXT
        )""")
        # 旧版DBを同じフォルダで使った場合も、新しい列を追加して移行できるようにする。
        cols = {r[1] for r in con.execute("PRAGMA table_info(people)").fetchall()}
        for col in ("furigana", "membership", "mobile"):
            if col not in cols:
                con.execute(f"ALTER TABLE people ADD COLUMN {col} TEXT")


def read_designated_sheet(file_bytes):
    wb = CalamineWorkbook.from_filelike(BytesIO(file_bytes))
    sheet = wb.get_sheet_by_name("指定演題")
    rows = sheet.to_python(skip_empty_area=False)
    if not rows:
        raise ValueError("「指定演題」シートにデータがありません。")

    headers = [clean(x).replace("\n", "") for x in rows[0]]
    # 同名見出し（例：メールアドレス、諾否）が複数あるため、
    # 最初の列＝指定演題マスター、最後の列＝旧Form Mailer回答として明示的に扱う。
    indices = {}
    for i, h in enumerate(headers):
        if h:
            indices.setdefault(h, []).append(i)
    required = ["セッション名", "役割", "元）氏名", "諾否", "諾否締切", "依頼送信"]
    missing = [x for x in required if x not in indices]
    if missing:
        raise ValueError("必要な列が見つかりません: " + "、".join(missing))

    def val(row, header, fallback="", occurrence="first"):
        locs = indices.get(header, [])
        if not locs:
            return fallback
        i = locs[-1] if occurrence == "last" else locs[0]
        return row[i] if i < len(row) else fallback

    records = []
    legacy_profiles = {}
    for excel_row, row in enumerate(rows[1:], start=2):
        name = clean(val(row, "元）氏名"))
        if not name:
            continue
        session = clean(val(row, "セッション名"))
        role = clean(val(row, "役割"))
        if not session or not role:
            continue

        affiliation = clean(val(row, "所属")) or clean(val(row, "元）所属"))
        email = clean(val(row, "メールアドレス", occurrence="first"))
        theme = clean(val(row, "元）セッションテーマ（キャッチーなテーマ名）"))
        schedule = clean(val(row, "候補日時"))
        source_answer = normalize_answer(val(row, "諾否", occurrence="first"))
        deadline = clean(val(row, "諾否締切"))
        sent = clean(val(row, "依頼送信"))
        seq = clean(val(row, "スプシセッション順No."))
        token = person_token(email, name)
        rid = request_id(excel_row, seq, session, role, name)

        records.append({
            "request_id": rid, "source_row": excel_row, "source_seq": seq, "token": token,
            "name": name, "affiliation": affiliation, "email": email, "session_name": session,
            "theme": theme, "role": role, "schedule": schedule, "source_answer": source_answer,
            "deadline": deadline, "request_sent": sent,
        })

        # 旧Form Mailerの回答が入っている行は「初回登録済み」として先生情報を取り込む。
        legacy_receipt = clean(val(row, "受付番号"))
        legacy_sent = clean(val(row, "送信日時"))
        legacy_form_name = clean(val(row, "氏名"))
        if legacy_receipt or legacy_sent or legacy_form_name:
            p = {
                "token": token,
                "name": legacy_form_name or name,
                "email": clean(val(row, "メールアドレス", occurrence="last")) or email,
                "furigana": clean(val(row, "ふりがな")),
                "membership": clean(val(row, "会員・非会員 （日本腎臓リハビリテーション学会）")),
                "mobile": clean(val(row, "携帯電話番号")),
                "correction": clean(val(row, "修正内容のご入力")),
                "invitation": normalize_invitation(val(row, "派遣依頼状（招聘状）")),
                "leader_org": clean(val(row, "所属長の所属機関名")),
                "leader_title": clean(val(row, "所属長の役職")),
                "leader_name": clean(val(row, "所属長の氏名")),
                "special_request": clean(val(row, "指定様式での発行、web申請のご希望")),
                "upload_name": clean(val(row, "派遣依頼状　指定様式アップロード")),
                "upload_path": "",
                "registered_at": legacy_sent or datetime.now().isoformat(timespec="seconds"),
                "source": "legacy_formmailer",
            }
            old = legacy_profiles.get(token)
            if not old or (p["registered_at"] >= old["registered_at"]):
                legacy_profiles[token] = p
    return records, list(legacy_profiles.values())


def upsert_profile(p, overwrite=False):
    with connect() as con:
        current = con.execute("SELECT * FROM people WHERE token=?", (p["token"],)).fetchone()
        if current and not overwrite and current["source"] == "system":
            return
        con.execute("""
        INSERT INTO people(token,name,email,furigana,membership,mobile,correction,invitation,leader_org,leader_title,leader_name,
                           special_request,upload_name,upload_path,registered_at,source)
        VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
        ON CONFLICT(token) DO UPDATE SET
            name=excluded.name,email=excluded.email,furigana=excluded.furigana,membership=excluded.membership,mobile=excluded.mobile,
            correction=excluded.correction,invitation=excluded.invitation,leader_org=excluded.leader_org,
            leader_title=excluded.leader_title,leader_name=excluded.leader_name,
            special_request=excluded.special_request,upload_name=excluded.upload_name,
            upload_path=CASE WHEN excluded.upload_path<>'' THEN excluded.upload_path ELSE people.upload_path END,
            registered_at=excluded.registered_at,source=excluded.source
        """, (p["token"], p.get("name", ""), p.get("email", ""), p.get("furigana", ""), p.get("membership", ""), p.get("mobile", ""), p.get("correction", ""),
              p.get("invitation", ""), p.get("leader_org", ""), p.get("leader_title", ""),
              p.get("leader_name", ""), p.get("special_request", ""), p.get("upload_name", ""),
              p.get("upload_path", ""), p.get("registered_at", ""), p.get("source", "system")))


def import_master(file_bytes):
    records, legacy_profiles = read_designated_sheet(file_bytes)
    now = datetime.now().isoformat(timespec="seconds")
    ids = {r["request_id"] for r in records}
    with connect() as con:
        existing = [x[0] for x in con.execute("SELECT request_id FROM requests").fetchall()]
        for rid in existing:
            if rid not in ids:
                con.execute("DELETE FROM responses WHERE request_id=?", (rid,))
                con.execute("DELETE FROM requests WHERE request_id=?", (rid,))
        for r in records:
            con.execute("""
            INSERT INTO requests
            (request_id, source_row, source_seq, token, name, affiliation, email, session_name,
             theme, role, schedule, source_answer, deadline, request_sent, imported_at)
            VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
            ON CONFLICT(request_id) DO UPDATE SET
              source_row=excluded.source_row,source_seq=excluded.source_seq,token=excluded.token,
              name=excluded.name,affiliation=excluded.affiliation,email=excluded.email,
              session_name=excluded.session_name,theme=excluded.theme,role=excluded.role,
              schedule=excluded.schedule,source_answer=excluded.source_answer,
              deadline=excluded.deadline,request_sent=excluded.request_sent,imported_at=excluded.imported_at
            """, (r["request_id"], r["source_row"], r["source_seq"], r["token"], r["name"],
                  r["affiliation"], r["email"], r["session_name"], r["theme"], r["role"],
                  r["schedule"], r["source_answer"], r["deadline"], r["request_sent"], now))
    for p in legacy_profiles:
        upsert_profile(p, overwrite=False)
    return records, legacy_profiles


def bootstrap():
    init_db()
    with connect() as con:
        count = con.execute("SELECT COUNT(*) FROM requests").fetchone()[0]
    if count == 0 and DEFAULT_MASTER.exists():
        import_master(DEFAULT_MASTER.read_bytes())


def effective_answer(row):
    return row["source_answer"] or row["new_answer"] or ""


def get_person_requests(token):
    with connect() as con:
        return con.execute("""
        SELECT r.*, s.answer AS new_answer, s.decline_reason, s.note, s.responded_at
        FROM requests r LEFT JOIN responses s ON s.request_id=r.request_id
        WHERE r.token=? ORDER BY r.source_row
        """, (token,)).fetchall()


def get_profile(token):
    with connect() as con:
        return con.execute("SELECT * FROM people WHERE token=?", (token,)).fetchone()


def save_response(request_id, answer, decline_reason, note):
    with connect() as con:
        source = con.execute("SELECT source_answer FROM requests WHERE request_id=?", (request_id,)).fetchone()
        if not source:
            raise ValueError("依頼が見つかりません。")
        if clean(source[0]):
            raise ValueError("この依頼は既に旧フォームで回答済みのため変更できません。")
        con.execute("""
        INSERT INTO responses(request_id, answer, decline_reason, note, responded_at)
        VALUES (?,?,?,?,?)
        ON CONFLICT(request_id) DO UPDATE SET answer=excluded.answer,
            decline_reason=excluded.decline_reason,note=excluded.note,responded_at=excluded.responded_at
        """, (request_id, answer, clean(decline_reason), clean(note), datetime.now().isoformat(timespec="seconds")))


def save_upload(token, upload):
    if not upload:
        return "", ""
    ATTACH_DIR.mkdir(parents=True, exist_ok=True)
    safe = re.sub(r"[^0-9A-Za-z._\-ぁ-んァ-ヶ一-龠]", "_", upload.name)
    path = ATTACH_DIR / f"{token[:10]}_{datetime.now().strftime('%Y%m%d%H%M%S')}_{safe}"
    path.write_bytes(upload.getvalue())
    return upload.name, str(path)


def all_rows():
    with connect() as con:
        return con.execute("""
        SELECT r.*, s.answer AS new_answer, s.decline_reason, s.note, s.responded_at
        FROM requests r LEFT JOIN responses s ON s.request_id=r.request_id
        ORDER BY r.source_row
        """).fetchall()


def all_people():
    with connect() as con:
        return con.execute("SELECT * FROM people ORDER BY name").fetchall()


def build_export_xlsx(rows):
    out = BytesIO()
    book = xlsxwriter.Workbook(out, {"in_memory": True})
    ws = book.add_worksheet("回答一覧")
    diff = book.add_worksheet("Excel反映用")
    pp = book.add_worksheet("先生情報")
    fmt_head = book.add_format({"bold": True, "bg_color": "#1E526D", "font_color": "#FFFFFF", "border": 1})
    fmt_yes = book.add_format({"bg_color": "#E7F6ED"})
    fmt_no = book.add_format({"bg_color": "#FDEAEA"})
    fmt_pending = book.add_format({"bg_color": "#FFF2D9"})

    headers = ["元Excel行","依頼ID","氏名","所属","メール","セッション名","テーマ","役割","日時","AF依頼送信","旧AD諾否","新システム回答","現在の諾否","回答日時","辞退理由","備考","マイページURL"]
    for c,h in enumerate(headers): ws.write(0,c,h,fmt_head)
    diff_headers = ["元Excel行","依頼ID","氏名","セッション名","役割","ADへ反映する値","回答日時","備考"]
    for c,h in enumerate(diff_headers): diff.write(0,c,h,fmt_head)
    drow = 1
    for rno,r in enumerate(rows,start=1):
        eff = effective_answer(r)
        vals = [r["source_row"],r["request_id"],r["name"],r["affiliation"],r["email"],r["session_name"],r["theme"],r["role"],r["schedule"],r["request_sent"],r["source_answer"],r["new_answer"] or "",eff or "未回答",r["responded_at"] or "",r["decline_reason"] or "",r["note"] or "",f"{BASE_URL}/?token={r['token']}"]
        fmt = fmt_yes if eff == "諾" else fmt_no if eff == "否" else fmt_pending
        for c,v in enumerate(vals): ws.write(rno,c,v,fmt if c==12 else None)
        if r["new_answer"]:
            dvals = [r["source_row"],r["request_id"],r["name"],r["session_name"],r["role"],r["new_answer"],r["responded_at"] or "",r["note"] or ""]
            for c,v in enumerate(dvals): diff.write(drow,c,v)
            drow += 1

    pheaders = ["氏名","ふりがな","メール","会員・非会員","携帯電話番号","氏名・所属等の修正依頼","招聘状","所属長の所属機関名","所属長の役職","所属長の氏名","指定様式・web申請","アップロードファイル","初回登録日時","登録元"]
    for c,h in enumerate(pheaders): pp.write(0,c,h,fmt_head)
    for rno,p in enumerate(all_people(),start=1):
        vals=[p["name"],p["furigana"],p["email"],p["membership"],p["mobile"],p["correction"],p["invitation"],p["leader_org"],p["leader_title"],p["leader_name"],p["special_request"],p["upload_name"],p["registered_at"],p["source"]]
        for c,v in enumerate(vals): pp.write(rno,c,v or "")

    ws.freeze_panes(1,0); diff.freeze_panes(1,0); pp.freeze_panes(1,0)
    ws.set_column(0,16,18); ws.set_column(6,6,45); ws.set_column(16,16,52)
    diff.set_column(0,7,20); pp.set_column(0,13,25)
    book.close(); out.seek(0)
    return out.getvalue()


def hero(text):
    st.markdown(f'<div class="hero"><h1>{APP_TITLE}</h1><p>{text}</p></div>', unsafe_allow_html=True)


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


def locked_profile(profile, fallback_name, fallback_email):
    st.markdown("### ご登録情報")
    st.caption("初回回答時に登録済みのため、今回は変更できません。変更が必要な場合は備考欄にご記入ください。")
    st.text_input("氏名", value=(profile["name"] or fallback_name), disabled=True)
    st.text_input("ふりがな", value=profile["furigana"] or "", disabled=True)
    st.text_input("メールアドレス", value=(profile["email"] or fallback_email), disabled=True)
    st.text_input("会員・非会員 （日本腎臓リハビリテーション学会）", value=profile["membership"] or "", disabled=True)
    st.text_input("当日の緊急連絡先（携帯電話番号）", value=profile["mobile"] or "", disabled=True)
    st.text_area("氏名・ご所属等の修正依頼", value=profile["correction"] or "", disabled=True)
    invitation_value = profile["invitation"] or "未登録"
    st.text_input("派遣依頼状（招聘状）", value=invitation_value, disabled=True)
    st.text_input("所属長の所属機関名", value=profile["leader_org"] or "", disabled=True)
    st.text_input("所属長の役職", value=profile["leader_title"] or "", disabled=True)
    st.text_input("所属長の氏名", value=profile["leader_name"] or "", disabled=True)
    st.text_area("指定様式での発行、web申請のご希望", value=profile["special_request"] or "", disabled=True)
    st.text_input("派遣依頼状 指定様式アップロード", value=profile["upload_name"] or "", disabled=True)


def initial_form(token, rows, pending):
    first = rows[0]
    st.markdown("### 初回ご回答")
    st.markdown('<div class="notice">初回のみ、ご連絡先・会員区分・修正依頼・緊急連絡先・招聘状についてもご登録ください。次回以降はこれらを表示のみとし、諾否と備考だけご回答いただけます。</div>', unsafe_allow_html=True)
    for r in pending:
        request_summary_card(r)

    name = st.text_input("氏名", value=first["name"] or "", key="first_name")
    furigana = st.text_input("ふりがな", key="first_furigana")
    email = st.text_input("メールアドレス", value=first["email"] or "", key="first_email")
    membership = st.radio(
        "会員・非会員 （日本腎臓リハビリテーション学会）",
        ["選択してください", "会員", "非会員", "入会申請中、入会予定"],
        horizontal=False, key="first_membership"
    )
    if membership == "非会員":
        st.info("非会員の先生へ：お引き受けいただける場合、謝金・旅費・宿泊費のご用意はございません。学会参加費は御招待（無料）です。")

    st.markdown("#### 諾否のご回答")
    answers = {}
    declines = {}
    for r in pending:
        label = f"{r['session_name']}｜{r['role']}"
        answers[r["request_id"]] = st.radio(
            label, ["選択してください", "承諾する", "辞退する"],
            horizontal=True, key=f"ans_{r['request_id']}"
        )
        declines[r["request_id"]] = st.text_area(
            f"辞退理由（{label}）※辞退の場合のみ", key=f"dec_{r['request_id']}"
        )

    correction = st.text_area(
        "氏名・ご所属等の修正依頼",
        placeholder="依頼状、メール本文内に記載の氏名・所属等に修正がある場合のみご入力ください。",
        key="first_correction"
    )
    mobile = st.text_input("当日の緊急連絡先（携帯電話番号）", placeholder="090-1234-5678", key="first_mobile")

    any_yes = any(v == "承諾する" for v in answers.values())
    invitation = ""
    leader_org = leader_title = leader_name = special_request = ""
    upload = None
    if any_yes:
        st.markdown("#### 派遣依頼状（招聘状）の発行について")
        st.caption("ご登録メールアドレス宛にPDFの添付ファイルでお送りします。")
        invitation = st.radio("派遣依頼状（招聘状）", ["選択してください", "必要", "不要"], horizontal=True, key="first_invitation")
        if invitation == "必要":
            st.markdown("##### 派遣依頼状（招聘状）の宛名について")
            leader_org = st.text_input("所属長の所属機関名", placeholder="例）〇〇大学〇〇学部、〇〇病院 等", key="first_leader_org")
            leader_title = st.text_input("所属長の役職", placeholder="例）学長、学部長、教授、病院長、理事長 等", key="first_leader_title")
            leader_name = st.text_input("所属長の氏名", key="first_leader_name")
            special_request = st.text_area(
                "指定様式での発行、web申請のご希望",
                placeholder="ご所属指定の様式がある場合やweb申請が必要な場合のみご入力ください。",
                key="first_special"
            )
            upload = st.file_uploader(
                "派遣依頼状 指定様式アップロード",
                type=["pdf","doc","docx","xls","xlsx","xlsm","jpg","jpeg","png"],
                key="first_upload"
            )

    note = st.text_area("備考", placeholder="ご連絡事項がございましたら、ご入力ください。", key="first_note")
    submitted = st.button("回答を登録する", type="primary", use_container_width=True, key="first_submit")

    if submitted:
        errors = []
        if not name.strip(): errors.append("氏名を入力してください。")
        if not furigana.strip(): errors.append("ふりがなを入力してください。")
        if not email.strip() or "@" not in email: errors.append("メールアドレスを入力してください。")
        if membership == "選択してください": errors.append("会員・非会員を選択してください。")
        if not mobile.strip(): errors.append("当日の緊急連絡先（携帯電話番号）を入力してください。")
        for r in pending:
            if answers[r["request_id"]] == "選択してください":
                errors.append(f"{r['session_name']}（{r['role']}）の諾否を選択してください。")
        if any_yes:
            if invitation == "選択してください": errors.append("招聘状の要否を選択してください。")
            if invitation == "必要" and not (leader_org.strip() and leader_title.strip() and leader_name.strip()):
                errors.append("招聘状が必要な場合は、所属長の所属機関名・役職・氏名を入力してください。")
        if errors:
            for e in errors: st.error(e)
            return

        upload_name, upload_path = save_upload(token, upload) if any_yes and invitation == "必要" else ("", "")
        upsert_profile({
            "token": token, "name": name.strip(), "email": email.strip(), "furigana": furigana.strip(),
            "membership": membership, "mobile": mobile.strip(), "correction": correction.strip(),
            "invitation": invitation if any_yes else "",
            "leader_org": leader_org.strip() if invitation == "必要" else "",
            "leader_title": leader_title.strip() if invitation == "必要" else "",
            "leader_name": leader_name.strip() if invitation == "必要" else "",
            "special_request": special_request.strip() if invitation == "必要" else "",
            "upload_name": upload_name if invitation == "必要" else "",
            "upload_path": upload_path if invitation == "必要" else "",
            "registered_at": datetime.now().isoformat(timespec="seconds"), "source": "system",
        }, overwrite=True)
        for r in pending:
            ans = "諾" if answers[r["request_id"]] == "承諾する" else "否"
            save_response(r["request_id"], ans, declines[r["request_id"]], note)
        st.success("回答を登録しました。次回以降は、諾否と備考のみご回答いただけます。")
        st.rerun()


def subsequent_form(profile, rows, pending):
    locked_profile(profile, rows[0]["name"], rows[0]["email"])
    if not pending:
        st.success("現在、新たにご回答いただく依頼はありません。")
        return

    st.markdown("### 今回ご回答いただくご依頼")
    st.markdown('<div class="notice">今回入力できるのは「諾否」と「備考」のみです。初回登録情報は上記に表示しています。</div>', unsafe_allow_html=True)
    for r in pending:
        request_summary_card(r)

    with st.form("subsequent_response"):
        answers = {}
        declines = {}
        for r in pending:
            label = f"{r['session_name']}｜{r['role']}"
            answers[r["request_id"]] = st.radio(label, ["選択してください", "承諾する", "辞退する"], horizontal=True, key=f"ans2_{r['request_id']}")
            declines[r["request_id"]] = st.text_area(f"辞退理由（{label}）※辞退の場合のみ", key=f"dec2_{r['request_id']}")
        note = st.text_area("備考", placeholder="ご連絡事項がございましたら、ご入力ください。")
        submitted = st.form_submit_button("今回の回答を登録する", type="primary", use_container_width=True)
        if submitted:
            missing = [r for r in pending if answers[r["request_id"]] == "選択してください"]
            if missing:
                st.error("すべてのご依頼について、承諾または辞退を選択してください。")
                return
            for r in pending:
                ans = "諾" if answers[r["request_id"]] == "承諾する" else "否"
                save_response(r["request_id"], ans, declines[r["request_id"]], note)
            st.success("回答を登録しました。")
            st.rerun()


def user_page(token):
    rows = get_person_requests(token)
    if not rows:
        hero("ご依頼内容の確認・諾否回答")
        st.error("このURLに該当するご依頼が見つかりません。")
        return

    hero("ご依頼内容の確認・諾否回答")
    st.markdown(f'<div class="person">{rows[0]["name"]} 先生</div><div class="subtle">{rows[0]["affiliation"]}</div>', unsafe_allow_html=True)
    pending = [r for r in rows if not effective_answer(r)]
    done = [r for r in rows if effective_answer(r)]
    profile = get_profile(token)

    if profile:
        subsequent_form(profile, rows, pending)
    else:
        if pending:
            initial_form(token, rows, pending)
        else:
            st.info("既に回答済みですが、初回登録情報が見つかりません。事務局へお問い合わせください。")

    if done:
        st.markdown("### これまでに回答済みのご依頼")
        for r in done:
            request_summary_card(r)


def admin_page():
    hero("管理画面｜指定演題Excel取込・回答状況確認")
    if not st.session_state.get("admin_ok"):
        pw = st.text_input("管理者パスワード", type="password")
        if st.button("ログイン", type="primary"):
            if hmac.compare_digest(pw, ADMIN_PASSWORD):
                st.session_state.admin_ok = True; st.rerun()
            else: st.error("パスワードが違います。")
        st.caption("試作版の初期パスワードは README に記載しています。本番公開前に必ず変更してください。")
        return

    tabs = st.tabs(["回答状況", "Excel取込", "木田先生テスト", "先生情報"])
    with tabs[0]:
        rows = all_rows(); data=[]
        for r in rows:
            eff=effective_answer(r)
            data.append({"元Excel行":r["source_row"],"氏名":r["name"],"セッション":r["session_name"],"役割":r["role"],"旧AD":r["source_answer"] or "","新回答":r["new_answer"] or "","現在":eff or "未回答","AF依頼送信":r["request_sent"],"回答日時":r["responded_at"] or "","マイページURL":f"{BASE_URL}/?token={r['token']}"})
        st.dataframe(data,use_container_width=True,hide_index=True)
        st.download_button("回答一覧・Excel反映用データをダウンロード",build_export_xlsx(rows),"JKRA17_諾否回答一覧.xlsx","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True)

    with tabs[1]:
        st.write("現在の『指定演題』Excelを取り込みます。AD列の既存諾否に加えて、旧Form Mailerの回答欄（送信日時～招聘状・備考等）も読み取り、初回登録済みとして引き継ぎます。")
        up=st.file_uploader("指定演題Excel（.xlsm / .xlsx）",type=["xlsm","xlsx"])
        if up and st.button("このExcelを取り込む",type="primary"):
            records,profiles=import_master(up.getvalue())
            st.success(f"{len(records)}件の依頼、旧フォーム登録済み先生 {len(profiles)}名分を取り込みました。")

    with tabs[2]:
        rows=[r for r in all_rows() if r["name"]=="木田圭亮"]
        if not rows: st.warning("木田圭亮先生のデータが見つかりません。")
        else:
            profile=get_profile(rows[0]["token"])
            st.write("初回登録判定：", "登録済み（2回目以降モード）" if profile else "未登録（初回モード）")
            if profile:
                st.write({"氏名":profile["name"],"ふりがな":profile["furigana"],"メール":profile["email"],"会員区分":profile["membership"],"携帯電話":profile["mobile"],"招聘状":profile["invitation"],"登録元":profile["source"]})
            for r in rows:
                st.write({"元Excel行":r["source_row"],"セッション":r["session_name"],"役割":r["role"],"AD":r["source_answer"] or "空欄","AF":r["request_sent"],"現在":effective_answer(r) or "未回答"})
            url=f"{BASE_URL}/?token={rows[0]['token']}"
            st.code(url,language=None); st.link_button("木田先生のマイページを開く",url,use_container_width=True)
            st.info("木田先生は旧Form Mailer回答済みのため、37行目の追加依頼では『諾否＋備考』のみ入力可能。初回登録情報は表示のみになります。")

    with tabs[3]:
        ps=[]
        for p in all_people():
            ps.append({"氏名":p["name"],"ふりがな":p["furigana"],"メール":p["email"],"会員区分":p["membership"],"携帯電話":p["mobile"],"招聘状":p["invitation"],"修正依頼":p["correction"],"登録元":p["source"],"登録日時":p["registered_at"]})
        st.dataframe(ps,use_container_width=True,hide_index=True)


bootstrap()
params=st.query_params
if params.get("admin")=="1": admin_page()
elif params.get("token"): user_page(params.get("token"))
else:
    hero("ご依頼内容の確認・諾否回答")
    st.info("メールに記載された専用URLからアクセスしてください。管理画面は ?admin=1 です。")
