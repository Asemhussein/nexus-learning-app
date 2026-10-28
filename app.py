import streamlit as st
import pandas as pd
from datetime import datetime
import os
import re
import json
import hmac
from io import BytesIO
import PyPDF2
from PIL import Image
from pptx import Presentation
from pptx.util import Pt
from google import genai
from google.genai import types

# --- Optional Imports ---
try:
    from supabase import create_client
    SUPABASE_INSTALLED = True
except ImportError:
    SUPABASE_INSTALLED = False

try:
    from streamlit_agraph import agraph, Node, Edge, Config
    AGRAPH_INSTALLED = True
except ImportError:
    AGRAPH_INSTALLED = False

# --- Page Config & CSS ---
st.set_page_config(page_title="Nexus | AI Second Brain", page_icon="🧠", layout="wide")

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
    :root {
        --nx-bg: #0b1020; --nx-border: rgba(148, 163, 184, .18);
        --nx-text: #e8eefb; --nx-muted: #9ba9c4;
    }
    html, body, [class*="css"] { font-family: 'Inter', sans-serif; }
    .stApp {
        background:
            radial-gradient(ellipse at 12% 0%, rgba(69, 96, 210, .18), transparent 34%),
            radial-gradient(ellipse at 92% 8%, rgba(27, 196, 191, .10), transparent 28%),
            #0b1020;
        color: var(--nx-text);
    }
    [data-testid="stHeader"] { background: rgba(11, 16, 32, .75); }
    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #10182c 0%, #0b1020 100%);
        border-right: 1px solid var(--nx-border);
    }
    [data-testid="stSidebar"] > div { padding-top: 1.25rem; }
    .block-container { padding-top: 1.5rem; padding-bottom: 3rem; max-width: 1600px; }
    .main-header {
        position: relative; overflow: hidden;
        background: linear-gradient(118deg, #18284e 0%, #263c78 48%, #087e91 100%);
        padding: 2rem 2.1rem; border: 1px solid rgba(255,255,255,.13);
        border-radius: 24px; color: white; margin-bottom: 1.35rem;
        box-shadow: 0 18px 55px rgba(0,0,0,.22);
    }
    .main-header:after {
        content: ""; position: absolute; width: 270px; height: 270px;
        right: -75px; top: -145px; border: 1px solid rgba(255,255,255,.16);
        border-radius: 50%;
        box-shadow: 0 0 0 30px rgba(255,255,255,.035), 0 0 0 65px rgba(255,255,255,.025);
        pointer-events: none;
    }
    .eyebrow { text-transform: uppercase; letter-spacing: .18em; font-size: .72rem; color: #b9f5ee; font-weight: 800; margin-bottom: .55rem; }
    .main-header h1 { color: #fff; margin: 0; font-size: clamp(1.7rem, 3vw, 2.55rem); font-weight: 800; letter-spacing: -.045em; }
    .main-header p { color: #dce8ff; margin: .65rem 0 0; font-size: .98rem; max-width: 760px; line-height: 1.65; }
    .status-badge {
        padding: 9px 11px; border-radius: 12px; font-size: .76rem; font-weight: 700;
        text-align: center; margin-bottom: 8px; border: 1px solid var(--nx-border);
        background: rgba(255,255,255,.04);
    }
    .badge-ok { color: #5eead4; border-color: rgba(45,212,191,.24); background: rgba(45,212,191,.08); }
    .badge-warn { color: #fbbf24; border-color: rgba(251,191,36,.24); background: rgba(251,191,36,.08); }
    div[data-testid="stMetric"] {
        background: linear-gradient(145deg, rgba(23,34,59,.95), rgba(16,24,43,.78));
        border: 1px solid var(--nx-border); padding: 1rem 1.1rem;
        border-radius: 18px; box-shadow: 0 8px 28px rgba(0,0,0,.10);
    }
    div[data-testid="stMetricLabel"] { color: #aab8d4; font-weight: 600; }
    div[data-testid="stMetricValue"] { color: #f4f7ff; font-weight: 800; }
    [data-testid="stVerticalBlockBorderWrapper"] > div {
        border-color: var(--nx-border) !important; border-radius: 18px !important;
        background: rgba(18,27,48,.42);
    }
    .stTabs [data-baseweb="tab-list"] {
        gap: 7px; border-bottom: 1px solid var(--nx-border);
        background: rgba(255,255,255,.025); padding: 7px; border-radius: 16px;
    }
    .stTabs [data-baseweb="tab"] {
        height: 43px; border-radius: 11px; padding: 0 15px;
        color: #aab8d4; font-weight: 650; border: 1px solid transparent;
    }
    .stTabs [aria-selected="true"] {
        background: linear-gradient(120deg, rgba(124,156,255,.22), rgba(77,226,208,.10));
        color: #f5f8ff !important; border-color: rgba(124,156,255,.22);
    }
    .stButton > button, .stDownloadButton > button {
        width: 100%; border-radius: 12px; min-height: 2.65rem; font-weight: 700;
        transition: transform .15s ease, box-shadow .15s ease, border-color .15s ease;
        border: 1px solid var(--nx-border);
    }
    .stButton > button:hover, .stDownloadButton > button:hover {
        transform: translateY(-1px); box-shadow: 0 8px 24px rgba(0,0,0,.18);
        border-color: rgba(124,156,255,.55);
    }
    .stButton > button[kind="primary"] { background: linear-gradient(115deg, #647ff5, #0eaaa8); color: white; border: 0; }
    input, textarea, [data-baseweb="select"] > div { border-radius: 11px !important; }
    .flashcard {
        background: linear-gradient(135deg, #1a2c51 0%, #111b31 100%);
        color: #f8fafc; padding: 1.2rem; border-radius: 16px;
        border: 1px solid rgba(124,156,255,.22); border-left: 4px solid #4de2d0;
        margin-bottom: .8rem; box-shadow: 0 8px 25px rgba(0,0,0,.12);
    }
    .empty-state {
        padding: 2rem; text-align: center; border-radius: 18px;
        border: 1px dashed rgba(148,163,184,.3); background: rgba(255,255,255,.025); color: #aab8d4;
    }
    hr { border-color: var(--nx-border) !important; }
    @media (max-width: 800px) {
        .block-container { padding: 1rem .8rem 2rem; }
        .main-header { padding: 1.4rem; border-radius: 18px; }
        .stTabs [data-baseweb="tab"] { padding: 0 9px; font-size: .78rem; }
    }
</style>
""", unsafe_allow_html=True)


# ================= SECRETS & CLIENTS =================
def get_secret(key):
    if key in os.environ and os.environ[key]:
        return os.environ[key].strip()
    try:
        if key in st.secrets:
            return str(st.secrets[key]).strip()
    except Exception:
        pass
    return None


def sanitize_supabase_url(raw_url):
    """Fix dashboard URLs / extra paths into a valid https://<ref>.supabase.co URL"""
    if not raw_url:
        return None
    raw_url = raw_url.strip().strip('"').strip("'")
    m = re.search(r'supabase\.com/dashboard/project/([a-z0-9]+)', raw_url)
    if m:
        return f"https://{m.group(1)}.supabase.co"
    m = re.search(r'(https://[a-z0-9]+\.supabase\.co)', raw_url)
    if m:
        return m.group(1)
    if re.match(r'^[a-z0-9]{15,25}$', raw_url):
        return f"https://{raw_url}.supabase.co"
    return raw_url.rstrip('/')


GEMINI_API_KEY = get_secret("GEMINI_API_KEY")
SUPABASE_URL = sanitize_supabase_url(get_secret("SUPABASE_URL"))
SUPABASE_KEY = get_secret("SUPABASE_KEY")
APP_PASSWORD = get_secret("APP_PASSWORD")  # optional: protects the app

# --- Optional password gate (recommended: the DB policy below is open to anyone with the key) ---
if APP_PASSWORD and not st.session_state.get("authed"):
    st.markdown("### 🔒 Nexus")
    pw = st.text_input("Password", type="password")
    if pw:
        if hmac.compare_digest(pw, APP_PASSWORD):
            st.session_state.authed = True
            st.rerun()
        else:
            st.error("Wrong password")
    st.stop()


@st.cache_resource(show_spinner=False)
def get_ai_client(key):
    return genai.Client(api_key=key) if key else None


@st.cache_resource(show_spinner=False)
def get_supabase_client(url, key):
    if SUPABASE_INSTALLED and url and key:
        return create_client(url, key)
    return None


ai_client = get_ai_client(GEMINI_API_KEY)
try:
    supabase_client = get_supabase_client(SUPABASE_URL, SUPABASE_KEY)
    _client_err = ""
except Exception as e:
    supabase_client = None
    _client_err = str(e)


# ================= DATABASE LAYER =================
def explain_db_error(e):
    """Turn raw Supabase/Postgres errors into an actionable message."""
    s = str(e)
    low = s.lower()
    if "row-level security" in low or "42501" in low:
        hint = "RLS شغال على الجدول ومفيش Policy بتسمح بالكتابة. شغّل SQL الـ policies من التعليمات."
    elif "42p10" in low or "no unique or exclusion constraint" in low:
        hint = "عمود title/term مش UNIQUE في الجدول، عشان كده الـ upsert بيفشل. شغّل SQL الإنشاء."
    elif "pgrst205" in low or "could not find the table" in low or "42p01" in low:
        hint = "الجدول مش موجود في Supabase. شغّل SQL الإنشاء في SQL Editor."
    elif "pgrst204" in low or ("column" in low and "does not exist" in low):
        hint = "ناقص عمود في الجدول (غالباً related_topics). شغّل ALTER TABLE من التعليمات."
    elif "invalid api key" in low or "jwt" in low or "apikey" in low:
        hint = "الـ SUPABASE_KEY غلط. استخدم anon/publishable key أو service_role من Project Settings > API."
    elif "name or service not known" in low or "connection" in low or "timed out" in low:
        hint = "مش قادر يوصل لـ Supabase. راجع SUPABASE_URL (أو المشروع Paused؟)."
    else:
        hint = "خطأ غير متوقع."
    return f"{hint}\n\nالتفاصيل: {s[:300]}"


def run_db_check():
    """Real read + write probe. Returns (ok, message)."""
    if not SUPABASE_INSTALLED:
        return False, "مكتبة supabase مش متسطبة (راجع requirements.txt)."
    if _client_err:
        return False, explain_db_error(_client_err)
    if not supabase_client:
        return False, "مفيش Secrets لـ SUPABASE_URL / SUPABASE_KEY → الداتا محلية بس وهتتمسح مع كل Refresh."
    today = datetime.now().strftime("%Y-%m-%d")
    try:
        supabase_client.table("topics").select("title").limit(1).execute()
        supabase_client.table("glossary").select("term").limit(1).execute()
        probe = {"title": "__nexus_probe__", "category": "x", "status": "x", "mastery": 0,
                 "notes": "", "ai_summary": "", "related_topics": "", "updated_at": today}
        supabase_client.table("topics").upsert(probe, on_conflict="title").execute()
        supabase_client.table("topics").delete().eq("title", "__nexus_probe__").execute()
        supabase_client.table("glossary").upsert(
            {"term": "__nexus_probe__", "category": "x", "context": "x"}, on_conflict="term").execute()
        supabase_client.table("glossary").delete().eq("term", "__nexus_probe__").execute()
        return True, ""
    except Exception as e:
        return False, explain_db_error(e)


def load_data_from_cloud():
    if not supabase_client:
        return False
    try:
        res_t = supabase_client.table("topics").select("*").order("updated_at", desc=True).execute()
        res_g = supabase_client.table("glossary").select("*").execute()
        st.session_state.topics = res_t.data or []
        st.session_state.glossary = res_g.data or []
        return True
    except Exception as e:
        st.session_state.db_error_msg = explain_db_error(e)
        return False


if 'data_initialized' not in st.session_state:
    st.session_state.topics = []
    st.session_state.glossary = []
    ok, msg = run_db_check()
    st.session_state.db_verified = ok
    st.session_state.db_error_msg = msg
    if ok:
        load_data_from_cloud()
    st.session_state.last_save_error = ""
    st.session_state.data_initialized = True


def _cloud_write(table, item, conflict_col):
    """Upsert to Supabase. Returns '' on success or an error message. Never hides failures."""
    if not supabase_client:
        return "Local mode: الداتا مش متسيفة في السحابة."
    try:
        supabase_client.table(table).upsert(item, on_conflict=conflict_col).execute()
        st.session_state.db_verified = True
        return ""
    except Exception as e:
        st.session_state.db_verified = False
        return explain_db_error(e)


def upsert_topic(title, category, status, mastery, notes=None, ai_summary=None, related_topics=None):
    clean_title = title.strip()
    today = datetime.now().strftime("%Y-%m-%d")
    idx = next((i for i, t in enumerate(st.session_state.topics)
                if t['title'].lower() == clean_title.lower()), None)
    curr = st.session_state.topics[idx] if idx is not None else {}

    item = {
        "title": curr.get("title", clean_title),
        "category": category or curr.get("category") or "General",
        "status": status or curr.get("status") or "In Progress",
        "mastery": int(mastery) if mastery is not None else int(curr.get("mastery") or 50),
        "notes": notes if notes is not None else (curr.get("notes") or ""),
        "ai_summary": ai_summary if ai_summary is not None else (curr.get("ai_summary") or ""),
        "related_topics": related_topics if related_topics is not None else (curr.get("related_topics") or ""),
        "updated_at": today,
    }
    if idx is not None:
        st.session_state.topics[idx] = item
        action = "updated"
    else:
        st.session_state.topics.insert(0, item)
        action = "created"

    err = _cloud_write("topics", item, "title")
    st.session_state.last_save_error = err
    return action, err


def delete_topic(title):
    st.session_state.topics = [t for t in st.session_state.topics if t['title'] != title]
    if supabase_client:
        try:
            supabase_client.table("topics").delete().eq("title", title).execute()
            st.session_state.last_save_error = ""
        except Exception as e:
            st.session_state.last_save_error = explain_db_error(e)


def upsert_glossary(term, category, context):
    clean_term = term.strip()
    item = {"term": clean_term, "category": category, "context": context}
    idx = next((i for i, g in enumerate(st.session_state.glossary)
                if g['term'].lower() == clean_term.lower()), None)
    if idx is not None:
        st.session_state.glossary[idx] = item
    else:
        st.session_state.glossary.insert(0, item)
    err = _cloud_write("glossary", item, "term")
    if err:
        st.session_state.last_save_error = err
    return err


# ================= AI HELPERS =================
def parse_json_loose(text):
    """Extract JSON from a model reply even if wrapped in ``` fences or extra prose."""
    t = text.strip()
    t = re.sub(r"^```(?:json)?", "", t).strip()
    t = re.sub(r"```$", "", t).strip()
    try:
        return json.loads(t)
    except Exception:
        pass
    starts = [p for p in (t.find("["), t.find("{")) if p != -1]
    if not starts:
        raise ValueError("No JSON found")
    start = min(starts)
    end = max(t.rfind("]"), t.rfind("}"))
    return json.loads(t[start:end + 1])


def ask_gemini(prompt, image_file=None, selected_model="gemini-2.5-flash", json_mode=False):
    if not ai_client:
        return False, "⚠️ AI Engine offline. Please verify GEMINI_API_KEY."
    contents = [prompt]
    if image_file:
        try:
            contents.append(Image.open(image_file))
        except Exception as e:
            return False, f"⚠️ Couldn't read the image: {e}"

    config = types.GenerateContentConfig(response_mime_type="application/json") if json_mode else None
    chain = []
    for m in [selected_model, "gemini-2.5-flash", "gemini-2.0-flash"]:
        if m and m not in chain:
            chain.append(m)

    last_err = "Empty response from model."
    for m in chain:
        try:
            response = ai_client.models.generate_content(model=m, contents=contents, config=config)
            if response and response.text:
                return True, response.text
        except Exception as e:
            last_err = f"{m}: {e}"
    return False, f"⚠️ AI Request Failed: {last_err[:220]}"


def build_advanced_pptx(topic, audience, slides_json):
    prs = Presentation()
    s0 = prs.slides.add_slide(prs.slide_layouts[0])
    s0.shapes.title.text = topic
    s0.placeholders[1].text = f"Target Audience: {audience}\nGenerated by Nexus AI Engine"
    for s in slides_json:
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        slide.shapes.title.text = s.get("title", "Slide")
        tf = slide.placeholders[1].text_frame
        tf.word_wrap = True
        for idx, b in enumerate(s.get("bullets", [])):
            p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
            p.text = str(b)
            p.font.size = Pt(20)
        slide.notes_slide.notes_text_frame.text = (
            f"⏱️ Time: {s.get('time', '2 mins')}\n\n🎙️ Script:\n{s.get('speaker_notes', '')}")
    stream = BytesIO()
    prs.save(stream)
    stream.seek(0)
    return stream


# ================= SIDEBAR =================
topic_names = [t['title'] for t in st.session_state.topics]

with st.sidebar:
    st.markdown("### ⚡ Nexus Control Center")
    ai_badge = ('<div class="status-badge badge-ok">🤖 AI: Ready</div>' if ai_client
                else '<div class="status-badge badge-warn">🤖 AI: Missing Key</div>')
    db_badge = ('<div class="status-badge badge-ok">☁️ Cloud DB: Read/Write OK</div>' if st.session_state.get("db_verified")
                else '<div class="status-badge badge-warn">💾 DB: NOT saving</div>')
    c1, c2 = st.columns(2)
    c1.markdown(ai_badge, unsafe_allow_html=True)
    c2.markdown(db_badge, unsafe_allow_html=True)

    preset = st.selectbox("⚡ AI Model:", ["gemini-2.5-flash", "gemini-2.5-pro", "gemini-2.0-flash"], index=0)
    custom_model = st.text_input("Custom model (optional):", placeholder="e.g. gemini-3-flash")
    model_choice = custom_model.strip() or preset
    st.divider()

    st.markdown("#### 🔄 Add / Update Topic")
    mode = st.radio("Action:", ["➕ Create New", "✏️ Edit Existing"], horizontal=True, label_visibility="collapsed")

    cat_options = ["Project Management (PM)", "Technical / Engineering", "AI & Automation", "Business & Strategy"]
    stat_options = ["Not Started", "In Progress", "Completed"]

    if mode == "✏️ Edit Existing" and topic_names:
        sel_edit = st.selectbox("Select Topic:", topic_names)
        curr_t = next(t for t in st.session_state.topics if t['title'] == sel_edit)
        in_title = curr_t['title']
        default_cat = curr_t.get('category') or cat_options[0]
        default_stat = curr_t.get('status') or "In Progress"
        default_mast = int(curr_t.get('mastery') or 50)
        existing_rels = [r.strip() for r in str(curr_t.get('related_topics') or '').split(',')
                         if r.strip() in topic_names]
        wkey = f"edit_{sel_edit}"
    else:
        in_title = st.text_input("Topic Title:", placeholder="e.g., 5G Site Rollout KPIs")
        default_cat, default_stat, default_mast, existing_rels = cat_options[0], "In Progress", 50, []
        wkey = "new"

    in_cat = st.selectbox("Category:", cat_options,
                          index=cat_options.index(default_cat) if default_cat in cat_options else 0, key=f"cat_{wkey}")
    in_stat = st.selectbox("Status:", stat_options,
                           index=stat_options.index(default_stat) if default_stat in stat_options else 1, key=f"stat_{wkey}")
    in_mast = st.slider("Mastery (%):", 0, 100, default_mast, step=5, key=f"mast_{wkey}")

    available_links = [t for t in topic_names if t.lower() != in_title.strip().lower()]
    linked_topics = st.multiselect("🔗 Link to Topics (Graph):", available_links,
                                   default=existing_rels, key=f"links_{wkey}")

    if st.button("💾 Save / Update", type="primary"):
        if in_title.strip():
            act, err = upsert_topic(in_title, in_cat, in_stat, in_mast, related_topics=", ".join(linked_topics))
            if not err:
                st.toast(f"☁️ Topic {act} & saved to cloud", icon="✅")
            st.rerun()
        else:
            st.warning("Enter a topic title.")

    if mode == "✏️ Edit Existing" and topic_names:
        if st.button("🗑️ Delete this topic"):
            delete_topic(in_title)
            st.rerun()

# ================= MAIN HEADER =================
st.markdown("""
<div class="main-header">
    <div class="eyebrow">PERSONAL KNOWLEDGE SYSTEM · AI-POWERED WORKSPACE</div>
    <h1>🧠 Nexus <span style="font-weight:500; opacity:.82;">/ Your Second Brain</span></h1>
    <p>Capture what you learn, connect ideas, measure mastery, and turn knowledge into clear presentations — all in one focused workspace.</p>
</div>
""", unsafe_allow_html=True)

# Honest persistence banners (the old version hid these errors)
if st.session_state.get("last_save_error"):
    st.error("⚠️ آخر عملية حفظ **ماوصلتش** للسحابة:\n\n" + st.session_state.last_save_error)
elif not st.session_state.get("db_verified"):
    st.warning("⚠️ الداتا حالياً **محلية بس** وهتتمسح مع أي Refresh. افتح تاب Cloud Diagnostics لتشوف السبب.")

topic_list = [t['title'] for t in st.session_state.topics]

tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
    "📊 1. Tracker", "🕸️ 2. Visual Map", "🧠 3. Knowledge Vault",
    "🎯 4. Pitch Studio", "🎙️ 5. AI Coach & Cards", "⚙️ 6. Cloud Diagnostics"
])

# ================= TAB 1: TRACKER =================
with tab1:
    if st.session_state.topics:
        df = pd.DataFrame(st.session_state.topics)
        for col in ["category", "status", "notes", "related_topics", "updated_at"]:
            if col not in df.columns:
                df[col] = ""
        df['mastery'] = pd.to_numeric(df.get('mastery'), errors='coerce').fillna(0).astype(int)

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("📚 Total Topics", len(df))
        m2.metric("✅ Completed", len(df[df['status'] == 'Completed']))
        m3.metric("⚡ In Progress", len(df[df['status'] == 'In Progress']))
        m4.metric("🎯 Avg Mastery", f"{df['mastery'].mean():.0f}%")

        done_rate = len(df[df['status'] == 'Completed']) / len(df) * 100
        st.caption(f"{done_rate:.0f}% of topics marked complete")
        st.divider()

        ch1, ch2 = st.columns(2)
        with ch1:
            st.markdown("##### 📈 Mastery by Topic")
            st.bar_chart(df.set_index("title")["mastery"], color="#0284c7")
        with ch2:
            st.markdown("##### 📂 Topics per Domain")
            st.bar_chart(df['category'].value_counts(), color="#1e3a8a")

        st.markdown("### 📌 Interactive Knowledge Directory")
        f1, f2 = st.columns([2, 1])
        with f1:
            search_q = st.text_input("🔍 Search topics or notes...", placeholder="Type keyword...")
        with f2:
            cats = list(df['category'].unique())
            filter_cat = st.multiselect("Filter by Category:", cats, default=cats)

        fdf = df[df['category'].isin(filter_cat)]
        if search_q:
            fdf = fdf[fdf['title'].str.contains(search_q, case=False, na=False, regex=False) |
                      fdf['notes'].astype(str).str.contains(search_q, case=False, na=False, regex=False)]

        cols_to_show = ['title', 'category', 'status', 'mastery', 'related_topics', 'updated_at', 'notes']
        st.dataframe(
            fdf[cols_to_show],
            column_config={
                "title": "Topic Title", "category": "Domain", "status": "Status",
                "mastery": st.column_config.ProgressColumn("Mastery", format="%d%%", min_value=0, max_value=100),
                "related_topics": "🔗 Linked Topics", "updated_at": "Updated", "notes": "Key Notes",
            },
            hide_index=True,
        )
        st.download_button("⬇️ Export filtered topics (CSV)",
                           data=fdf[cols_to_show].to_csv(index=False).encode("utf-8-sig"),
                           file_name="nexus_topics.csv", mime="text/csv")
    else:
        st.markdown('<div class="empty-state"><h3>Your knowledge base starts here</h3>'
                    '<p>Add a topic from the sidebar to start tracking progress, connecting ideas, and building your second brain.</p></div>',
                    unsafe_allow_html=True)

# ================= TAB 2: VISUAL MAP =================
with tab2:
    st.subheader("🕸️ Interactive Second-Brain Knowledge Graph")
    g1, g2 = st.columns([3, 1])
    with g2:
        show_categories = st.checkbox("Show Category Hubs", value=True)
        physics_enabled = st.checkbox("Enable Bouncy Physics", value=True)

        if st.button("🤖 AI Auto-Connect Topics", type="primary"):
            if len(st.session_state.topics) >= 2:
                with st.spinner("Discovering topic connections..."):
                    t_list_str = "\n".join(f"- {t['title']} ({t.get('category', '')})" for t in st.session_state.topics)
                    link_prompt = f"""Connect related topics from this list:
{t_list_str}
Return ONLY valid JSON mapping exact topic titles to a comma-separated string of related exact topic titles:
{{"Topic A": "Topic B, Topic C"}}"""
                    ok, res = ask_gemini(link_prompt, selected_model=model_choice, json_mode=True)
                    if ok:
                        try:
                            links_map = parse_json_loose(res)
                            for t_title, rels in links_map.items():
                                m = next((t for t in st.session_state.topics if t['title'].lower() == str(t_title).lower()), None)
                                if m:
                                    upsert_topic(m['title'], m.get('category'), m.get('status'), m.get('mastery'),
                                                 related_topics=rels if isinstance(rels, str) else ", ".join(rels))
                            st.rerun()
                        except Exception as e:
                            st.warning(f"Parse error: {e}")
                    else:
                        st.error(res)
            else:
                st.info("Add at least 2 topics first.")
        st.markdown("**🎨 Legend:**\n- 🟦 Domain Hub\n- 🟢 Completed (≥80%)\n- 🟠 In Progress\n- 🔴 Early Stage")

    with g1:
        if AGRAPH_INSTALLED and st.session_state.topics:
            nodes, edges, added, seen_edges = [], [], set(), set()
            if show_categories:
                for c in set(t.get('category') or 'General' for t in st.session_state.topics):
                    hub_id = f"HUB_{c}"
                    nodes.append(Node(id=hub_id, label=f"📂 {c}", size=26, color="#1e3a8a",
                                      shape="hexagon", font={"color": "#ffffff"}))
                    added.add(hub_id)
            for t in st.session_state.topics:
                t_id = t['title']
                mast = int(t.get('mastery') or 50)
                color = "#10b981" if (t.get('status') == "Completed" or mast >= 80) else ("#f59e0b" if mast >= 40 else "#ef4444")
                if t_id not in added:
                    nodes.append(Node(id=t_id, label=f"{t_id}\n({mast}%)", size=14 + int(mast / 6),
                                      color=color, shape="dot", font={"color": "#e8eefb"}))
                    added.add(t_id)
                if show_categories:
                    edges.append(Edge(source=f"HUB_{t.get('category') or 'General'}", target=t_id, color="#94a3b8"))
            for t in st.session_state.topics:
                for r in [x.strip() for x in str(t.get('related_topics') or '').split(',') if x.strip()]:
                    tgt = next((x['title'] for x in st.session_state.topics if x['title'].lower() == r.lower()), None)
                    if tgt and tgt != t['title']:
                        key = frozenset((t['title'], tgt))
                        if key not in seen_edges:
                            seen_edges.add(key)
                            edges.append(Edge(source=t['title'], target=tgt, color="#38bdf8", width=2))
            agraph(nodes=nodes, edges=edges, config=Config(width=900, height=540, directed=False, physics=physics_enabled))
        elif not st.session_state.topics:
            st.markdown('<div class="empty-state"><h3>No nodes to visualize yet</h3>'
                        '<p>Create a few topics first, then use AI Auto-Connect.</p></div>', unsafe_allow_html=True)
        else:
            st.info("Install `streamlit-agraph` to enable the interactive knowledge graph.")

# ================= TAB 3: KNOWLEDGE VAULT =================
with tab3:
    st.subheader("📥 Fast Multi-Modal Knowledge Ingestion")
    if topic_list:
        active_name = st.selectbox("Select Topic:", topic_list)
        active = next(t for t in st.session_state.topics if t['title'] == active_name)

        v1, v2 = st.columns(2)
        with v1:
            user_notes = st.text_area("✍️ Personal Notes & Key Formulas:", value=active.get('notes') or '',
                                      height=180, key=f"notes_{active_name}")
            if st.button("💾 Save Notes Only"):
                _, err = upsert_topic(active_name, active.get('category'), active.get('status'),
                                      active.get('mastery'), notes=user_notes)
                if err:
                    st.error(err)
                else:
                    st.success("Notes saved to cloud ✅")

        with v2:
            up_file = st.file_uploader("📄 Upload PDF or Slide Image (PNG/JPG):", type=['pdf', 'png', 'jpg', 'jpeg'])
            extract_glossary = st.checkbox("Auto-extract terms to Glossary", value=True)

            if st.button("⚡ Instant AI Summarize", type="primary"):
                if not up_file and not user_notes:
                    st.warning("Upload a file or write notes first.")
                else:
                    with st.spinner(f"Analyzing with {model_choice}..."):
                        extracted_text, img_to_send = "", None
                        if up_file:
                            if up_file.name.lower().endswith('.pdf'):
                                try:
                                    reader = PyPDF2.PdfReader(up_file)
                                    for page in reader.pages[:8]:
                                        extracted_text += (page.extract_text() or "") + "\n"
                                    if not extracted_text.strip():
                                        st.warning("الـ PDF ده شكله صور (Scanned) ومفيش نص يتقرا. جرب ترفعه كصورة.")
                                except Exception as e:
                                    st.error(f"PDF read error: {e}")
                            else:
                                img_to_send = up_file

                        prompt = f"""Summarize this material for '{active_name}'.
Notes: {user_notes}
Content: {extracted_text[:6000]}

Format:
### 📌 Executive Summary (4 bullets)
### ⚙️ Key Engineering/PM Takeaways
### 📖 Key Terms (Strictly 3 lines: TERM: [name] | DEF: [short definition])"""
                        ok, ai_res = ask_gemini(prompt, image_file=img_to_send, selected_model=model_choice)
                        if ok:
                            _, err = upsert_topic(active_name, active.get('category'), active.get('status'),
                                                  active.get('mastery'), notes=user_notes, ai_summary=ai_res)
                            if extract_glossary:
                                for line in ai_res.split('\n'):
                                    if "TERM:" in line and "DEF:" in line:
                                        term_part, def_part = line.split("DEF:", 1)
                                        term = term_part.replace("TERM:", "").replace("*", "").replace("|", "").strip(" -•:[]")
                                        if term:
                                            upsert_glossary(term, "Technical", def_part.strip(" *[]"))
                            st.rerun()
                        else:
                            st.error(ai_res)

        saved_sum = active.get('ai_summary', '')
        if saved_sum:
            st.divider()
            st.markdown("#### 🤖 Saved AI Synthesis")
            with st.container(border=True):
                st.markdown(saved_sum)
    else:
        st.info("Add a topic from the sidebar first.")

# ================= TAB 4: PITCH STUDIO =================
with tab4:
    st.subheader("🎯 Automated Executive Presentation Studio")
    p1, p2, p3 = st.columns(3)
    with p1:
        pitch_topic = st.selectbox("Select Topic:", topic_list if topic_list else ["General"], key="pitch_sel")
    with p2:
        audience = st.selectbox("Target Audience:", ["Executive Board / PMO", "Engineering Deep-Dive", "Oral Assessment Panel"])
    with p3:
        slide_count = st.slider("Slides:", 3, 6, 4)

    if st.button("✨ Generate Slide Deck + Script", type="primary"):
        t_obj = next((t for t in st.session_state.topics if t['title'] == pitch_topic), {})
        with st.spinner("Generating presentation..."):
            pitch_prompt = f"""Create a {slide_count}-slide presentation on '{pitch_topic}' for '{audience}'.
Notes: {(t_obj.get('notes') or '')[:1500]}
Summary: {(t_obj.get('ai_summary') or '')[:1500]}
Return ONLY a JSON array:
[{{"title": "Slide Title", "bullets": ["Point 1", "Point 2"], "speaker_notes": "Script...", "time": "2 mins"}}]"""
            ok, raw = ask_gemini(pitch_prompt, selected_model=model_choice, json_mode=True)
            if ok:
                try:
                    slides_data = parse_json_loose(raw)
                    st.session_state.last_deck = (pitch_topic, audience, slides_data)
                except Exception:
                    st.write(raw)
            else:
                st.error(raw)

    if st.session_state.get("last_deck"):
        d_topic, d_aud, slides_data = st.session_state.last_deck
        cols = st.columns(2)
        for i, s in enumerate(slides_data):
            with cols[i % 2]:
                with st.container(border=True):
                    st.markdown(f"#### 📊 Slide {i+1}: {s.get('title')} `{s.get('time')}`")
                    for b in s.get("bullets", []):
                        st.markdown(f"- {b}")
                    with st.expander("🎙️ Speaker Script"):
                        st.write(s.get("speaker_notes"))
        st.download_button("📥 Download PowerPoint (.pptx)",
                           data=build_advanced_pptx(d_topic, d_aud, slides_data),
                           file_name=f"{d_topic}_Deck.pptx",
                           mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                           type="primary")

# ================= TAB 5: AI COACH & FLASHCARDS =================
with tab5:
    c_left, c_right = st.columns([3, 2])
    with c_left:
        st.subheader("🎙️ Active Recall & Oral Coach")
        coach_mode = st.radio("Mode:", ["🧠 Feynman Evaluator", "🃏 Smart Flashcards", "📝 Quiz"], horizontal=True)
        coach_topic = st.selectbox("Topic:", topic_list if topic_list else ["General"], key="coach_sel")

        if coach_mode == "🧠 Feynman Evaluator":
            user_exp = st.text_area("Explain the concept in English:", height=120)
            if st.button("🔍 Evaluate Delivery", type="primary") and user_exp:
                with st.spinner("Evaluating..."):
                    ok, res = ask_gemini(
                        f"Evaluate this explanation of '{coach_topic}': '{user_exp}'. "
                        "Give a score /10, missing terms, and a polished 3-sentence version.",
                        selected_model=model_choice)
                    st.markdown(res) if ok else st.error(res)
        elif coach_mode == "🃏 Smart Flashcards":
            if st.button("⚡ Generate 4 Flashcards", type="primary"):
                with st.spinner("Creating cards..."):
                    ok, res = ask_gemini(
                        f"Return ONLY a JSON array of 4 flashcards for '{coach_topic}': "
                        '[{"q":"Question?","a":"Answer"}]',
                        selected_model=model_choice, json_mode=True)
                    if ok:
                        try:
                            st.session_state.flashcards = parse_json_loose(res)
                        except Exception:
                            st.write(res)
                    else:
                        st.error(res)
            for idx, fc in enumerate(st.session_state.get("flashcards", [])):
                st.markdown(f'<div class="flashcard"><b>Q{idx+1}: {fc.get("q")}</b></div>', unsafe_allow_html=True)
                with st.expander("👀 Reveal Answer"):
                    st.success(fc.get('a'))
        else:
            if st.button("🎲 Generate 3 Questions", type="primary"):
                with st.spinner("Generating quiz..."):
                    ok, res = ask_gemini(
                        f"Create 3 multiple-choice questions on '{coach_topic}' with answers at the end.",
                        selected_model=model_choice)
                    st.markdown(res) if ok else st.error(res)

    with c_right:
        st.subheader("📖 Smart Glossary")
        with st.expander("➕ Add Term"):
            g_t = st.text_input("Term:")
            g_c = st.selectbox("Domain:", ["PM", "Technical", "Business"])
            g_d = st.text_area("Definition:")
            if st.button("Save Term") and g_t.strip():
                err = upsert_glossary(g_t, g_c, g_d)
                if not err:
                    st.toast("Term saved ✅")
                st.rerun()
        if st.session_state.glossary:
            gdf = pd.DataFrame(st.session_state.glossary)
            show = [c for c in ['term', 'category', 'context'] if c in gdf.columns]
            st.dataframe(gdf[show], hide_index=True)

# ================= TAB 6: CLOUD DIAGNOSTICS =================
with tab6:
    st.subheader("⚙️ Cloud Database Diagnostics & Backup")
    st.write(f"**Supabase API URL:** `{SUPABASE_URL or 'Not Set'}`")
    st.write(f"**Supabase key set:** `{'yes' if SUPABASE_KEY else 'no'}`  |  "
             f"**Topics in memory:** `{len(st.session_state.topics)}`")

    if st.session_state.get("db_verified"):
        st.success("✅ قراءة وكتابة Supabase شغالين تمام.")
    else:
        st.error("❌ الحفظ في السحابة مش شغال. السبب:")
        st.code(st.session_state.get("db_error_msg") or "Unknown")

    if st.button("🔄 Run full read/write test & reload from cloud", type="primary"):
        ok, msg = run_db_check()
        st.session_state.db_verified = ok
        st.session_state.db_error_msg = msg
        if ok:
            load_data_from_cloud()
            st.session_state.last_save_error = ""
        st.rerun()

    with st.expander("🧱 SQL setup (run once in Supabase → SQL Editor)"):
        st.code("""
create table if not exists topics (
  id bigint generated always as identity primary key,
  title text unique not null,
  category text, status text,
  mastery int default 50,
  notes text default '', ai_summary text default '', related_topics text default '',
  updated_at text
);
create table if not exists glossary (
  id bigint generated always as identity primary key,
  term text unique not null,
  category text, context text
);

-- if the tables already existed, make sure of these:
alter table topics add column if not exists related_topics text default '';
alter table topics add column if not exists ai_summary text default '';

alter table topics enable row level security;
alter table glossary enable row level security;
drop policy if exists "nexus_all_topics" on topics;
drop policy if exists "nexus_all_glossary" on glossary;
create policy "nexus_all_topics" on topics for all to anon, authenticated using (true) with check (true);
create policy "nexus_all_glossary" on glossary for all to anon, authenticated using (true) with check (true);
""", language="sql")

    st.divider()
    st.markdown("#### 📦 Manual JSON Backup & Restore")
    backup_data = json.dumps({"topics": st.session_state.topics, "glossary": st.session_state.glossary},
                             indent=2, default=str)
    st.download_button("⬇️ Download Full Backup (.json)", data=backup_data,
                       file_name="nexus_backup.json", mime="application/json")

    restore = st.file_uploader("Restore from backup (.json)", type=["json"], key="restore_up")
    if restore and st.button("♻️ Restore into app (and cloud)"):
        try:
            data = json.load(restore)
            for t in data.get("topics", []):
                upsert_topic(t.get("title", "Untitled"), t.get("category"), t.get("status"), t.get("mastery"),
                             notes=t.get("notes"), ai_summary=t.get("ai_summary"),
                             related_topics=t.get("related_topics"))
            for g in data.get("glossary", []):
                upsert_glossary(g.get("term", ""), g.get("category"), g.get("context"))
            st.success("Restored ✅")
            st.rerun()
        except Exception as e:
            st.error(f"Restore failed: {e}")
