import streamlit as st
import pandas as pd
from datetime import datetime
import os
import re
import json
from io import BytesIO
import PyPDF2
from PIL import Image
from pptx import Presentation
from pptx.util import Pt
from google import genai

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

# --- Page Config & Refined Modern CSS ---
st.set_page_config(page_title="Nexus | AI Second Brain", page_icon="🧠", layout="wide")

st.markdown("""
<style>
    .main-header {
        background: linear-gradient(135deg, #0f172a 0%, #1e3a8a 60%, #0284c7 100%);
        padding: 1.4rem 2rem;
        border-radius: 14px;
        color: white;
        margin-bottom: 1.2rem;
        box-shadow: 0 4px 15px rgba(0,0,0,0.15);
    }
    .main-header h1 { color: #ffffff; margin: 0; font-size: 1.9rem; font-weight: 800; }
    .main-header p { color: #e2e8f0; margin: 0.3rem 0 0 0; font-size: 0.95rem; }
    
    .status-badge {
        padding: 8px 12px;
        border-radius: 8px;
        font-size: 0.82rem;
        font-weight: 600;
        text-align: center;
        margin-bottom: 8px;
    }
    .badge-ok { background-color: rgba(16, 185, 129, 0.15); color: #10b981; border: 1px solid rgba(16, 185, 129, 0.3); }
    .badge-warn { background-color: rgba(245, 158, 11, 0.15); color: #f59e0b; border: 1px solid rgba(245, 158, 11, 0.3); }
    
    div[data-testid="stMetric"] {
        background-color: rgba(30, 41, 59, 0.04);
        border: 1px solid rgba(148, 163, 184, 0.25);
        padding: 0.8rem 1rem;
        border-radius: 12px;
    }
    .flashcard {
        background: linear-gradient(135deg, #1e293b 0%, #0f172a 100%);
        color: #f8fafc;
        padding: 1.2rem;
        border-radius: 12px;
        border-left: 5px solid #38bdf8;
        margin-bottom: 0.8rem;
    }
</style>
""", unsafe_allow_html=True)

# --- Smart Secret & URL Sanitizer ---
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
    """Automatically fixes dashboard URLs or extra paths into valid *.supabase.co API URLs"""
    if not raw_url:
        return None
    raw_url = raw_url.strip().strip('"').strip("'")
    # If user pasted: https://supabase.com/dashboard/project/abcdefghijk/...
    match_dash = re.search(r'supabase\.com/dashboard/project/([a-z0-9]+)', raw_url)
    if match_dash:
        return f"https://{match_dash.group(1)}.supabase.co"
    # If user pasted: https://abcdefghijk.supabase.co/rest/v1/...
    match_api = re.search(r'(https://[a-z0-9]+\.supabase\.co)', raw_url)
    if match_api:
        return match_api.group(1)
    # If user pasted just the project ID: abcdefghijk
    if re.match(r'^[a-z0-9]{15,25}\$', raw_url):
        return f"https://{raw_url}.supabase.co"
    return raw_url.rstrip('/')

GEMINI_API_KEY = get_secret("GEMINI_API_KEY")
RAW_SUPABASE_URL = get_secret("SUPABASE_URL")
SUPABASE_URL = sanitize_supabase_url(RAW_SUPABASE_URL)
SUPABASE_KEY = get_secret("SUPABASE_KEY")

ai_client = genai.Client(api_key=GEMINI_API_KEY) if GEMINI_API_KEY else None

supabase_client = None
if SUPABASE_INSTALLED and SUPABASE_URL and SUPABASE_KEY:
    try:
        supabase_client = create_client(SUPABASE_URL, SUPABASE_KEY)
    except Exception as e:
        st.session_state.db_error_msg = str(e)

# --- Cloud Sync & Local Persistence ---
def load_data_from_cloud():
    if supabase_client:
        try:
            res_t = supabase_client.table("topics").select("*").order("id", desc=True).execute()
            res_g = supabase_client.table("glossary").select("*").order("id", desc=True).execute()
            st.session_state.topics = res_t.data if res_t.data else []
            st.session_state.glossary = res_g.data if res_g.data else []
            st.session_state.db_verified = True
            st.session_state.db_error_msg = ""
            return True
        except Exception as e:
            st.session_state.db_verified = False
            st.session_state.db_error_msg = str(e)
            return False
    return False

if 'data_initialized' not in st.session_state:
    if not load_data_from_cloud():
        st.session_state.topics = [
            {"title": "Telecom Rollout KPIs & Risk Mgmt", "category": "Project Management (PM)", "status": "Completed", "mastery": 90, "notes": "Critical path & vendor SLA tracking.", "ai_summary": "", "related_topics": "NVDLA Architecture & LUT Optimization", "updated_at": "2026-10-28"},
            {"title": "NVDLA Architecture & LUT Optimization", "category": "Technical / Engineering", "status": "Completed", "mastery": 95, "notes": "Hardware accelerator lookup-table logic.", "ai_summary": "", "related_topics": "Telecom Rollout KPIs & Risk Mgmt", "updated_at": "2026-10-28"}
        ]
        st.session_state.glossary = [
            {"term": "Critical Path", "category": "PM", "context": "The longest sequence of tasks that determines the project completion date."}
        ]
    st.session_state.data_initialized = True

def upsert_topic(title, category, status, mastery, notes=None, ai_summary=None, related_topics=None):
    clean_title = title.strip()
    today = datetime.now().strftime("%Y-%m-%d")
    
    existing_idx = next((i for i, t in enumerate(st.session_state.topics) if t['title'].lower() == clean_title.lower()), None)
    
    if existing_idx is not None:
        curr = st.session_state.topics[existing_idx]
        clean_title = curr['title']
        updated_item = {
            "title": clean_title,
            "category": category or curr.get("category", "General"),
            "status": status or curr.get("status", "In Progress"),
            "mastery": int(mastery) if mastery is not None else int(curr.get("mastery", 50)),
            "notes": notes if notes is not None else curr.get("notes", ""),
            "ai_summary": ai_summary if ai_summary is not None else curr.get("ai_summary", ""),
            "related_topics": related_topics if related_topics is not None else curr.get("related_topics", ""),
            "updated_at": today
        }
        st.session_state.topics[existing_idx] = updated_item
        action = "updated"
    else:
        updated_item = {
            "title": clean_title,
            "category": category or "General",
            "status": status or "In Progress",
            "mastery": int(mastery) if mastery is not None else 50,
            "notes": notes or "",
            "ai_summary": ai_summary or "",
            "related_topics": related_topics or "",
            "updated_at": today
        }
        st.session_state.topics.insert(0, updated_item)
        action = "created"

    if supabase_client:
        try:
            supabase_client.table("topics").upsert(updated_item, on_conflict="title").execute()
            st.session_state.db_verified = True
            st.toast(f"☁️ Saved '{clean_title}' to Cloud!", icon="✅")
        except Exception as e:
            try:
                # Fallback if related_topics column isn't in Supabase table yet
                fallback_item = {k: v for k, v in updated_item.items() if k != "related_topics"}
                supabase_client.table("topics").upsert(fallback_item, on_conflict="title").execute()
                st.session_state.db_verified = True
                st.toast(f"☁️ Saved '{clean_title}' to Cloud!", icon="✅")
            except Exception as e2:
                st.session_state.db_error_msg = str(e2)
                st.toast("⚠️ Saved locally (Check Cloud Settings tab)", icon="⚠️")
    return action

def delete_topic(title):
    st.session_state.topics = [t for t in st.session_state.topics if t['title'] != title]
    if supabase_client:
        try:
            supabase_client.table("topics").delete().eq("title", title).execute()
            st.toast(f"🗑️ Deleted '{title}'", icon="🗑️")
        except Exception:
            pass

def upsert_glossary(term, category, context):
    clean_term = term.strip()
    existing_idx = next((i for i, g in enumerate(st.session_state.glossary) if g['term'].lower() == clean_term.lower()), None)
    item = {"term": clean_term, "category": category, "context": context}
    
    if existing_idx is not None:
        st.session_state.glossary[existing_idx] = item
    else:
        st.session_state.glossary.insert(0, item)
        
    if supabase_client:
        try:
            supabase_client.table("glossary").upsert(item, on_conflict="term").execute()
        except Exception:
            pass

# --- Fast, Non-Blocking AI Caller (No Long Loops!) ---
def ask_gemini(prompt, image_file=None, selected_model="gemini-3.8-flash"):
    if not ai_client:
        return False, "⚠️ AI Engine offline. Please verify GEMINI_API_KEY."
    
    contents = [prompt]
    if image_file:
        img = Image.open(image_file)
        contents.append(img)

    # Try selected model once, then immediate fast fallback if needed (zero sleep delay!)
    backup = "gemini-2.0-flash" if selected_model != "gemini-2.0-flash" else "gemini-3.8-flash"
    for m in [selected_model, backup]:
        try:
            response = ai_client.models.generate_content(model=m, contents=contents)
            if response and response.text:
                return True, response.text
        except Exception as e:
            last_err = str(e)
            continue
            
    return False, f"⚠️ AI Request Failed: {last_err[:180]}"

# --- PPTX Builder ---
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
        notes_tf = slide.notes_slide.notes_text_frame
        notes_tf.text = f"⏱️ Time: {s.get('time', '2 mins')}\n\n🎙️ Script:\n{s.get('speaker_notes', '')}"
        
    stream = BytesIO()
    prs.save(stream)
    stream.seek(0)
    return stream

# ================= CLEAN SIDEBAR =================
with st.sidebar:
    st.markdown("### ⚡ Nexus Control Center")
    
    # Sleek, Clean Status Badges (No ugly broken error boxes!)
    ai_badge = '<div class="status-badge badge-ok">🤖 AI Engine: Ready</div>' if ai_client else '<div class="status-badge badge-warn">🤖 AI: Missing Key</div>'
    db_badge = '<div class="status-badge badge-ok">☁️ Cloud DB: Synced</div>' if st.session_state.get("db_verified") else '<div class="status-badge badge-warn">💾 DB: Local Mode</div>'
    
    col_b1, col_b2 = st.columns(2)
    col_b1.markdown(ai_badge, unsafe_allow_html=True)
    col_b2.markdown(db_badge, unsafe_allow_html=True)

    model_choice = st.selectbox("⚡ Fast AI Model:", ["gemini-3.8-flash", "gemini-2.0-flash", "gemini-2.5-pro"], index=0)
    st.divider()
    
    st.markdown("#### 🔄 Add / Update Topic")
    topic_names = [t['title'] for t in st.session_state.topics]
    mode = st.radio("Action:", ["➕ Create New", "✏️ Edit Existing"], horizontal=True, label_visibility="collapsed")
    
    if mode == "✏️ Edit Existing" and topic_names:
        sel_edit = st.selectbox("Select Topic:", topic_names)
        curr_t = next(t for t in st.session_state.topics if t['title'] == sel_edit)
        in_title = curr_t['title']
        default_cat = curr_t.get('category', 'Project Management (PM)')
        default_stat = curr_t.get('status', 'In Progress')
        default_mast = int(curr_t.get('mastery', 50))
        existing_rels = [r.strip() for r in str(curr_t.get('related_topics', '')).split(',') if r.strip() in topic_names]
    else:
        in_title = st.text_input("Topic Title:", placeholder="e.g., 5G Site Rollout KPIs")
        default_cat = "Project Management (PM)"
        default_stat = "In Progress"
        default_mast = 50
        existing_rels = []
        
    cat_options = ["Project Management (PM)", "Technical / Engineering", "AI & Automation", "Business & Strategy"]
    stat_options = ["Not Started", "In Progress", "Completed"]
    
    in_cat = st.selectbox("Category:", cat_options, index=cat_options.index(default_cat) if default_cat in cat_options else 0)
    in_stat = st.selectbox("Status:", stat_options, index=stat_options.index(default_stat) if default_stat in stat_options else 1)
    in_mast = st.slider("Mastery (%):", 0, 100, default_mast, step=5)
    
    available_links = [t for t in topic_names if t.lower() != in_title.strip().lower()]
    linked_topics = st.multiselect("🔗 Link to Topics (Graph):", available_links, default=existing_rels)
    
    col_btn1, col_btn2 = st.columns([3, 1])
    with col_btn1:
        if st.button("💾 Save / Update", use_container_width=True, type="primary"):
            if in_title.strip():
                rel_str = ", ".join(linked_topics)
                act = upsert_topic(in_title, in_cat, in_stat, in_mast, related_topics=rel_str)
                st.success(f"Topic {act}!")
                st.rerun()
            else:
                st.warning("Enter a topic title.")
    with col_btn2:
        if mode == "✏️ Edit Existing" and st.button("🗑️", help="Delete Topic", use_container_width=True):
            delete_topic(in_title)
            st.rerun()

# ================= MAIN HEADER =================
st.markdown("""
<div class="main-header">
    <h1>🧠 Nexus | Personal AI Second Brain & Knowledge Graph</h1>
    <p>Cloud-synced learning tracker, visual concept mapper, fast document analyzer, and executive pitch studio.</p>
</div>
""", unsafe_allow_html=True)

tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
    "📊 1. Tracker", 
    "🕸️ 2. Visual Map",
    "🧠 3. Knowledge Vault", 
    "🎯 4. Pitch Studio", 
    "🎙️ 5. AI Coach & Cards",
    "⚙️ 6. Cloud Diagnostics"
])

# ================= TAB 1: TRACKER & ANALYTICS =================
with tab1:
    if st.session_state.topics:
        df = pd.DataFrame(st.session_state.topics)
        df['mastery'] = pd.to_numeric(df['mastery'], errors='coerce').fillna(0).astype(int)
        
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("📚 Total Topics", len(df))
        m2.metric("✅ Completed", len(df[df['status'] == 'Completed']))
        m3.metric("⚡ In Progress", len(df[df['status'] == 'In Progress']))
        m4.metric("🎯 Avg Mastery", f"{df['mastery'].mean():.0f}%")
        
        st.divider()
        ch1, ch2 = st.columns(2)
        with ch1:
            st.markdown("##### 📈 Mastery by Topic")
            st.bar_chart(df.set_index("title")["mastery"], color="#0284c7")
        with ch2:
            st.markdown("##### 📂 Topics per Domain")
            st.bar_chart(df['category'].value_counts(), color="#1e3a8a")

        st.markdown("### 📌 Interactive Knowledge Directory")
        f_col1, f_col2 = st.columns([2, 1])
        with f_col1:
            search_q = st.text_input("🔍 Search topics or notes...", placeholder="Type keyword...")
        with f_col2:
            filter_cat = st.multiselect("Filter by Category:", df['category'].unique(), default=df['category'].unique())
            
        filtered_df = df[df['category'].isin(filter_cat)]
        if search_q:
            filtered_df = filtered_df[
                filtered_df['title'].str.contains(search_q, case=False, na=False) | 
                filtered_df['notes'].str.contains(search_q, case=False, na=False)
            ]
            
        cols_to_show = [c for c in ['title', 'category', 'status', 'mastery', 'related_topics', 'updated_at', 'notes'] if c in filtered_df.columns]
        st.dataframe(
            filtered_df[cols_to_show],
            column_config={
                "title": "Topic Title",
                "category": "Domain",
                "status": "Status",
                "mastery": st.column_config.ProgressColumn("Mastery", format="%d%%", min_value=0, max_value=100),
                "related_topics": "🔗 Linked Topics",
                "updated_at": "Updated",
                "notes": "Key Notes"
            },
            use_container_width=True,
            hide_index=True
        )
    else:
        st.info("No topics logged yet. Use the sidebar to add your first topic!")

# ================= TAB 2: VISUAL KNOWLEDGE MAP =================
with tab2:
    st.subheader("🕸️ Interactive Second-Brain Knowledge Graph")
    g_col1, g_col2 = st.columns([3, 1])
    with g_col2:
        show_categories = st.checkbox("Show Category Hubs", value=True)
        physics_enabled = st.checkbox("Enable Bouncy Physics", value=True)
        
        if st.button("🤖 AI Auto-Connect Topics", type="primary", use_container_width=True):
            if len(st.session_state.topics) >= 2:
                with st.spinner("Discovering topic connections..."):
                    t_list_str = "\n".join([f"- {t['title']} ({t['category']})" for t in st.session_state.topics])
                    link_prompt = f"""
                    Connect related topics from this list:
                    {t_list_str}
                    Return ONLY valid JSON mapping exact topic titles to comma-separated related topic titles:
                    {{"Topic A": "Topic B"}}
                    """
                    ok, res = ask_gemini(link_prompt, selected_model=model_choice)
                    if ok:
                        try:
                            clean_j = res.strip().removeprefix("```json").removeprefix("```").removesuffix("```").strip()
                            links_map = json.loads(clean_j)
                            for t_title, rels in links_map.items():
                                matched = next((t for t in st.session_state.topics if t['title'].lower() == t_title.lower()), None)
                                if matched:
                                    upsert_topic(matched['title'], matched['category'], matched['status'], matched['mastery'], related_topics=rels)
                            st.success("Graph links updated!")
                            st.rerun()
                        except Exception as e:
                            st.warning(f"Parse error: {e}")
        st.markdown("**🎨 Legend:**\n- 🟦 Domain Hub\n- 🟢 Completed (≥80%)\n- 🟠 In Progress\n- 🔴 Early Stage")

    with g_col1:
        if AGRAPH_INSTALLED and st.session_state.topics:
            nodes, edges, added_nodes = [], [], set()
            if show_categories:
                for c in set(t.get('category', 'General') for t in st.session_state.topics):
                    hub_id = f"HUB_{c}"
                    nodes.append(Node(id=hub_id, label=f"📂 {c}", size=26, color="#1e3a8a", shape="hexagon", font={"color": "#ffffff"}))
                    added_nodes.add(hub_id)
            
            for t in st.session_state.topics:
                t_id = t['title']
                mast = int(t.get('mastery', 50))
                color = "#10b981" if t.get('status') == "Completed" or mast >= 80 else ("#f59e0b" if mast >= 40 else "#ef4444")
                if t_id not in added_nodes:
                    nodes.append(Node(id=t_id, label=f"{t_id}\n({mast}%)", size=14 + int(mast/6), color=color, shape="dot"))
                    added_nodes.add(t_id)
                if show_categories:
                    edges.append(Edge(source=f"HUB_{t.get('category', 'General')}", target=t_id, color="#94a3b8"))
                    
            for t in st.session_state.topics:
                rels = [r.strip() for r in str(t.get('related_topics', '')).split(',') if r.strip()]
                for r in rels:
                    target_match = next((x['title'] for x in st.session_state.topics if x['title'].lower() == r.lower()), None)
                    if target_match:
                        edges.append(Edge(source=t['title'], target=target_match, color="#38bdf8", width=2))
                        
            agraph(nodes=nodes, edges=edges, config=Config(width="100%", height=500, directed=False, physics=physics_enabled))

# ================= TAB 3: KNOWLEDGE VAULT =================
with tab3:
    st.subheader("📥 Fast Multi-Modal Knowledge Ingestion")
    topic_list = [t['title'] for t in st.session_state.topics]
    
    if topic_list:
        active_t_name = st.selectbox("Select Topic:", topic_list)
        active_t_obj = next(t for t in st.session_state.topics if t['title'] == active_t_name)
        
        col_v1, col_v2 = st.columns(2)
        with col_v1:
            user_notes = st.text_area("✍️ Personal Notes & Key Formulas:", value=active_t_obj.get('notes', ''), height=180)
            if st.button("💾 Save Notes Only", use_container_width=True):
                upsert_topic(active_t_name, active_t_obj['category'], active_t_obj['status'], active_t_obj['mastery'], notes=user_notes)
                st.success("Notes saved!")
                
        with col_v2:
            up_file = st.file_uploader("📄 Upload PDF or Slide Image (PNG/JPG):", type=['pdf', 'png', 'jpg', 'jpeg'])
            extract_glossary = st.checkbox("Auto-extract terms to Glossary", value=True)
            
            if st.button("⚡ Instant AI Summarize", type="primary", use_container_width=True):
                if not up_file and not user_notes:
                    st.warning("Upload a file or write notes first.")
                else:
                    with st.spinner(f"Fast-analyzing with {model_choice} (takes ~3 seconds)..."):
                        extracted_text, img_to_send = "", None
                        if up_file:
                            if up_file.name.lower().endswith('.pdf'):
                                reader = PyPDF2.PdfReader(up_file)
                                for page in reader.pages[:8]: # Limit to first 8 pages for blazing speed!
                                    extracted_text += (page.extract_text() or "") + "\n"
                            else:
                                img_to_send = up_file
                                
                        prompt = f"""
                        Summarize this material for '{active_t_name}'.
                        Notes: {user_notes}
                        Content: {extracted_text[:4500]}
                        
                        Format:
                        ### 📌 Executive Summary (4 bullets)
                        ### ⚙️ Key Engineering/PM Takeaways
                        ### 📖 Key Terms (Strictly 3 lines: TERM: [name] | DEF: [short definition])
                        """
                        ok, ai_res = ask_gemini(prompt, image_file=img_to_send, selected_model=model_choice)
                        if ok:
                            upsert_topic(active_t_name, active_t_obj['category'], active_t_obj['status'], active_t_obj['mastery'], notes=user_notes, ai_summary=ai_res)
                            if extract_glossary and "TERM:" in ai_res:
                                for line in ai_res.split('\n'):
                                    if "TERM:" in line and "| DEF:" in line:
                                        parts = line.split("| DEF:")
                                        upsert_glossary(parts[0].replace("TERM:", "").replace("*", "").replace("-", "").strip(), active_t_obj['category'], parts[1].strip())
                            st.rerun()
                        else:
                            st.error(ai_res)

        saved_sum = active_t_obj.get('ai_summary', '')
        if saved_sum and not saved_sum.startswith("AI Error:"):
            st.divider()
            st.markdown("#### 🤖 Saved AI Synthesis")
            st.info(saved_sum)

# ================= TAB 4: PITCH STUDIO =================
with tab4:
    st.subheader("🎯 Automated Executive Presentation Studio")
    p_col1, p_col2, p_col3 = st.columns(3)
    with p_col1:
        pitch_topic = st.selectbox("Select Topic:", topic_list if topic_list else ["General"], key="pitch_sel")
    with p_col2:
        audience = st.selectbox("Target Audience:", ["Executive Board / PMO", "Engineering Deep-Dive", "Oral Assessment Panel"])
    with p_col3:
        slide_count = st.slider("Slides:", 3, 6, 4)
        
    if st.button("✨ Generate Slide Deck + Script", type="primary"):
        t_obj = next((t for t in st.session_state.topics if t['title'] == pitch_topic), {})
        with st.spinner("Generating presentation..."):
            pitch_prompt = f"""
            Create a {slide_count}-slide presentation on '{pitch_topic}' for '{audience}'.
            Notes: {t_obj.get('notes', '')[:1500]}
            Return ONLY a JSON array:
            [{{"title": "Slide Title", "bullets": ["Point 1", "Point 2"], "speaker_notes": "Script...", "time": "2 mins"}}]
            """
            ok, raw_json = ask_gemini(pitch_prompt, selected_model=model_choice)
            if ok:
                try:
                    clean_json = raw_json.strip().removeprefix("```json").removeprefix("```").removesuffix("```").strip()
                    slides_data = json.loads(clean_json)
                    cols = st.columns(2)
                    for i, s in enumerate(slides_data):
                        with cols[i % 2]:
                            with st.container(border=True):
                                st.markdown(f"#### 📊 Slide {i+1}: {s.get('title')} `{s.get('time')}`")
                                for b in s.get("bullets", []): st.markdown(f"- {b}")
                                with st.expander("🎙️ Speaker Script"): st.write(s.get("speaker_notes"))
                    ppt_bytes = build_advanced_pptx(pitch_topic, audience, slides_data)
                    st.download_button("📥 Download PowerPoint (.pptx)", data=ppt_bytes, file_name=f"{pitch_topic}_Deck.pptx", type="primary", use_container_width=True)
                except Exception:
                    st.write(raw_json)
            else:
                st.error(raw_json)

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
                    ok, res = ask_gemini(f"Evaluate this explanation of '{coach_topic}': '{user_exp}'. Give a score /10, missing terms, and a polished 3-sentence version.", selected_model=model_choice)
                    st.markdown(res) if ok else st.error(res)
        elif coach_mode == "🃏 Smart Flashcards":
            if st.button("⚡ Generate 4 Flashcards", type="primary"):
                with st.spinner("Creating cards..."):
                    ok, res = ask_gemini(f"Return ONLY a JSON array of 4 flashcards for '{coach_topic}': [{{\"q\":\"Question?\",\"a\":\"Answer\"}}]", selected_model=model_choice)
                    if ok:
                        try:
                            st.session_state.flashcards = json.loads(res.strip().removeprefix("```json").removeprefix("```").removesuffix("```").strip())
                        except Exception: st.write(res)
            for idx, fc in enumerate(st.session_state.get("flashcards", [])):
                st.markdown(f'<div class="flashcard"><b>Q{idx+1}: {fc.get("q")}</b></div>', unsafe_allow_html=True)
                with st.expander("👀 Reveal Answer"): st.success(fc.get('a'))
        else:
            if st.button("🎲 Generate 3 Questions", type="primary"):
                with st.spinner("Generating quiz..."):
                    ok, res = ask_gemini(f"Create 3 multiple-choice questions on '{coach_topic}' with answers at the end.", selected_model=model_choice)
                    st.markdown(res) if ok else st.error(res)

    with c_right:
        st.subheader("📖 Smart Glossary")
        with st.expander("➕ Add Term"):
            g_t = st.text_input("Term:")
            g_c = st.selectbox("Domain:", ["PM", "Technical", "Business"])
            g_d = st.text_area("Definition:")
            if st.button("Save Term") and g_t:
                upsert_glossary(g_t, g_c, g_d)
                st.rerun()
        if st.session_state.glossary:
            st.dataframe(pd.DataFrame(st.session_state.glossary)[['term', 'category', 'context']], use_container_width=True, hide_index=True)

# ================= TAB 6: CLOUD DIAGNOSTICS & BACKUP =================
with tab6:
    st.subheader("⚙️ Cloud Database Diagnostics & Backup")
    st.write(f"**Sanitized Supabase API URL:** `{SUPABASE_URL or 'Not Set'}`")
    
    if st.session_state.get("db_verified"):
        st.success("✅ Your Supabase Cloud Database is properly connected and syncing automatically!")
    else:
        st.warning("⚠️ Currently running in Local Memory Mode. See diagnostic message below:")
        if st.session_state.get("db_error_msg"):
            st.code(st.session_state.get("db_error_msg"))
        st.markdown("""
        **How to fix `SUPABASE_URL` if it fails:**
        1. In Supabase, go to **Project Settings (⚙️) -> API**.
        2. Copy the **Project URL** (looks like `[https://abcdefghijk.supabase.co](https://abcdefghijk.supabase.co)`).
        3. Make sure you ran the SQL table creation script in **SQL Editor**.
        """)
        if st.button("🔄 Test & Reconnect Cloud DB", type="primary"):
            if load_data_from_cloud():
                st.success("Connected successfully!")
                st.rerun()
                
    st.divider()
    st.markdown("#### 📦 Manual JSON Backup & Restore")
    backup_data = json.dumps({"topics": st.session_state.topics, "glossary": st.session_state.glossary}, indent=2)
    st.download_button("⬇️ Download Full Backup (.json)", data=backup_data, file_name="nexus_backup.json", mime="application/json")
