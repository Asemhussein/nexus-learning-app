import streamlit as st
import pandas as pd
from datetime import datetime
import os
import json
import time
from io import BytesIO
import PyPDF2
from PIL import Image
from pptx import Presentation
from pptx.util import Pt
from google import genai

# --- Optional Imports (Supabase & Visual Graph) ---
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

# --- Page Config & Styling ---
st.set_page_config(page_title="Nexus | AI Second Brain", page_icon="🧠", layout="wide")

st.markdown("""
<style>
    .main-header {
        background: linear-gradient(135deg, #0f172a 0%, #1e3a8a 60%, #0284c7 100%);
        padding: 1.5rem 2rem;
        border-radius: 14px;
        color: white;
        margin-bottom: 1.5rem;
        box-shadow: 0 4px 20px rgba(0,0,0,0.18);
    }
    .main-header h1 { color: #ffffff; margin: 0; font-size: 2rem; font-weight: 800; }
    .main-header p { color: #e2e8f0; margin: 0.3rem 0 0 0; font-size: 0.95rem; }
    div[data-testid="stMetric"] {
        background-color: rgba(30, 41, 59, 0.04);
        border: 1px solid rgba(148, 163, 184, 0.3);
        padding: 0.9rem 1.2rem;
        border-radius: 12px;
    }
    .flashcard {
        background: linear-gradient(135deg, #1e293b 0%, #0f172a 100%);
        color: #f8fafc;
        padding: 1.5rem;
        border-radius: 12px;
        border-left: 5px solid #38bdf8;
        margin-bottom: 1rem;
    }
</style>
""", unsafe_allow_html=True)

# --- Secrets Helper ---
def get_secret(key):
    if key in os.environ and os.environ[key]:
        return os.environ[key].strip()
    try:
        if key in st.secrets:
            return str(st.secrets[key]).strip()
    except Exception:
        pass
    return None

GEMINI_API_KEY = get_secret("GEMINI_API_KEY")
SUPABASE_URL = get_secret("SUPABASE_URL")
SUPABASE_KEY = get_secret("SUPABASE_KEY")

ai_client = genai.Client(api_key=GEMINI_API_KEY) if GEMINI_API_KEY else None

supabase_client = None
if SUPABASE_INSTALLED and SUPABASE_URL and SUPABASE_KEY:
    try:
        supabase_client = create_client(SUPABASE_URL, SUPABASE_KEY)
    except Exception as e:
        st.sidebar.error(f"Supabase Init Error: {e}")

# --- Database Sync Functions (With Real Verification) ---
def load_data_from_cloud():
    if supabase_client:
        try:
            res_t = supabase_client.table("topics").select("*").order("id", desc=True).execute()
            res_g = supabase_client.table("glossary").select("*").order("id", desc=True).execute()
            st.session_state.topics = res_t.data if res_t.data else []
            st.session_state.glossary = res_g.data if res_g.data else []
            st.session_state.db_verified = True
            return True
        except Exception as e:
            st.session_state.db_verified = False
            st.session_state.db_error_msg = str(e)
            return False
    return False

if 'data_initialized' not in st.session_state:
    success = load_data_from_cloud()
    if not success:
        st.session_state.topics = [
            {"title": "Telecom Rollout KPIs & Risk Mgmt", "category": "Project Management (PM)", "status": "Completed", "mastery": 90, "notes": "Critical path & vendor SLA tracking.", "ai_summary": "", "related_topics": "5G Site Deployment", "updated_at": "2026-10-28"},
            {"title": "NVDLA Architecture & LUT Optimization", "category": "Technical / Engineering", "status": "Completed", "mastery": 95, "notes": "Hardware accelerator lookup-table logic.", "ai_summary": "", "related_topics": "FPGA DSP Architecture", "updated_at": "2026-10-28"}
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

    # Cloud Save with Explicit Error Reporting
    if supabase_client:
        try:
            supabase_client.table("topics").upsert(updated_item, on_conflict="title").execute()
            st.toast(f"☁️ Synced '{clean_title}' to Supabase Cloud!", icon="✅")
        except Exception as e:
            # Fallback if 'related_topics' column wasn't created in SQL yet
            try:
                fallback_item = {k: v for k, v in updated_item.items() if k != "related_topics"}
                supabase_client.table("topics").upsert(fallback_item, on_conflict="title").execute()
                st.toast("☁️ Saved to Cloud (Run new SQL script to enable Graph links column)", icon="⚠️")
            except Exception as e2:
                st.error(f"❌ Cloud Save Failed! Please run the SQL script in Supabase. Details: {e2}")
    return action

def delete_topic(title):
    st.session_state.topics = [t for t in st.session_state.topics if t['title'] != title]
    if supabase_client:
        try:
            supabase_client.table("topics").delete().eq("title", title).execute()
            st.toast(f"🗑️ Deleted '{title}' from Cloud.", icon="🗑️")
        except Exception as e:
            st.error(f"Cloud Delete Error: {e}")

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
        except Exception as e:
            st.error(f"Glossary Cloud Sync Error: {e}")

# --- Smart AI Caller with Auto-Retry & Fallback for 503 Errors ---
def ask_gemini(prompt, image_file=None, preferred_model="gemini-3.8-flash"):
    if not ai_client:
        return False, "⚠️ AI Engine offline. Please verify GEMINI_API_KEY."
    
    # Fallback chain if a model has 503 high demand
    fallback_models = [preferred_model, "gemini-3.8-flash", "gemini-2.5-pro", "gemini-2.0-flash"]
    models_to_try = list(dict.fromkeys(fallback_models))
    
    contents = [prompt]
    if image_file:
        img = Image.open(image_file)
        contents.append(img)

    last_err = ""
    for model_name in models_to_try:
        for attempt in range(2): # 2 retries per model
            try:
                response = ai_client.models.generate_content(model=model_name, contents=contents)
                return True, response.text
            except Exception as e:
                err_str = str(e)
                last_err = err_str
                if "503" in err_str or "UNAVAILABLE" in err_str or "429" in err_str:
                    time.sleep(1.5) # Wait briefly and retry
                    continue
                elif "404" in err_str or "NOT_FOUND" in err_str:
                    break # Move immediately to next model in fallback list
                else:
                    return False, f"AI Error: {err_str}"
                    
    return False, f"⚠️ Google AI servers are currently overloaded (503). Please try again in a few seconds. (Details: {last_err})"

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

# ================= SIDEBAR =================
with st.sidebar:
    st.markdown("## ⚡ Nexus Control Center")
    
    c1, c2 = st.columns(2)
    with c1:
        st.success("🤖 AI: Online") if ai_client else st.error("🤖 AI: Offline")
    with c2:
        if supabase_client and st.session_state.get("db_verified", False):
            st.success("☁️ DB: Synced")
        elif supabase_client:
            st.warning("⚠️ DB: Auth/SQL Err")
        else:
            st.warning("💾 DB: Local")

    if supabase_client and not st.session_state.get("db_verified", False):
        st.error(f"DB Error: Run the SQL script in Supabase! ({st.session_state.get('db_error_msg', '')[:60]})")

    if st.button("🔄 Force Sync from Cloud", use_container_width=True):
        if load_data_from_cloud():
            st.toast("Cloud data loaded!", icon="☁️")
            st.rerun()

    model_choice = st.selectbox("🧠 Primary AI Model:", ["gemini-3.8-flash", "gemini-2.5-pro"], index=0)
    st.divider()
    
    # Smart Add / Update Topic with Connections
    st.subheader("🔄 Add / Update Topic")
    topic_names = [t['title'] for t in st.session_state.topics]
    mode = st.radio("Mode:", ["➕ Create New", "✏️ Edit Existing"], horizontal=True)
    
    if mode == "✏️ Edit Existing" and topic_names:
        sel_edit = st.selectbox("Select Topic:", topic_names)
        curr_t = next(t for t in st.session_state.topics if t['title'] == sel_edit)
        in_title = curr_t['title']
        default_cat = curr_t.get('category', 'Project Management (PM)')
        default_stat = curr_t.get('status', 'In Progress')
        default_mast = int(curr_t.get('mastery', 50))
        existing_rels = [r.strip() for r in curr_t.get('related_topics', '').split(',') if r.strip() in topic_names]
    else:
        in_title = st.text_input("Topic Title:", placeholder="e.g., Agile Sprint Planning")
        default_cat = "Project Management (PM)"
        default_stat = "In Progress"
        default_mast = 50
        existing_rels = []
        
    cat_options = ["Project Management (PM)", "Technical / Engineering", "AI & Automation", "Business & Strategy"]
    stat_options = ["Not Started", "In Progress", "Completed"]
    
    in_cat = st.selectbox("Category:", cat_options, index=cat_options.index(default_cat) if default_cat in cat_options else 0)
    in_stat = st.selectbox("Status:", stat_options, index=stat_options.index(default_stat) if default_stat in stat_options else 1)
    in_mast = st.slider("Mastery (%):", 0, 100, default_mast, step=5)
    
    # Connect to other topics for the Knowledge Graph!
    available_links = [t for t in topic_names if t.lower() != in_title.strip().lower()]
    linked_topics = st.multiselect("🔗 Link to Related Topics (Graph):", available_links, default=existing_rels)
    
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
    <p>Cloud-synced learning tracker, visual concept mapper, multi-modal document analyzer, and executive pitch studio.</p>
</div>
""", unsafe_allow_html=True)

tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📊 1. Tracker & Analytics", 
    "🕸️ 2. Visual Knowledge Map",
    "🧠 3. Knowledge Vault (OCR)", 
    "🎯 4. Pitch Studio", 
    "🎙️ 5. AI Coach & Flashcards"
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
            cat_counts = df['category'].value_counts()
            st.bar_chart(cat_counts, color="#1e3a8a")

        st.markdown("### 📌 Interactive Knowledge Directory")
        f_col1, f_col2 = st.columns([2, 1])
        with f_col1:
            search_q = st.text_input("🔍 Search topics, notes, or connections...", placeholder="Type keyword...")
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
    st.caption("Visualize how your Engineering, PM, and AI concepts interconnect. Drag nodes to explore!")
    
    g_col1, g_col2 = st.columns([3, 1])
    with g_col2:
        show_categories = st.checkbox("Show Category Hubs", value=True)
        physics_enabled = st.checkbox("Enable Bouncy Physics", value=True)
        
        if st.button("🤖 AI Auto-Discover Links", type="primary", use_container_width=True):
            if len(st.session_state.topics) >= 2:
                with st.spinner("AI is finding hidden connections between your topics..."):
                    t_list_str = "\n".join([f"- {t['title']} ({t['category']}): {t.get('notes','')}" for t in st.session_state.topics])
                    link_prompt = f"""
                    Given these topics in my knowledge base:
                    {t_list_str}
                    Identify meaningful conceptual connections between them.
                    Return ONLY a JSON object mapping exact topic titles to a comma-separated string of related exact topic titles from the list:
                    {{"Topic A": "Topic B, Topic C"}}
                    """
                    ok, res = ask_gemini(link_prompt, preferred_model=model_choice)
                    if ok:
                        try:
                            clean_j = res.strip().removeprefix("```json").removeprefix("```").removesuffix("```").strip()
                            links_map = json.loads(clean_j)
                            for t_title, rels in links_map.items():
                                matched = next((t for t in st.session_state.topics if t['title'].lower() == t_title.lower()), None)
                                if matched:
                                    upsert_topic(matched['title'], matched['category'], matched['status'], matched['mastery'], related_topics=rels)
                            st.success("AI updated your Knowledge Graph links!")
                            st.rerun()
                        except Exception as e:
                            st.warning(f"Could not parse AI links automatically: {e}")
            else:
                st.warning("Add at least 2 topics first.")

        st.markdown("""
        **🎨 Node Legend:**
        - 🟦 **Domain Hub** (Category)
        - 🟢 **Completed** (Mastery ≥ 80%)
        - 🟠 **In Progress**
        - 🔴 **Early Stage**
        """)

    with g_col1:
        if AGRAPH_INSTALLED and st.session_state.topics:
            nodes = []
            edges = []
            added_nodes = set()
            
            # Category Hub Nodes
            if show_categories:
                cats = set(t.get('category', 'General') for t in st.session_state.topics)
                for c in cats:
                    hub_id = f"HUB_{c}"
                    nodes.append(Node(id=hub_id, label=f"📂 {c}", size=28, color="#1e3a8a", shape="hexagon", font={"color": "#ffffff", "size": 14}))
                    added_nodes.add(hub_id)
            
            # Topic Nodes
            for t in st.session_state.topics:
                t_id = t['title']
                mast = int(t.get('mastery', 50))
                color = "#10b981" if t.get('status') == "Completed" or mast >= 80 else ("#f59e0b" if mast >= 40 else "#ef4444")
                size = 15 + int(mast / 6)
                
                if t_id not in added_nodes:
                    nodes.append(Node(id=t_id, label=f"{t_id}\n({mast}%)", size=size, color=color, shape="dot"))
                    added_nodes.add(t_id)
                
                if show_categories:
                    edges.append(Edge(source=f"HUB_{t.get('category', 'General')}", target=t_id, color="#94a3b8", type="CURVE_SMOOTH"))
                    
            # Cross-Topic Edges
            for t in st.session_state.topics:
                rels = [r.strip() for r in str(t.get('related_topics', '')).split(',') if r.strip()]
                for r in rels:
                    target_match = next((x['title'] for x in st.session_state.topics if x['title'].lower() == r.lower()), None)
                    if target_match:
                        edges.append(Edge(source=t['title'], target=target_match, label="related", color="#38bdf8", width=2))
                        
            config = Config(width="100%", height=520, directed=False, physics=physics_enabled, hierarchical=False)
            agraph(nodes=nodes, edges=edges, config=config)
        elif not AGRAPH_INSTALLED:
            st.warning("Install `streamlit-agraph` in requirements.txt to view the interactive 2D map.")

# ================= TAB 3: KNOWLEDGE VAULT =================
with tab3:
    st.subheader("📥 Multi-Modal Knowledge Ingestion (PDFs, Slide Images & Notes)")
    topic_list = [t['title'] for t in st.session_state.topics]
    
    if not topic_list:
        st.warning("Create a topic in the sidebar first.")
    else:
        active_t_name = st.selectbox("Select Topic to Enrich:", topic_list)
        active_t_obj = next(t for t in st.session_state.topics if t['title'] == active_t_name)
        
        col_v1, col_v2 = st.columns(2)
        with col_v1:
            user_notes = st.text_area("✍️ Personal Notes & Key Formulas:", value=active_t_obj.get('notes', ''), height=200)
            if st.button("💾 Save Personal Notes", use_container_width=True):
                upsert_topic(active_t_name, active_t_obj['category'], active_t_obj['status'], active_t_obj['mastery'], notes=user_notes)
                st.success("Notes saved to Cloud!")
                
        with col_v2:
            up_file = st.file_uploader("📄 Upload Lecture PDF or Slide Image (PNG/JPG):", type=['pdf', 'png', 'jpg', 'jpeg'])
            extract_glossary = st.checkbox("Auto-extract technical/PM terms to Glossary", value=True)
            
            if st.button("🚀 Analyze & Summarize with AI", type="primary", use_container_width=True):
                if not up_file and not user_notes:
                    st.warning("Please upload a file or write notes first.")
                else:
                    with st.spinner("AI is reading your material (Auto-Retry enabled for high demand)..."):
                        extracted_text = ""
                        img_to_send = None
                        
                        if up_file:
                            if up_file.name.lower().endswith('.pdf'):
                                reader = PyPDF2.PdfReader(up_file)
                                for page in reader.pages:
                                    extracted_text += (page.extract_text() or "") + "\n"
                            else:
                                img_to_send = up_file
                                
                        prompt = f"""
                        Analyze this material for the topic '{active_t_name}'.
                        User Notes: {user_notes}
                        Document Content: {extracted_text[:7000]}
                        
                        Provide a structured response:
                        ### 📌 Executive Summary
                        (4 clear bullet points)
                        ### ⚙️ Core Engineering / PM Takeaways
                        (Practical insights & metrics)
                        ### 📖 Extracted Terms
                        (Format strictly 3 lines as -> TERM: [name] | DEF: [1-sentence definition])
                        """
                        ok, ai_res = ask_gemini(prompt, image_file=img_to_send, preferred_model=model_choice)
                        
                        if ok:
                            upsert_topic(active_t_name, active_t_obj['category'], active_t_obj['status'], active_t_obj['mastery'], notes=user_notes, ai_summary=ai_res)
                            if extract_glossary and "TERM:" in ai_res:
                                for line in ai_res.split('\n'):
                                    if "TERM:" in line and "| DEF:" in line:
                                        parts = line.split("| DEF:")
                                        t_name = parts[0].replace("TERM:", "").replace("*", "").replace("-", "").strip()
                                        t_def = parts[1].strip()
                                        upsert_glossary(t_name, active_t_obj['category'], t_def)
                            st.success("✅ Analysis complete & synced to Cloud!")
                            st.rerun()
                        else:
                            # Do NOT overwrite saved summary with a 503 error!
                            st.error(ai_res)

        saved_sum = active_t_obj.get('ai_summary', '')
        if saved_sum and not saved_sum.startswith("AI Error:"):
            st.divider()
            col_s1, col_s2 = st.columns([5, 1])
            with col_s1:
                st.markdown("#### 🤖 Saved AI Synthesis & Summary")
            with col_s2:
                if st.button("🗑️ Clear Summary"):
                    upsert_topic(active_t_name, active_t_obj['category'], active_t_obj['status'], active_t_obj['mastery'], ai_summary="")
                    st.rerun()
            st.info(saved_sum)

# ================= TAB 4: PITCH STUDIO =================
with tab4:
    st.subheader("🎯 Automated Executive Presentation & Speaker Script Studio")
    p_col1, p_col2, p_col3 = st.columns(3)
    with p_col1:
        pitch_topic = st.selectbox("Select Topic:", topic_list if topic_list else ["General Topic"], key="pitch_sel")
    with p_col2:
        audience = st.selectbox("Target Audience:", ["Executive Board / PMO (KPIs & ROI)", "Engineering Team (Architecture & Deep-Dive)", "Oral Assessment / Interview Panel"])
    with p_col3:
        slide_count = st.slider("Number of Slides:", 3, 8, 4)
        
    if st.button("✨ Generate Slide Deck + Speaker Notes", type="primary"):
        t_obj = next((t for t in st.session_state.topics if t['title'] == pitch_topic), {})
        context_data = f"Notes: {t_obj.get('notes', '')}\nSummary: {t_obj.get('ai_summary', '')}"
        
        with st.spinner("Architecting slides and speaker scripts..."):
            pitch_prompt = f"""
            Create a {slide_count}-slide presentation on '{pitch_topic}' for '{audience}'.
            Context: {context_data}
            Return ONLY a valid JSON array of objects:
            [
              {{"title": "Slide Title", "bullets": ["Point 1", "Point 2", "Point 3"], "speaker_notes": "Exact English script to say...", "time": "2 mins"}}
            ]
            """
            ok, raw_json = ask_gemini(pitch_prompt, preferred_model=model_choice)
            if ok:
                try:
                    clean_json = raw_json.strip().removeprefix("```json").removeprefix("```").removesuffix("```").strip()
                    slides_data = json.loads(clean_json)
                    st.success("✅ Presentation Deck & Script Ready!")
                    cols = st.columns(2)
                    for i, s in enumerate(slides_data):
                        with cols[i % 2]:
                            with st.container(border=True):
                                st.markdown(f"#### 📊 Slide {i+1}: {s.get('title')} `⏱️ {s.get('time')}`")
                                for b in s.get("bullets", []):
                                    st.markdown(f"- {b}")
                                with st.expander("🎙️ View Speaker Script"):
                                    st.write(s.get("speaker_notes"))
                                    
                    ppt_bytes = build_advanced_pptx(pitch_topic, audience, slides_data)
                    st.download_button(
                        "📥 Download PowerPoint (.pptx) with Speaker Notes",
                        data=ppt_bytes,
                        file_name=f"{pitch_topic.replace(' ', '_')}_Deck.pptx",
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                        type="primary",
                        use_container_width=True
                    )
                except Exception:
                    st.write(raw_json)
            else:
                st.error(raw_json)

# ================= TAB 5: AI COACH, FLASHCARDS & GLOSSARY =================
with tab5:
    c_left, c_right = st.columns([3, 2])
    
    with c_left:
        st.subheader("🎙️ Active Recall, Flashcards & Oral Coach")
        coach_mode = st.radio("Select Practice Mode:", ["🧠 Feynman Oral Evaluator", "🃏 AI Smart Flashcards", "📝 Scenario Quiz"], horizontal=True)
        coach_topic = st.selectbox("Topic to Practice:", topic_list if topic_list else ["General"], key="coach_sel")
        t_obj_c = next((t for t in st.session_state.topics if t['title'] == coach_topic), {})
        
        if coach_mode == "🧠 Feynman Oral Evaluator":
            user_exp = st.text_area("Explain the concept in English (as if in an interview or PMO review):", height=130)
            if st.button("🔍 Evaluate My Delivery", type="primary"):
                if user_exp:
                    with st.spinner("Evaluating technical accuracy & English delivery..."):
                        eval_prompt = f"""
                        Act as a senior Engineering & PMO interviewer. Evaluate this candidate's explanation of '{coach_topic}':
                        "{user_exp}"
                        Provide:
                        1. Score out of 10 (Clarity, Structure & Technical Accuracy).
                        2. Missing industry keywords/jargon they should include.
                        3. Polished, executive-level version of their answer (3-4 sentences).
                        """
                        ok, res = ask_gemini(eval_prompt, preferred_model=model_choice)
                        st.markdown(res) if ok else st.error(res)
                        
        elif coach_mode == "🃏 AI Smart Flashcards":
            if st.button("⚡ Generate 4 Active Recall Flashcards", type="primary"):
                with st.spinner("Generating flashcards..."):
                    fc_prompt = f"""
                    Generate 4 high-impact study flashcards for '{coach_topic}'.
                    Context: {t_obj_c.get('notes', '')}
                    Return ONLY a valid JSON array:
                    [{{"q": "Question?", "a": "Concise Answer"}}]
                    """
                    ok, res = ask_gemini(fc_prompt, preferred_model=model_choice)
                    if ok:
                        try:
                            clean_fc = res.strip().removeprefix("```json").removeprefix("```").removesuffix("```").strip()
                            st.session_state.flashcards = json.loads(clean_fc)
                        except Exception:
                            st.write(res)
                    else:
                        st.error(res)
                        
            if st.session_state.get("flashcards"):
                for idx, fc in enumerate(st.session_state.flashcards):
                    st.markdown(f"""<div class="flashcard"><b>Q{idx+1}: {fc.get('q')}</b></div>""", unsafe_allow_html=True)
                    with st.expander("👀 Flip Card to Reveal Answer"):
                        st.success(fc.get('a'))

        else:
            if st.button("🎲 Generate 3 Scenario Questions", type="primary"):
                with st.spinner("Creating quiz..."):
                    q_prompt = f"Create 3 practical multiple-choice scenario questions about '{coach_topic}' with answers and brief explanations at the bottom."
                    ok, res = ask_gemini(q_prompt, preferred_model=model_choice)
                    st.markdown(res) if ok else st.error(res)

    with c_right:
        st.subheader("📖 Smart Terminology Glossary")
        with st.expander("➕ Manually Add Term"):
            g_t = st.text_input("Term:")
            g_c = st.selectbox("Domain:", ["PM", "Technical", "Business"])
            g_d = st.text_area("Definition / Context:")
            if st.button("Save Term"):
                if g_t:
                    upsert_glossary(g_t, g_c, g_d)
                    st.success("Term synced!")
                    st.rerun()
                    
        if st.session_state.glossary:
            st.dataframe(pd.DataFrame(st.session_state.glossary)[['term', 'category', 'context']], use_container_width=True, hide_index=True)
