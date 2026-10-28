import streamlit as st
import pandas as pd
from datetime import datetime
import os
import json
from io import BytesIO
import PyPDF2
from PIL import Image
from pptx import Presentation
from pptx.util import Inches, Pt
from google import genai

# --- Try importing Supabase ---
try:
    from supabase import create_client, Client
    SUPABASE_INSTALLED = True
except ImportError:
    SUPABASE_INSTALLED = False

# --- Page Configuration & Modern UI CSS ---
st.set_page_config(page_title="Nexus | AI Second Brain", page_icon="🧠", layout="wide")

st.markdown("""
<style>
    .main-header {
        background: linear-gradient(90deg, #0f172a 0%, #1e3a8a 100%);
        padding: 1.5rem 2rem;
        border-radius: 14px;
        color: white;
        margin-bottom: 1.5rem;
        box-shadow: 0 4px 15px rgba(0,0,0,0.15);
    }
    .main-header h1 { color: #f8fafc; margin: 0; font-size: 2rem; font-weight: 800; }
    .main-header p { color: #94a3b8; margin: 0.3rem 0 0 0; font-size: 1rem; }
    
    div[data-testid="stMetric"] {
        background-color: rgba(30, 41, 59, 0.04);
        border: 1px solid rgba(148, 163, 184, 0.25);
        padding: 1rem 1.2rem;
        border-radius: 12px;
    }
    .stTabs [data-baseweb="tab-list"] { gap: 10px; }
    .stTabs [data-baseweb="tab"] {
        border-radius: 8px 8px 0 0;
        padding: 10px 20px;
        font-weight: 600;
    }
</style>
""", unsafe_allow_html=True)

# --- Helper: Get Secrets from Env or Streamlit Secrets ---
def get_secret(key):
    if key in os.environ:
        return os.environ[key]
    try:
        if key in st.secrets:
            return st.secrets[key]
    except Exception:
        pass
    return None

GEMINI_API_KEY = get_secret("GEMINI_API_KEY")
SUPABASE_URL = get_secret("SUPABASE_URL")
SUPABASE_KEY = get_secret("SUPABASE_KEY")

# --- Initialize AI & Database Connections ---
ai_client = genai.Client(api_key=GEMINI_API_KEY) if GEMINI_API_KEY else None

supabase_client = None
if SUPABASE_INSTALLED and SUPABASE_URL and SUPABASE_KEY:
    try:
        supabase_client = create_client(SUPABASE_URL, SUPABASE_KEY)
    except Exception as e:
        st.sidebar.error(f"DB Connection Error: {e}")

# --- Load & Sync Data (Cloud + Local Fallback) ---
def load_data():
    if supabase_client:
        try:
            res_t = supabase_client.table("topics").select("*").order("id", desc=True).execute()
            res_g = supabase_client.table("glossary").select("*").order("id", desc=True).execute()
            st.session_state.topics = res_t.data if res_t.data else []
            st.session_state.glossary = res_g.data if res_g.data else []
            return
        except Exception:
            pass
    # Fallback to Session State if Cloud DB isn't connected yet
    if 'topics' not in st.session_state:
        st.session_state.topics = [
            {"title": "Telecom Rollout KPIs & Risk Mgmt", "category": "Project Management (PM)", "status": "Completed", "mastery": 90, "notes": "Critical path & vendor SLA tracking.", "ai_summary": "", "updated_at": "2026-10-28"},
            {"title": "NVDLA Architecture & LUT Optimization", "category": "Technical / Engineering", "status": "Completed", "mastery": 95, "notes": "Hardware accelerator lookup-table logic.", "ai_summary": "", "updated_at": "2026-10-28"}
        ]
    if 'glossary' not in st.session_state:
        st.session_state.glossary = [
            {"term": "Critical Path", "category": "PM", "context": "The longest sequence of tasks that determines the project completion date."}
        ]

if 'data_initialized' not in st.session_state:
    load_data()
    st.session_state.data_initialized = True

# --- Smart Upsert Functions (Prevents Duplicates!) ---
def upsert_topic(title, category, status, mastery, notes=None, ai_summary=None):
    clean_title = title.strip()
    today = datetime.now().strftime("%Y-%m-%d")
    
    existing_idx = next((i for i, t in enumerate(st.session_state.topics) if t['title'].lower() == clean_title.lower()), None)
    
    if existing_idx is not None:
        curr = st.session_state.topics[existing_idx]
        clean_title = curr['title'] # Keep original casing
        updated_item = {
            "title": clean_title,
            "category": category or curr.get("category", "General"),
            "status": status or curr.get("status", "In Progress"),
            "mastery": mastery if mastery is not None else curr.get("mastery", 50),
            "notes": notes if notes is not None else curr.get("notes", ""),
            "ai_summary": ai_summary if ai_summary is not None else curr.get("ai_summary", ""),
            "updated_at": today
        }
        st.session_state.topics[existing_idx] = updated_item
        action = "updated"
    else:
        updated_item = {
            "title": clean_title,
            "category": category or "General",
            "status": status or "In Progress",
            "mastery": mastery if mastery is not None else 50,
            "notes": notes or "",
            "ai_summary": ai_summary or "",
            "updated_at": today
        }
        st.session_state.topics.insert(0, updated_item)
        action = "created"

    if supabase_client:
        try:
            supabase_client.table("topics").upsert(updated_item, on_conflict="title").execute()
        except Exception as e:
            st.toast(f"Cloud Sync Warning: {e}", icon="⚠️")
    return action

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

# --- AI Helper Function (Updated to gemini-3.8-flash) ---
def ask_gemini(prompt, image_File=None, model_name="gemini-3.8-flash"):
    if not ai_client:
        return "⚠️ AI Engine offline. Please verify GEMINI_API_KEY."
    try:
        contents = [prompt]
        if image_File:
            img = Image.open(image_File)
            contents.append(img)
        response = ai_client.models.generate_content(model=model_name, contents=contents)
        return response.text
    except Exception as e:
        return f"AI Error: {e}"

# --- PPTX Generator with Speaker Notes ---
def build_advanced_pptx(topic, audience, slides_json):
    prs = Presentation()
    
    # Title Slide
    s0 = prs.slides.add_slide(prs.slide_layouts[0])
    s0.shapes.title.text = topic
    s0.placeholders[1].text = f"Target Audience: {audience}\nGenerated by Nexus AI Engine"
    
    for s in slides_json:
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        slide.shapes.title.text = s.get("title", "Slide")
        tf = slide.placeholders[1].text_frame
        tf.word_wrap = True
        
        bullets = s.get("bullets", [])
        for idx, b in enumerate(bullets):
            p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
            p.text = str(b)
            p.font.size = Pt(20)
            
        # Add Speaker Notes & Timing
        notes_slide = slide.notes_slide
        notes_tf = notes_slide.notes_text_frame
        notes_tf.text = f"⏱️ Estimated Time: {s.get('time', '2 mins')}\n\n🎙️ Speaker Script:\n{s.get('speaker_notes', '')}"
        
    stream = BytesIO()
    prs.save(stream)
    stream.seek(0)
    return stream

# ================= SIDEBAR =================
with st.sidebar:
    st.markdown("## ⚡ Nexus Control Center")
    
    c_stat1, c_stat2 = st.columns(2)
    with c_stat1:
        if ai_client:
            st.success("🤖 AI: Online")
        else:
            st.error("🤖 AI: Offline")
    with c_stat2:
        if supabase_client:
            st.success("☁️ DB: Cloud")
        else:
            st.warning("💾 DB: Local")

    if not supabase_client:
        st.caption("💡 Add `SUPABASE_URL` & `SUPABASE_KEY` to save data permanently across restarts.")

    model_choice = st.selectbox("🧠 Gemini Model:", ["gemini-3.8-flash", "gemini-2.5-pro"], index=0)
    st.divider()
    
    # Smart Add / Update Form
    st.subheader("🔄 Add / Update Topic")
    topic_names = [t['title'] for t in st.session_state.topics]
    mode = st.radio("Action Mode:", ["➕ Create New", "✏️ Edit Existing"], horizontal=True)
    
    if mode == "✏️ Edit Existing" and topic_names:
        sel_edit = st.selectbox("Select Topic:", topic_names)
        curr_t = next(t for t in st.session_state.topics if t['title'] == sel_edit)
        in_title = curr_t['title']
        default_cat = curr_t.get('category', 'Project Management (PM)')
        default_stat = curr_t.get('status', 'In Progress')
        default_mast = int(curr_t.get('mastery', 50))
    else:
        in_title = st.text_input("Topic Title:", placeholder="e.g., 5G Rollout Schedule")
        default_cat = "Project Management (PM)"
        default_stat = "In Progress"
        default_mast = 50
        
    cat_options = ["Project Management (PM)", "Technical / Engineering", "AI & Automation", "Business & Strategy"]
    stat_options = ["Not Started", "In Progress", "Completed"]
    
    in_cat = st.selectbox("Category:", cat_options, index=cat_options.index(default_cat) if default_cat in cat_options else 0)
    in_stat = st.selectbox("Status:", stat_options, index=stat_options.index(default_stat) if default_stat in stat_options else 1)
    in_mast = st.slider("Mastery (%):", 0, 100, default_mast, step=5)
    
    if st.button("💾 Save / Update Topic", use_container_width=True, type="primary"):
        if in_title.strip():
            act = upsert_topic(in_title, in_cat, in_stat, in_mast)
            st.success(f"Topic '{in_title}' {act} successfully!")
            st.rerun()
        else:
            st.warning("Please enter a topic title.")

    st.divider()
    with st.expander("📦 Backup & Restore Data (JSON)"):
        backup_data = json.dumps({"topics": st.session_state.topics, "glossary": st.session_state.glossary}, indent=2)
        st.download_button("⬇️ Download Backup", data=backup_data, file_name="nexus_backup.json", mime="application/json", use_container_width=True)
        uploaded_backup = st.file_uploader("⬆️ Restore JSON:", type=['json'])
        if uploaded_backup and st.button("Restore Now", use_container_width=True):
            data = json.load(uploaded_backup)
            st.session_state.topics = data.get("topics", [])
            st.session_state.glossary = data.get("glossary", [])
            st.success("Backup Restored!")
            st.rerun()

# ================= MAIN HEADER =================
st.markdown("""
<div class="main-header">
    <h1>🧠 Nexus | Personal AI Second Brain & Pitch Engine</h1>
    <p>Track engineering & PM milestones, digest technical PDFs/Slides, generate executive decks, and master active recall.</p>
</div>
""", unsafe_allow_html=True)

tab1, tab2, tab3, tab4 = st.tabs([
    "📊 1. Progress Tracker", 
    "🧠 2. Knowledge Vault & OCR", 
    "🎯 3. AI Pitch Studio", 
    "🎙️ 4. AI Coach & Glossary"
])

# ================= TAB 1: TRACKER =================
with tab1:
    if st.session_state.topics:
        df = pd.DataFrame(st.session_state.topics)
        
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("📚 Total Topics", len(df))
        m2.metric("✅ Completed", len(df[df['status'] == 'Completed']))
        m3.metric("⚡ In Progress", len(df[df['status'] == 'In Progress']))
        m4.metric("🎯 Avg Mastery", f"{df['mastery'].astype(int).mean():.0f}%")
        
        st.markdown("### 📌 Interactive Knowledge Directory")
        f_col1, f_col2 = st.columns([2, 1])
        with f_col1:
            search_q = st.text_input("🔍 Search topics or notes...", placeholder="Type keyword...")
        with f_col2:
            filter_cat = st.multiselect("Filter by Category:", df['category'].unique(), default=df['category'].unique())
            
        filtered_df = df[df['category'].isin(filter_cat)]
        if search_q:
            filtered_df = filtered_df[filtered_df['title'].str.contains(search_q, case=False) | filtered_df['notes'].str.contains(search_q, case=False)]
            
        st.dataframe(
            filtered_df[['title', 'category', 'status', 'mastery', 'updated_at', 'notes']],
            column_config={
                "title": "Topic Title",
                "category": "Domain",
                "status": "Status",
                "mastery": st.column_config.ProgressColumn("Mastery Level", format="%d%%", min_value=0, max_value=100),
                "updated_at": "Last Updated",
                "notes": "Key Notes"
            },
            use_container_width=True,
            hide_index=True
        )
    else:
        st.info("No topics logged yet. Add your first topic from the left sidebar!")

# ================= TAB 2: KNOWLEDGE VAULT =================
with tab2:
    st.subheader("📥 Multi-Modal Knowledge Ingestion (PDFs, Slide Images & Notes)")
    topic_list = [t['title'] for t in st.session_state.topics]
    
    if not topic_list:
        st.warning("Please create a topic in the sidebar first.")
    else:
        active_t_name = st.selectbox("Select Topic to Enrich:", topic_list)
        active_t_obj = next(t for t in st.session_state.topics if t['title'] == active_t_name)
        
        col_v1, col_v2 = st.columns(2)
        with col_v1:
            user_notes = st.text_area("✍️ Your Personal Notes & Takeaways:", value=active_t_obj.get('notes', ''), height=200)
            if st.button("💾 Save Notes Only"):
                upsert_topic(active_t_name, active_t_obj['category'], active_t_obj['status'], active_t_obj['mastery'], notes=user_notes)
                st.success("Notes updated!")
                
        with col_v2:
            up_file = st.file_uploader("📄 Upload Lecture PDF or Slide Screenshot (PNG/JPG):", type=['pdf', 'png', 'jpg', 'jpeg'])
            extract_glossary = st.checkbox("Auto-extract technical/PM terms to Glossary", value=True)
            
            if st.button("🚀 Analyze with Gemini AI", type="primary", use_container_width=True):
                if not up_file and not user_notes:
                    st.warning("Upload a file or write notes to analyze.")
                else:
                    with st.spinner(f"Analyzing with {model_choice}..."):
                        extracted_text = ""
                        img_to_send = None
                        
                        if up_file:
                            if up_file.name.endswith('.pdf'):
                                reader = PyPDF2.PdfReader(up_file)
                                for page in reader.pages:
                                    extracted_text += (page.extract_text() or "") + "\n"
                            else:
                                img_to_send = up_file
                                
                        prompt = f"""
                        Analyze this material for the topic '{active_t_name}'.
                        User Notes: {user_notes}
                        Document Text: {extracted_text[:6000]}
                        
                        Provide:
                        1. Executive Summary (4 clear bullet points).
                        2. Core Practical / Engineering / PM Takeaways.
                        3. 3 Key Terms in format -> TERM: [name] | DEF: [1-sentence definition]
                        """
                        ai_res = ask_gemini(prompt, image_File=img_to_send, model_name=model_choice)
                        upsert_topic(active_t_name, active_t_obj['category'], active_t_obj['status'], active_t_obj['mastery'], notes=user_notes, ai_summary=ai_res)
                        
                        if extract_glossary and "TERM:" in ai_res:
                            for line in ai_res.split('\n'):
                                if "TERM:" in line and "| DEF:" in line:
                                    parts = line.split("| DEF:")
                                    t_name = parts[0].replace("TERM:", "").replace("*", "").strip()
                                    t_def = parts[1].strip()
                                    upsert_glossary(t_name, active_t_obj['category'], t_def)
                        st.success("Analysis saved & synced to topic!")
                        st.rerun()

        if active_t_obj.get('ai_summary'):
            st.divider()
            st.markdown("#### 🤖 Saved AI Synthesis & Summary")
            st.info(active_t_obj['ai_summary'])

# ================= TAB 3: PITCH STUDIO =================
with tab3:
    st.subheader("🎯 Automated Executive Presentation & Speaker Script Studio")
    p_col1, p_col2, p_col3 = st.columns(3)
    with p_col1:
        pitch_topic = st.selectbox("Select Topic:", topic_list if topic_list else ["General Topic"])
    with p_col2:
        audience = st.selectbox("Target Audience:", ["Executive Board / PMO (KPIs & ROI)", "Engineering Team (Architecture & Deep-Dive)", "Interview / Oral Assessment Panel"])
    with p_col3:
        slide_count = st.slider("Number of Slides:", 3, 8, 4)
        
    if st.button("✨ Generate Slide Deck + Speaker Notes", type="primary"):
        t_obj = next((t for t in st.session_state.topics if t['title'] == pitch_topic), {})
        context_data = f"Notes: {t_obj.get('notes', '')}\nSummary: {t_obj.get('ai_summary', '')}"
        
        with st.spinner("Architecting slides and speaker scripts..."):
            pitch_prompt = f"""
            Create a {slide_count}-slide presentation on '{pitch_topic}' for '{audience}'.
            Context: {context_data}
            Return ONLY a valid JSON array of objects. No markdown ticks, just pure JSON:
            [
              {{"title": "Slide Title", "bullets": ["Point 1", "Point 2", "Point 3"], "speaker_notes": "Exact English script to say...", "time": "2 mins"}}
            ]
            """
            raw_json = ask_gemini(pitch_prompt, model_name=model_choice)
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
            except Exception as e:
                st.warning("Displaying raw outline (JSON parse fallback):")
                st.write(raw_json)

# ================= TAB 4: AI COACH & GLOSSARY =================
with tab4:
    c_left, c_right = st.columns([3, 2])
    
    with c_left:
        st.subheader("🎙️ Active Recall & Oral Assessment Coach")
        coach_mode = st.radio("Select Practice Mode:", ["🧠 Feynman Explanation Evaluator", "📝 Generate Interactive Quiz"], horizontal=True)
        coach_topic = st.selectbox("Topic to Practice:", topic_list if topic_list else ["General"])
        
        if coach_mode == "🧠 Feynman Explanation Evaluator":
            user_exp = st.text_area("Explain the concept in English (as if in an interview or PMO meeting):", height=140)
            if st.button("🔍 Evaluate My Delivery", type="primary"):
                if user_exp:
                    with st.spinner("Evaluating technical accuracy & English delivery..."):
                        eval_prompt = f"""
                        Act as a senior Engineering & PMO interviewer. Evaluate this candidate's explanation of '{coach_topic}':
                        "{user_exp}"
                        Provide:
                        1. Score out of 10 (Clarity & Technical Accuracy).
                        2. Missing industry keywords/jargon they should have used.
                        3. Polished, executive-level version of their answer (3-4 sentences).
                        """
                        st.markdown(ask_gemini(eval_prompt, model_name=model_choice))
        else:
            if st.button("🎲 Generate 3 Scenario Questions", type="primary"):
                with st.spinner("Creating quiz..."):
                    q_prompt = f"Create 3 practical multiple-choice scenario questions about '{coach_topic}' with answers and brief explanations at the bottom."
                    st.markdown(ask_gemini(q_prompt, model_name=model_choice))

    with c_right:
        st.subheader("📖 Smart Terminology Glossary")
        with st.expander("➕ Manually Add Term"):
            g_t = st.text_input("Term:")
            g_c = st.selectbox("Domain:", ["PM", "Technical", "Business"])
            g_d = st.text_area("Definition / Usage in a sentence:")
            if st.button("Save Term"):
                if g_t:
                    upsert_glossary(g_t, g_c, g_d)
                    st.success("Term added!")
                    st.rerun()
                    
        if st.session_state.glossary:
            st.dataframe(pd.DataFrame(st.session_state.glossary), use_container_width=True, hide_index=True)
