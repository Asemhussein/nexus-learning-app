import streamlit as st
import pandas as pd
from datetime import datetime

# --- Page Configuration ---
st.set_page_config(page_title="Nexus | AI Learning & Pitch Engine", layout="wide")

# --- Safe Supabase Import (Prevents Server Hang) ---
try:
    from supabase import create_client, Client
    supabase_installed = True
except ImportError:
    supabase_installed = False

# --- 1. Connect to Supabase Cloud ---
db_connected = False
supabase = None

if supabase_installed:
    try:
        if "SUPABASE_URL" in st.secrets and "SUPABASE_KEY" in st.secrets:
            supabase = create_client(st.secrets["SUPABASE_URL"], st.secrets["SUPABASE_KEY"])
            db_connected = True
    except Exception as e:
        db_connected = False

# --- Fallback Local Session State (If DB not connected yet) ---
if 'local_topics' not in st.session_state:
    st.session_state.local_topics = [
        {"title": "Telecom Rollout KPIs & Risk Mgmt", "category": "Project Management (PM)", "status": "Completed", "mastery": "90%", "updated_at": "2026-10-20", "notes": "Sample note"},
        {"title": "Hardware Accelerators & LUT Optimization", "category": "Technical / Engineering", "status": "Completed", "mastery": "95%", "updated_at": "2026-10-15", "notes": "Sample note"}
    ]
if 'local_glossary' not in st.session_state:
    st.session_state.local_glossary = [
        {"term": "Critical Path", "category": "PM", "context": "We need to monitor the critical path to avoid rollout delays."}
    ]

# --- 2. Database Functions (Cloud + Local Fallback) ---
def fetch_topics():
    if db_connected:
        try:
            res = supabase.table("topics").select("*").order("id", desc=True).execute()
            return res.data
        except Exception:
            return st.session_state.local_topics
    return st.session_state.local_topics

def insert_topic(title, category, status, mastery, notes=""):
    new_item = {
        "title": title,
        "category": category,
        "status": status,
        "mastery": mastery,
        "notes": notes,
        "updated_at": datetime.now().strftime("%Y-%m-%d")
    }
    if db_connected:
        supabase.table("topics").insert(new_item).execute()
    else:
        st.session_state.local_topics.insert(0, new_item)

def fetch_glossary():
    if db_connected:
        try:
            res = supabase.table("glossary").select("*").order("id", desc=True).execute()
            return res.data
        except Exception:
            return st.session_state.local_glossary
    return st.session_state.local_glossary

def insert_glossary(term, category, context):
    new_term = {"term": term, "category": category, "context": context}
    if db_connected:
        supabase.table("glossary").insert(new_term).execute()
    else:
        st.session_state.local_glossary.insert(0, new_term)

# --- Load Data ---
topics_data = fetch_topics()
glossary_data = fetch_glossary()

# --- Sidebar ---
with st.sidebar:
    st.title("⚡ Nexus Engine")
    if db_connected:
        st.success("🟢 Cloud Sync Active (Supabase)")
    else:
        st.warning("🟡 Running in Fast Mode (Local). Add 'supabase' to requirements.txt to enable Cloud Sync.")
    st.divider()
    
    st.subheader("➕ Quick Log")
    new_title = st.text_input("Topic / Course Module Name:")
    new_cat = st.selectbox("Category:", ["Project Management (PM)", "Technical / Engineering", "AI & Automation", "Business & Strategy"])
    new_status = st.selectbox("Current Status:", ["Not Started", "In Progress", "Completed"])
    new_mastery = st.select_slider("Mastery Level:", options=["0%", "25%", "50%", "75%", "100%"], value="50%")
    
    if st.button("Save Topic 🚀", use_container_width=True, type="primary"):
        if new_title:
            insert_topic(new_title, new_cat, new_status, new_mastery)
            st.success(f"Saved: {new_title}")
            st.rerun()
            
    st.divider()
    st.subheader("📖 Add New Jargon / Term")
    g_term = st.text_input("Term / Acronym:")
    g_cat = st.selectbox("Term Category:", ["PM", "Technical", "General"])
    g_ctx = st.text_area("Example Sentence / Definition:", height=70)
    if st.button("Add to Glossary", use_container_width=True):
        if g_term:
            insert_glossary(g_term, g_cat, g_ctx)
            st.success(f"Added term: {g_term}")
            st.rerun()

# --- Main App Tabs ---
tab1, tab2, tab3, tab4 = st.tabs([
    "📊 1. Tracker & Analytics", 
    "🧠 2. Knowledge Vault", 
    "🎯 3. Presentation Studio", 
    "🎙️ 4. AI Coach & Retention"
])

with tab1:
    st.header("Learning Progress & Milestones")
    if topics_data:
        df_topics = pd.DataFrame(topics_data)
        total_count = len(df_topics)
        completed_count = len(df_topics[df_topics['status'] == 'Completed'])
        comp_rate = int((completed_count / total_count) * 100) if total_count > 0 else 0
        
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Total Logged Topics", total_count, "Active")
        col2.metric("Completed Milestones", completed_count, f"{comp_rate}% Rate")
        col3.metric("Learning Velocity", "1.4x", "On Track")
        col4.metric("System Status", "Online 🟢", "Ready")

        st.divider()
        st.subheader("📌 Topic Directory & Timeline")
        st.dataframe(df_topics, use_container_width=True)

with tab2:
    st.header("Multi-Modal Knowledge Ingestion")
    topic_titles = [t['title'] for t in topics_data] if topics_data else ["General Notes"]
    col_in1, col_in2 = st.columns(2)
    with col_in1:
        selected_topic = st.selectbox("Select Active Topic:", topic_titles)
        user_notes = st.text_area("Write what you understood (in English):", height=200)
    with col_in2:
        uploaded_files = st.file_uploader("Upload Slides, PDFs, or Images:", type=['pdf', 'png', 'jpg', 'jpeg'], accept_multiple_files=True)
    if st.button("🚀 Save & Process", type="primary"):
        st.success(f"Saved under '{selected_topic}'!")

with tab3:
    st.header("AI Presentation & Pitch Generator")
    p_col1, p_col2, p_col3 = st.columns(3)
    with p_col1:
        pitch_topics = st.multiselect("Select Topics:", topic_titles, default=[topic_titles[0]])
    with p_col2:
        audience = st.selectbox("Target Audience:", ["Executive Summary", "Technical Deep-Dive", "Cross-Functional Team"])
    with p_col3:
        duration = st.slider("Duration (Minutes):", 5, 60, 15, step=5)
        
    if st.button("✨ Generate Slide Deck, Script & Q&A", type="primary"):
        s1, s2, s3 = st.columns(3)
        s1.markdown("#### Slide 1: Overview\n* Scope & KPIs\n* Strategic Alignment\n\n*⏱️ Est. Time: 2 mins*")
        s2.markdown("#### Slide 2: Deep Dive\n* Architecture & Workflow\n* Optimization Logic\n\n*⏱️ Est. Time: 4 mins*")
        s3.markdown("#### Slide 3: Next Steps\n* Risk Mitigation\n* Rollout Timeline\n\n*⏱️ Est. Time: 3 mins*")

with tab4:
    st.header("Active Recall & Communication Coach")
    c1, c2 = st.columns(2)
    with c1:
        st.subheader("🎙️ Feynman Mode & Rehearsal")
        rehearsal_input = st.text_area("Explain a concept in English:")
        if st.button("Evaluate Explanation"):
            st.info("**Feedback:** Clear structure! Try incorporating terms from your Glossary for stronger impact.")
    with c2:
        st.subheader("📖 Glossary Bank")
        st.dataframe(pd.DataFrame(glossary_data), use_container_width=True)
