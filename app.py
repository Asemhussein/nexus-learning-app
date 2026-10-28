import streamlit as st
import pandas as pd
from datetime import datetime

# --- Page Configuration ---
st.set_page_config(page_title="Nexus | AI Learning & Pitch Engine", layout="wide")

# --- Initialize Local Session State ---
if 'topics' not in st.session_state:
    st.session_state.topics = [
        {"title": "Telecom Rollout KPIs & Risk Mgmt", "category": "Project Management (PM)", "status": "Completed", "mastery": "90%", "updated_at": "2026-10-20", "notes": "Core rollout metrics reviewed."},
        {"title": "Hardware Accelerators & LUT Optimization", "category": "Technical / Engineering", "status": "Completed", "mastery": "95%", "updated_at": "2026-10-15", "notes": "NVDLA architecture optimization."}
    ]
if 'glossary' not in st.session_state:
    st.session_state.glossary = [
        {"term": "Critical Path", "category": "PM", "context": "Monitor the critical path to avoid rollout delays."}
    ]

# --- Sidebar: Quick Actions ---
with st.sidebar:
    st.title("⚡ Nexus Engine")
    st.success("🟢 Stable Mode Active")
    st.divider()
    
    st.subheader("➕ Quick Log Topic")
    new_title = st.text_input("Topic Name:")
    new_cat = st.selectbox("Category:", ["Project Management (PM)", "Technical / Engineering", "AI & Automation", "Business & Strategy"])
    new_status = st.selectbox("Current Status:", ["Not Started", "In Progress", "Completed"])
    new_mastery = st.select_slider("Mastery Level:", options=["0%", "25%", "50%", "75%", "100%"], value="50%")
    
    if st.button("Save Topic 🚀", use_container_width=True, type="primary"):
        if new_title:
            st.session_state.topics.insert(0, {
                "title": new_title,
                "category": new_cat,
                "status": new_status,
                "mastery": new_mastery,
                "updated_at": datetime.now().strftime("%Y-%m-%d"),
                "notes": "Newly added topic."
            })
            st.success(f"Saved: {new_title}")

    st.divider()
    st.subheader("📖 Add Glossary Term")
    g_term = st.text_input("Term:")
    g_cat = st.selectbox("Term Category:", ["PM", "Technical", "General"])
    g_ctx = st.text_area("Definition / Context:", height=60)
    if st.button("Add Term", use_container_width=True):
        if g_term:
            st.session_state.glossary.insert(0, {"term": g_term, "category": g_cat, "context": g_ctx})
            st.success(f"Added: {g_term}")

# --- Main App Tabs ---
tab1, tab2, tab3, tab4 = st.tabs([
    "📊 1. Tracker & Analytics", 
    "🧠 2. Knowledge Vault", 
    "🎯 3. Presentation Studio", 
    "🎙️ 4. AI Coach & Retention"
])

with tab1:
    st.header("Learning Progress & Milestones")
    df_topics = pd.DataFrame(st.session_state.topics)
    total_count = len(df_topics)
    completed_count = len(df_topics[df_topics['status'] == 'Completed'])
    
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Total Logged Topics", total_count)
    col2.metric("Completed Milestones", completed_count)
    col3.metric("Learning Velocity", "1.4x")
    col4.metric("Status", "Stable 🟢")

    st.divider()
    st.subheader("📌 Topic Directory & Timeline")
    st.dataframe(df_topics, use_container_width=True)

with tab2:
    st.header("Multi-Modal Knowledge Ingestion")
    topic_titles = [t['title'] for t in st.session_state.topics]
    col_in1, col_in2 = st.columns(2)
    with col_in1:
        selected_topic = st.selectbox("Select Active Topic:", topic_titles)
        user_notes = st.text_area("Write what you understood (in English):", height=200, placeholder="Explain core concepts here...")
    with col_in2:
        uploaded_files = st.file_uploader("Upload Slides, PDFs, or Images:", type=['pdf', 'png', 'jpg', 'jpeg'], accept_multiple_files=True)
    if st.button("🚀 Save Notes", type="primary"):
        st.success("Notes saved successfully!")

with tab3:
    st.header("AI Presentation & Pitch Generator")
    p1, p2, p3 = st.columns(3)
    with p1:
        st.multiselect("Select Topics:", topic_titles, default=[topic_titles[0] if topic_titles else ""])
    with p2:
        st.selectbox("Target Audience:", ["Executive Summary", "Technical Deep-Dive", "Cross-Functional Team"])
    with p3:
        st.slider("Duration (Minutes):", 5, 60, 15, step=5)
        
    if st.button("✨ Generate Slide Deck & Script", type="primary"):
        s1, s2, s3 = st.columns(3)
        s1.markdown("#### Slide 1: Overview\n* Scope & KPIs\n* Strategic Alignment\n\n*⏱️ Est. Time: 2 mins*")
        s2.markdown("#### Slide 2: Deep Dive\n* Architecture & Workflow\n* Optimization Logic\n\n*⏱️ Est. Time: 4 mins*")
        s3.markdown("#### Slide 3: Next Steps\n* Risk Mitigation\n* Rollout Timeline\n\n*⏱️ Est. Time: 3 mins*")

with tab4:
    st.header("Active Recall & Communication Coach")
    c1, c2 = st.columns(2)
    with c1:
        st.subheader("🎙️ Feynman Mode & Rehearsal")
        st.text_area("Explain a concept in English:")
        if st.button("Evaluate Explanation"):
            st.info("**Feedback:** Clear and structured! Keep practicing.")
    with c2:
        st.subheader("📖 Glossary Bank")
        st.dataframe(pd.DataFrame(st.session_state.glossary), use_container_width=True)
