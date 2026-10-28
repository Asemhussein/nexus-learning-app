import streamlit as st
from supabase import create_client, Client
import pandas as pd
from datetime import datetime

# --- Page Configuration ---
st.set_page_config(page_title="Nexus | AI Learning & Pitch Engine", layout="wide")

# --- 1. Connect to Supabase Cloud ---
@st.cache_resource
def init_connection() -> Client:
    url = st.secrets["SUPABASE_URL"]
    key = st.secrets["SUPABASE_KEY"]
    return create_client(url, key)

try:
    supabase = init_connection()
    db_connected = True
except Exception as e:
    db_connected = False
    st.error("⚠️ Database not connected yet. Please add SUPABASE_URL and SUPABASE_KEY in Streamlit Secrets.")

# --- 2. Cloud Database Functions ---
def fetch_topics():
    if not db_connected:
        return []
    res = supabase.table("topics").select("*").order("id", desc=True).execute()
    return res.data

def insert_topic(title, category, status, mastery, notes=""):
    if db_connected:
        supabase.table("topics").insert({
            "title": title,
            "category": category,
            "status": status,
            "mastery": mastery,
            "notes": notes,
            "updated_at": datetime.now().strftime("%Y-%m-%d")
        }).execute()

def update_topic_notes(topic_title, new_notes):
    if db_connected:
        supabase.table("topics").update({
            "notes": new_notes,
            "updated_at": datetime.now().strftime("%Y-%m-%d")
        }).eq("title", topic_title).execute()

def fetch_glossary():
    if not db_connected:
        return []
    res = supabase.table("glossary").select("*").order("id", desc=True).execute()
    return res.data

def insert_glossary(term, category, context):
    if db_connected:
        supabase.table("glossary").insert({
            "term": term,
            "category": category,
            "context": context
        }).execute()

# --- Load Data from Cloud ---
topics_data = fetch_topics()
glossary_data = fetch_glossary()

# --- Sidebar: Quick Actions & Cloud Sync ---
with st.sidebar:
    st.title("⚡ Nexus Engine")
    st.caption("Cloud-Synced Learning & Pitch Studio")
    st.divider()
    
    st.subheader("➕ Quick Log to Cloud")
    new_title = st.text_input("Topic / Course Module Name:")
    new_cat = st.selectbox("Category:", ["Project Management (PM)", "Technical / Engineering", "AI & Automation", "Business & Strategy"])
    new_status = st.selectbox("Current Status:", ["Not Started", "In Progress", "Completed"])
    new_mastery = st.select_slider("Mastery Level:", options=["0%", "25%", "50%", "75%", "100%"], value="50%")
    
    if st.button("Save Topic to Cloud ☁️", use_container_width=True, type="primary"):
        if new_title:
            insert_topic(new_title, new_cat, new_status, new_mastery)
            st.success(f"Synced: {new_title}")
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
    "🧠 2. Knowledge Vault (Input/OCR)", 
    "🎯 3. Presentation Studio", 
    "🎙️ 4. AI Coach & Retention"
])

# ==========================================
# TAB 1: TRACKER, MILESTONES & ANALYTICS
# ==========================================
with tab1:
    st.header("Learning Progress & Milestones")
    
    if topics_data:
        df_topics = pd.DataFrame(topics_data)
        total_count = len(df_topics)
        completed_count = len(df_topics[df_topics['status'] == 'Completed'])
        comp_rate = int((completed_count / total_count) * 100) if total_count > 0 else 0
        
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Total Logged Topics", total_count, "Cloud Synced")
        col2.metric("Completed Milestones", completed_count, f"{comp_rate}% Completion Rate")
        col3.metric("Learning Velocity", "Active", "Tracking across categories")
        col4.metric("Sync Status", "Online 🟢", "Multi-device Ready")

        st.divider()
        
        col_filter, col_digest = st.columns([2, 1])
        with col_filter:
            st.subheader("📌 Topic Directory & Timeline")
            available_cats = df_topics['category'].unique().tolist()
            cat_filter = st.multiselect("Filter by Category:", options=available_cats, default=available_cats)
            
            display_cols = ["title", "category", "status", "mastery", "updated_at", "notes"]
            existing_cols = [c for c in display_cols if c in df_topics.columns]
            st.dataframe(df_topics[df_topics['category'].isin(cat_filter)][existing_cols], use_container_width=True)
            
        with col_digest:
            st.subheader("📋 Weekly Executive Digest")
            st.info(f"""
            **Current Cloud Snapshot:**
            * **Total Topics Tracked:** {total_count}
            * **Completed:** {completed_count}
            * **Latest Topic:** {df_topics['title'].iloc[0]}
            * **Action Item:** Review notes in the Knowledge Vault and generate a summary deck.
            """)
    else:
        st.info("👋 Your cloud database is currently empty! Add your first topic from the left sidebar to start tracking.")

# ==========================================
# TAB 2: MULTI-MODAL KNOWLEDGE VAULT
# ==========================================
with tab2:
    st.header("Multi-Modal Knowledge Ingestion")
    
    topic_titles = [t['title'] for t in topics_data] if topics_data else ["No topics available - Add one in sidebar"]
    
    col_in1, col_in2 = st.columns(2)
    with col_in1:
        st.subheader("📝 My Understanding & Notes")
        selected_topic = st.selectbox("Select Active Topic:", topic_titles)
        
        # Load existing notes for the selected topic
        current_notes = ""
        if topics_data:
            for t in topics_data:
                if t['title'] == selected_topic and t.get('notes'):
                    current_notes = t['notes']
                    
        user_notes = st.text_area("Write what you understood (in English):", value=current_notes, height=200, 
                                  placeholder="Explain the core concepts, formulas, or PM frameworks here...")
        
    with col_in2:
        st.subheader("📂 Upload Course Material (PDF / Images / Syllabus)")
        uploaded_files = st.file_uploader("Upload Slides, PDFs, or Whiteboard Photos (OCR Ready):", 
                                          type=['pdf', 'png', 'jpg', 'jpeg'], accept_multiple_files=True)
        parse_mode = st.radio("AI Processing Action:", [
            "Summarize & Append to Topic Notes",
            "Syllabus Auto-Parser (Convert Outline to Tracker Milestones)",
            "Extract Technical/PM Jargon to Glossary"
        ])
        
    if st.button("🚀 Save Notes & Process Files to Cloud", type="primary"):
        if topics_data:
            update_topic_notes(selected_topic, user_notes)
            st.success(f"Notes saved to cloud under '{selected_topic}'!")
            st.rerun()
        else:
            st.warning("Please add a topic from the sidebar first.")

# ==========================================
# TAB 3: PRESENTATION & PITCH STUDIO
# ==========================================
with tab3:
    st.header("AI Presentation & Pitch Generator")
    
    p_col1, p_col2, p_col3 = st.columns(3)
    with p_col1:
        pitch_topics = st.multiselect("Select Topics to Combine in Presentation:", topic_titles, default=[topic_titles[0]] if topics_data else [])
    with p_col2:
        audience = st.selectbox("Target Audience:", [
            "Executive / Management (KPIs, ROI, High-Level)",
            "Technical Deep-Dive (Architecture, Specs, Logic)",
            "Cross-Functional Team (Workflows, Milestones, Action Items)"
        ])
    with p_col3:
        duration = st.slider("Target Presentation Duration (Minutes):", 5, 60, 15, step=5)
        
    include_diagrams = st.checkbox("Generate Process Flowcharts / Block Diagrams (Mermaid.js)", value=True)
    
    if st.button("✨ Generate Slide Deck, Script & Q&A", type="primary"):
        st.subheader("📑 Generated Slide Deck Preview")
        
        slide1, slide2, slide3 = st.columns(3)
        with slide1:
            st.markdown("#### Slide 1: Executive Overview\n* Core Objective & Scope\n* Key Milestones Achieved\n* Strategic Alignment")
            st.caption("⏱️ Est. Speaking Time: 2 mins")
            with st.expander("🎙️ Speaker Notes (English)"):
                st.write('"Good morning everyone. Today I will walk you through our progress on..."')
                
        with slide2:
            st.markdown("#### Slide 2: Architecture & Workflow\n* System Bottlenecks Identified\n* Optimization Strategy\n* Resource Allocation")
            st.caption("⏱️ Est. Speaking Time: 4 mins")
            with st.expander("🎙️ Speaker Notes (English)"):
                st.write('"Moving to the technical workflow, you can see how the components interact..."')

        with slide3:
            st.markdown("#### Slide 3: Impact & Next Steps\n* Efficiency Gains & KPIs\n* Risk Mitigation Plan\n* Rollout Timeline")
            st.caption("⏱️ Est. Speaking Time: 3 mins")
            with st.expander("🎙️ Speaker Notes (English)"):
                st.write('"To wrap up, our immediate next steps focus on scaling this across..."')
        
        st.divider()
        st.download_button("📥 Export Ready-to-Present (.PPTX)", data="Mock PPTX Binary", file_name="Nexus_Presentation.pptx")

# ==========================================
# TAB 4: AI COACH, FEYNMAN & RETENTION LAB
# ==========================================
with tab4:
    st.header("Active Recall & Communication Coach")
    
    coach_col1, coach_col2 = st.columns(2)
    
    with coach_col1:
        st.subheader("🎙️ Feynman Mode & Rehearsal")
        st.write("Explain a concept in text. AI will evaluate your accuracy, English fluency, and highlight knowledge gaps.")
        rehearsal_input = st.text_area("Type your explanation:", placeholder="So basically, this concept works by...")
        if st.button("Evaluate My Explanation"):
            st.warning("**Knowledge Gap Check:** Make sure to connect the technical execution with the project timeline.")
            st.info("**Language Polish:** Use structured transitions like *'Consequently'*, *'From an architectural standpoint'*, and *'To mitigate this risk'*.")
            
        st.divider()
        st.subheader("🔥 Tough Q&A Simulator")
        if st.button("Generate 3 Challenging Assessment Questions"):
            st.markdown("""
            1. **Scenario Question:** *If a critical dependency fails mid-rollout, how does your proposed framework mitigate downtime?*
            2. **Technical Probe:** *What is the trade-off between accuracy and resource utilization in this model?*
            3. **PM Alignment:** *How do you justify the time invested in this optimization to non-technical stakeholders?*
            """)
            
    with coach_col2:
        st.subheader("📖 Cloud Jargon & Glossary Bank")
        if glossary_data:
            df_glossary = pd.DataFrame(glossary_data)
            st.dataframe(df_glossary[["term", "category", "context"]], use_container_width=True)
        else:
            st.info("No terms added yet. Add new terminology from the left sidebar!")
        
        st.subheader("⚡ Spaced Repetition Flashcards")
        with st.container(border=True):
            st.markdown("**Flashcard Practice**")
            if glossary_data:
                sample_term = glossary_data[0]
                st.markdown(f"❓ *How do you define or use the term:* **{sample_term['term']}** *(Category: {sample_term['category']})*?")
                if st.button("Reveal Context / Answer"):
                    st.success(sample_term['context'])
            else:
                st.markdown("❓ *What is the difference between a Milestone and a Deliverable?*")
                if st.button("Reveal Answer"):
                    st.success("A **Deliverable** is a tangible/measurable output, while a **Milestone** is a significant checkpoint in the project timeline.")
