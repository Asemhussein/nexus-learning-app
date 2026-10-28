import streamlit as st
import pandas as pd
from datetime import datetime
import os
from io import BytesIO
import PyPDF2
from pptx import Presentation
from google import genai

# --- Page Configuration ---
st.set_page_config(page_title="Nexus | AI Second Brain", page_icon="🧠", layout="wide")

# --- Initialize Session State (Local DB until Supabase is linked) ---
if 'topics' not in st.session_state:
    st.session_state.topics = []
if 'glossary' not in st.session_state:
    st.session_state.glossary = []
if 'vault_notes' not in st.session_state:
    st.session_state.vault_notes = {}

# --- AI Engine Setup ---
api_key = os.environ.get("GEMINI_API_KEY")
client = genai.Client(api_key=api_key) if api_key else None

def ask_gemini(prompt):
    if not client:
        return "⚠️ AI Engine offline. Please set GEMINI_API_KEY."
    try:
        response = client.models.generate_content(model='gemini-2.5-flash', contents=prompt)
        return response.text
    except Exception as e:
        return f"Error: {e}"

# --- Helper: Extract PDF Text ---
def extract_text_from_pdf(uploaded_file):
    reader = PyPDF2.PdfReader(uploaded_file)
    text = ""
    for page in reader.pages:
        text += page.extract_text() + "\n"
    return text

# --- Helper: Generate PPTX with AI Content ---
def create_pptx(topic, audience, ai_content):
    prs = Presentation()
    
    # Title Slide
    title_slide = prs.slides.add_slide(prs.slide_layouts[0])
    title_slide.shapes.title.text = topic
    title_slide.placeholders[1].text = f"Tailored for: {audience}"

    # Parse AI Content (Expected simple format: Slide Title \n - point 1...)
    slides_data = ai_content.split("Slide:")
    for slide_data in slides_data[1:]:
        lines = slide_data.strip().split('\n')
        slide_title = lines[0].strip()
        bullets = [line.replace('-', '').strip() for line in lines[1:] if line.strip().startswith('-')]
        
        bullet_slide = prs.slides.add_slide(prs.slide_layouts[1])
        bullet_slide.shapes.title.text = slide_title
        tf = bullet_slide.placeholders[1].text_frame
        for i, bullet in enumerate(bullets):
            p = tf.add_paragraph()
            p.text = bullet
            if i > 0: p.level = 1

    ppt_stream = BytesIO()
    prs.save(ppt_stream)
    ppt_stream.seek(0)
    return ppt_stream

# --- Sidebar: Quick Actions ---
with st.sidebar:
    st.title("⚡ Nexus Engine")
    if api_key:
        st.success("🟢 AI Engine: Online")
    else:
        st.error("🔴 AI Engine: Offline")
    
    st.divider()
    st.subheader("➕ Log New Topic")
    new_title = st.text_input("Topic Name:")
    new_cat = st.selectbox("Category:", ["Project Management (PM)", "Technical / Engineering", "General"])
    new_status = st.selectbox("Status:", ["Not Started", "In Progress", "Completed"])
    new_mastery = st.slider("Mastery Level (%)", 0, 100, 10)
    
    if st.button("Log Topic 🚀", use_container_width=True):
        if new_title:
            st.session_state.topics.append({
                "Title": new_title, "Category": new_cat, 
                "Status": new_status, "Mastery": f"{new_mastery}%", 
                "Date": datetime.now().strftime("%Y-%m-%d")
            })
            st.success("Topic Saved!")

# --- Main App Tabs ---
tab1, tab2, tab3, tab4 = st.tabs([
    "📊 1. Tracker", "🧠 2. Knowledge Vault", "🎯 3. Pitch Studio", "🎙️ 4. AI Coach"
])

with tab1:
    st.header("Learning Progress & Milestones")
    if st.session_state.topics:
        df = pd.DataFrame(st.session_state.topics)
        
        c1, c2, c3 = st.columns(3)
        c1.metric("Total Topics", len(df))
        c2.metric("Completed", len(df[df['Status'] == 'Completed']))
        c3.metric("Avg Mastery", f"{df['Mastery'].str.replace('%','').astype(int).mean():.1f}%")
        
        st.dataframe(df, use_container_width=True)
    else:
        st.info("No topics logged yet. Use the sidebar to add your first topic!")

with tab2:
    st.header("Multi-Modal Knowledge Ingestion")
    topic_list = [t["Title"] for t in st.session_state.topics] if st.session_state.topics else ["General"]
    active_topic = st.selectbox("Assign material to topic:", topic_list)
    
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("Upload Material")
        uploaded_file = st.file_uploader("Upload PDF Course Material:", type=['pdf'])
    with col2:
        st.subheader("Your Notes")
        user_notes = st.text_area("Write key takeaways:", height=150)
        
    if st.button("Analyze & Summarize Material", type="primary"):
        if uploaded_file and api_key:
            with st.spinner("AI is reading the document..."):
                raw_text = extract_text_from_pdf(uploaded_file)
                # Limit text size to avoid token limits for this demo
                summary_prompt = f"Summarize the following technical/PM material in 4 bullet points. Material: {raw_text[:3000]}"
                ai_summary = ask_gemini(summary_prompt)
                
                st.session_state.vault_notes[active_topic] = ai_summary
                st.success("Analysis Complete!")
                st.info(f"**AI Summary:**\n{ai_summary}")
        elif not api_key:
            st.warning("Please configure Gemini API Key.")

with tab3:
    st.header("AI Presentation Generator")
    st.markdown("Transform your notes into structured slide decks.")
    
    pitch_topic = st.selectbox("Select Topic to Pitch:", topic_list, key="pitch")
    audience = st.selectbox("Target Audience:", ["Executives (High-level, KPIs)", "Technical Team (Deep-dive, Architecture)", "Clients (Benefits, ROI)"])
    
    if st.button("✨ Generate AI Slide Deck", type="primary"):
        if api_key:
            with st.spinner("AI is architecting your presentation..."):
                prompt = f"""
                Create a 3-slide presentation outline for the topic '{pitch_topic}'.
                The audience is '{audience}'.
                Format STRICTLY as follows:
                Slide: Introduction
                - Point 1
                - Point 2
                Slide: Main Concept
                - Point 1
                - Point 2
                Slide: Conclusion
                - Point 1
                - Point 2
                """
                ai_deck_content = ask_gemini(prompt)
                st.markdown("### Preview of AI Generated Deck:")
                st.write(ai_deck_content)
                
                # Generate PPTX
                ppt_file = create_pptx(pitch_topic, audience, ai_deck_content)
                st.download_button(
                    label="📥 Download Presentation (.pptx)",
                    data=ppt_file,
                    file_name=f"{pitch_topic.replace(' ', '_')}.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )
        else:
            st.warning("AI Engine offline.")

with tab4:
    st.header("Active Recall & Feynman Coach")
    st.markdown("Explain a concept as if teaching a beginner. The AI will evaluate your gaps.")
    
    rehearsal = st.text_area("Your Explanation (English):", placeholder="e.g., The critical path in project management is...")
    
    if st.button("Evaluate with AI"):
        if api_key and rehearsal:
            with st.spinner("AI Coach is analyzing..."):
                feedback_prompt = f"""
                You are a strict but helpful technical and project management coach. 
                Evaluate this explanation: "{rehearsal}"
                1. Give a score out of 10.
                2. Point out any missing technical or PM jargon.
                3. Provide a 1-sentence improved version.
                """
                feedback = ask_gemini(feedback_prompt)
                st.markdown("### Coach Feedback:")
                st.write(feedback)
        elif not rehearsal:
            st.warning("Please write an explanation first.")
