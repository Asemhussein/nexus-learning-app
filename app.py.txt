import streamlit as st
from supabase import create_client, Client
import pandas as pd
from datetime import datetime

st.set_page_config(page_title="Nexus | Cloud Engine", layout="wide")

# --- 1. Connect to Supabase Cloud ---
@st.cache_resource
def init_connection() -> Client:
    url = st.secrets["SUPABASE_URL"]
    key = st.secrets["SUPABASE_KEY"]
    return create_client(url, key)

supabase = init_connection()

# --- 2. Cloud Database Functions ---
def fetch_topics():
    res = supabase.table("topics").select("*").order("id", desc=True).execute()
    return res.data

def insert_topic(title, category, status, mastery, notes):
    supabase.table("topics").insert({
        "title": title,
        "category": category,
        "status": status,
        "mastery": mastery,
        "notes": notes,
        "updated_at": datetime.now().strftime("%Y-%m-%d")
    }).execute()

# --- Sidebar: Add New Topic from Phone or Laptop ---
with st.sidebar:
    st.title("⚡ Nexus Cloud")
    st.subheader("➕ Log New Topic")
    t_title = st.text_input("Topic Name:")
    t_cat = st.selectbox("Category:", ["Project Management (PM)", "Technical / Engineering", "AI & Automation"])
    t_status = st.selectbox("Status:", ["Not Started", "In Progress", "Completed"])
    t_mastery = st.select_slider("Mastery Level:", options=["0%", "25%", "50%", "75%", "100%"], value="50%")
    t_notes = st.text_area("Key Takeaways / Summary:")

    if st.button("Save to Cloud ☁️", use_container_width=True, type="primary"):
        if t_title:
            insert_topic(t_title, t_cat, t_status, t_mastery, t_notes)
            st.success("Synced to Cloud!")
            st.rerun()

# --- Main Dashboard ---
st.header("📊 Cloud-Synced Learning Tracker")
topics_data = fetch_topics()

if topics_data:
    df = pd.DataFrame(topics_data)
    st.dataframe(df[["title", "category", "status", "mastery", "updated_at", "notes"]], use_container_width=True)
else:
    st.info("No topics logged yet. Add your first topic from the sidebar!")