import streamlit as st

from agent import Agent
from llm_client import LLMClient
from conversation_context import ConversationContext
from tools.tools import tools

st.set_page_config(
    page_title="Software Architect Agent",
    page_icon="🏗️",
    layout="wide"
)

st.title("🏗️ Software Architect Agent")

# Inițializare
if "agent" not in st.session_state:
    context = ConversationContext()
    llm_client = LLMClient()
    st.session_state.agent = Agent(llm_client, context, tools=tools)
    st.session_state.chat_history = []

# Mesaj afișat la prima deschidere
if len(st.session_state.chat_history) == 0:
    st.info("""
    👋 Bine ai venit!

    Sunt Software Architect Agent și te pot ajuta cu:
    - Design arhitectural
    - Microservicii
    - Documentație tehnică
    - UML și diagrame
    - Code review
    - Best practices

    Încearcă:
    - "Proiectează o arhitectură pentru un magazin online"
    - "Generează o diagramă UML pentru sistemul de autentificare"
    - "Ce microservicii sunt necesare pentru o aplicație bancară?"
    """)

# Afișare istoric
for role, content in st.session_state.chat_history:
    with st.chat_message(role):
        st.markdown(content)

# Input utilizator
user_input = st.chat_input("Scrie un mesaj...")

if user_input:
    with st.chat_message("user"):
        st.markdown(user_input)

    st.session_state.chat_history.append(("user", user_input))

    with st.spinner("Agentul analizează cerința..."):
        response = st.session_state.agent.process_message(user_input)

    with st.chat_message("assistant"):
        st.markdown(response)

    st.session_state.chat_history.append(("assistant", response))