import streamlit as st
import pandas as pd
import os
import logging
import json
from docx import Document
import re

# ========== SETTINGS MANAGEMENT FUNCTIONS ==========
def load_default_headers():
    default_headers = ["Title", "Inventors", "Abstract", "Claims", "Description"]
    try:
        if not os.path.exists(DEFAULT_HEADERS_FILE):
            pd.DataFrame({'header': default_headers}).to_csv(DEFAULT_HEADERS_FILE, index=False)
        return pd.read_csv(DEFAULT_HEADERS_FILE)['header'].tolist()
    except Exception as e:
        logging.error(f"Header load error: {e}")
        return default_headers

def load_column_settings():
    try:
        if os.path.exists(COLUMN_SETTINGS_FILE):
            df = pd.read_csv(COLUMN_SETTINGS_FILE, index_col=0)
            return {**{h: True for h in load_default_headers()}, **df['selected'].to_dict()}
        return {h: True for h in load_default_headers()}
    except Exception as e:
        logging.error(f"Column load error: {e}")
        return {h: True for h in load_default_headers()}

def save_column_settings(settings):
    try:
        existing = pd.read_csv(COLUMN_SETTINGS_FILE, index_col=0) if os.path.exists(COLUMN_SETTINGS_FILE) else pd.DataFrame()
        updated = {**existing.to_dict().get('selected', {}), **settings}
        pd.DataFrame({'selected': updated}).to_csv(COLUMN_SETTINGS_FILE)
    except Exception as e:
        logging.error(f"Column save error: {e}")

def load_filter_settings():
    try:
        if os.path.exists(FILTER_SETTINGS_FILE):
            return pd.read_csv(FILTER_SETTINGS_FILE, index_col=0).to_dict(orient='index')
        return {}
    except Exception as e:
        logging.error(f"Filter load error: {e}")
        return {}

def save_filter_settings(filters):
    try:
        existing = pd.read_csv(FILTER_SETTINGS_FILE, index_col=0) if os.path.exists(FILTER_SETTINGS_FILE) else pd.DataFrame()
        updated = {**existing.to_dict(orient='index'), **filters}
        pd.DataFrame.from_dict(updated, orient='index').to_csv(FILTER_SETTINGS_FILE)
    except Exception as e:
        logging.error(f"Filter save error: {e}")

def save_user_defaults(data):
    try:
        pd.DataFrame.from_dict(data, orient='index', columns=['value']).to_csv(USER_DEFAULTS_FILE)
    except Exception as e:
        logging.error(f"User defaults save error: {e}")

# ========== PANEL RENDERING ==========
def render_settings_panel(panel_id):
    """Render a settings panel with proper UI components"""
    config = PANEL_CONFIG[panel_id]
    with st.expander(config["label"], expanded=True):
        st.markdown(f"##### {config['icon']} {config['label']}")

        updates = {}
        filter_updates = {}

        for i, header in enumerate(config["headers"]()):
            sanitized_key = header.replace('/', '_').replace(' ', '_').replace('&', '_')
            col1, col2, col3 = st.columns([1.2, 1, 2])

            # Add column headers only once
            if i == 0:  # Only for first row
                with col2:
                    st.markdown("<b>Condition</b>", unsafe_allow_html=True)
                with col3:
                    st.markdown("**Value**")

            # Row content
            with col1:
                updates[header] = st.checkbox(
                    header,
                    value=st.session_state.column_settings.get(header, False),
                    key=f"{panel_id}_cb_{sanitized_key}",
                )

            with col2:
                condition = st.selectbox(
                    "Condition",  # Hidden label
                    options=["--", "IS", "BLANK", "CONTAINS", "STARTS WITH", "ENDS WITH"],
                    index=0,
                    key=f"{panel_id}_cond_{sanitized_key}",
                    label_visibility="collapsed"
                )

            with col3:
                value = st.text_input(
                    "Value",  # Hidden label
                    value=st.session_state.filter_settings.get(header, {}).get('value', ""),
                    key=f"{panel_id}_val_{sanitized_key}",
                    label_visibility="collapsed"
                )

            filter_updates[header] = {'condition': condition, 'value': value}

        return updates, filter_updates

# ========== CORE FUNCTIONALITY CLASSES ==========
class PatentProcessor:
    def __init__(self, text):
        self.text = text

    def clean_text(self, text):
        return ' '.join(text.split()).strip()

    def extract_inventors(self):
        match = re.search(r"(?:Inventor(?:s)?):\s*(.+?)(?:\n|<span)", self.text, re.IGNORECASE|re.MULTILINE)
        return [self.clean_text(n.strip()) for n in match.group(1).split(',')] if match else []

    def extract_title(self):
        match = re.search(r"(?:Title):\s*(.+?)(?:\n|<span)", self.text, re.IGNORECASE|re.MULTILINE)
        return self.clean_text(match.group(1)) if match else "No Title Found"

    def extract_abstract(self):
        match = re.search(r"(?:Abstract):\s*(.+?)(?:\n(?:Claims:|Description:))", self.text, re.IGNORECASE|re.MULTILINE|re.DOTALL)
        return self.clean_text(match.group(1)) if match else "No Abstract Found"

    def extract_claims(self):
        match = re.search(r"Claims:(.+?)(?:\nDescription:)", self.text, re.IGNORECASE|re.MULTILINE|re.DOTALL)
        if match:
            claims = re.split(r"\n*\d+\.\s*|\n*and\s*\d+\.\s*", match.group(1).strip())
            return [self.clean_text(c) for c in claims if c.strip()]
        return []

    def extract_description(self):
        match = re.search(r"Description:(.+)", self.text, re.IGNORECASE|re.MULTILINE|re.DOTALL)
        return match.group(1).strip() if match else "No Description Found"

    def analyze(self):
        return {
            "Title": self.extract_title(),
            "Inventors": self.extract_inventors(),
            "Abstract": self.extract_abstract(),
            "Claims": self.extract_claims(),
            "Description": self.extract_description()
        }

# ========== CONSTANTS AND CONFIGURATION ==========
CONFIG_FOLDER = "config"
COLUMN_SETTINGS_FILE = os.path.join(CONFIG_FOLDER, "patent_column_settings.csv")
FILTER_SETTINGS_FILE = os.path.join(CONFIG_FOLDER, "patent_filter_settings.csv")
DEFAULT_HEADERS_FILE = os.path.join(CONFIG_FOLDER, "patent_headers_default.csv")
USER_DEFAULTS_FILE = os.path.join(CONFIG_FOLDER, "patent_user_default.csv")

PANEL_CONFIG = {
    "main": {
        "label": "📋 Main Settings",
        "headers": lambda: [h for h in st.session_state.column_settings if not h.endswith('_AD')],
        "icon": "⚙️"
    },
    "details": {
        "label": "📑 Details Settings",
        "headers": lambda: [h for h in st.session_state.column_settings if h.endswith('_AD')],
        "icon": "🔍"
    }
}

# ========== MAIN APPLICATION ==========
def main():
    st.set_page_config(layout="wide")
    st.title("Patent Information Management System")

    # Increase tab label size
    st.markdown("""
    <style>
    button[data-baseweb="tab"] { font-size: 1.2rem !important; }
    </style>
    """, unsafe_allow_html=True)

    # ========== INITIALIZE SESSION STATE ==========
    if 'column_settings' not in st.session_state:
        st.session_state.column_settings = load_column_settings()
    if 'filter_settings' not in st.session_state:
        st.session_state.filter_settings = load_filter_settings()

    # ========== INPUT FIELDS SECTION ==========
    with st.container():
        st.subheader("Data Input")
        col1, col2 = st.columns(2)

        with col1:
            emails = st.text_area("Emails", height=100, 
                                 help="Enter emails (one per line or comma-separated)")

        with col2:
            app_nums = st.text_area("Application Numbers", height=100,
                                  help="Enter application numbers (one per line or comma-separated)")

        # Save inputs immediately
        save_user_defaults({"emails": emails, "application_numbers": app_nums})

    # ========== TABBED PANELS SECTION ==========
    with st.container():
        col_tabs, col_buttons = st.columns([3, 1])
        with col_tabs:
            # Create tabs with larger labels
            tab_main, tab_details = st.tabs(["📌 MAIN SETTINGS", "📌 DETAILS SETTINGS"])

            # Main Panel
            with tab_main:
                panel_updates_main, filter_updates_main = render_settings_panel("main")

            # Details Panel
            with tab_details:
                panel_updates_details, filter_updates_details = render_settings_panel("details")

            # Combine updates
            panel_updates = {**panel_updates_main, **panel_updates_details}
            filter_updates = {**filter_updates_main, **filter_updates_details}

        with col_buttons:
            st.write("")  # Vertical spacer
            if st.button("💾 Save Settings", type="primary", use_container_width=True):
                st.session_state.column_settings.update(panel_updates)
                st.session_state.filter_settings.update(filter_updates)
                save_column_settings(st.session_state.column_settings)
                save_filter_settings(st.session_state.filter_settings)
                st.toast("Settings saved successfully!", icon="✅")

            if st.button("🔄 Reset to Defaults", type="secondary", use_container_width=True):
                st.session_state.column_settings = {h: True for h in load_default_headers()}
                st.session_state.filter_settings = {}
                save_column_settings(st.session_state.column_settings)
                save_filter_settings(st.session_state.filter_settings)
                st.rerun()

if __name__ == "__main__":
    main()
