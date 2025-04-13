import streamlit as st
import pandas as pd
import os
import logging
import json
from docx import Document
import re
# First define the constants FIRST
CONFIG_FOLDER = "config"
COLUMN_SETTINGS_FILE = os.path.join(CONFIG_FOLDER, "patent_column_settings.csv")
FILTER_SETTINGS_FILE = os.path.join(CONFIG_FOLDER, "patent_filter_settings.csv")
DEFAULT_HEADERS_FILE = os.path.join(CONFIG_FOLDER, "patent_headers_default.csv")  # NOW DEFINED
USER_DEFAULTS_FILE = os.path.join(CONFIG_FOLDER, "patent_user_default.csv")

# THEN perform the initialization check
if not os.path.exists(DEFAULT_HEADERS_FILE):
    st.warning("Initializing default headers...")
    load_default_headers()
    st.rerun()
    
    
# Configure logging
logging.basicConfig(level=logging.INFO,
                    format='%(asctime)s - %(levelname)s - %(message)s')

# Ensure config folder exists
os.makedirs(CONFIG_FOLDER, exist_ok=True)

# --- PatentProcessor Class (Keep your existing PatentProcessor class here) ---


class PatentProcessor:
    def __init__(self, text):
        self.text = text

    def clean_text(self, text):
        text = ' '.join(text.split())
        return text.strip()

    def extract_inventors(self):
        inventor_pattern = r"(?:Inventor(?:s)?):\s*(.+?)(?:\n|<span class=\"math-inline\">)"
        match = re.search(inventor_pattern, self.text,
                          re.IGNORECASE | re.MULTILINE)
        if match:
            return [self.clean_text(name.strip()) for name in match.group(1).split(',')]
        return []

    def extract_title(self):
        title_pattern = r"(?:Title):\s*(.+?)(?:\n|<span class=\"math-inline\">)"
        match = re.search(title_pattern, self.text,
                          re.IGNORECASE | re.MULTILINE)
        if match:
            return self.clean_text(match.group(1))
        return "No Title Found"

    def extract_abstract(self):
        abstract_pattern = r"(?:Abstract):\s*(.+?)(?:\n(?:Claims:|Description:|<span class=\"math-inline\">))"
        match = re.search(abstract_pattern, self.text,
                          re.IGNORECASE | re.MULTILINE | re.DOTALL)
        if match:
            return self.clean_text(match.group(1))
        return "No Abstract Found"

    def extract_claims(self):
        claims_pattern = r"(?:Claims:)(.+?)(?:\nDescription:|<span class=\"math-inline\">)"
        match = re.search(claims_pattern, self.text,
                          re.IGNORECASE | re.MULTILINE | re.DOTALL)
        if match:
            claims_text = match.group(1).strip()
            claims = re.split(
                r"\n*\d+\.\s*|\n*and\s*\d+\.\s*", claims_text)
            return [self.clean_text(claim).strip() for claim in claims if claim.strip()]
        return []

    def extract_description(self):
        description_pattern = r"(?:Description:)(.+)"
        match = re.search(description_pattern, self.text,
                          re.IGNORECASE | re.MULTILINE | re.DOTALL)
        if match:
            return match.group(1).strip()
        return "No Description Found"

    def analyze(self):
        title = self.extract_title()
        inventors = self.extract_inventors()
        abstract = self.extract_abstract()
        claims = self.extract_claims()
        description = self.extract_description()
        return {
            "Title": title,
            "Inventors": inventors,
            "Abstract": abstract,
            "Claims": claims,
            "Description": description
        }

# --- Helper Functions for Settings ---


# --- Modified load_default_headers Function ---
def load_default_headers():
    filepath = DEFAULT_HEADERS_FILE
    default_headers = ["Title", "Inventors", "Abstract", "Claims", "Description"]  # Core headers
    
    try:
        if not os.path.exists(filepath):
            # Create default headers file if missing
            pd.DataFrame({'header': default_headers}).to_csv(filepath, index=False)
            logging.info(f"Created default headers file at: {filepath}")
            
        # Always load from file to ensure synchronization
        df = pd.read_csv(filepath)
        return df['header'].tolist()
    
    except Exception as e:
        logging.error(f"Error in load_default_headers: {e}")
        return default_headers  # Fallback to hardcoded defaults

# --- Modified load_column_settings Function ---
def load_column_settings():
    """Load settings while preserving existing columns"""
    filepath = COLUMN_SETTINGS_FILE
    default_headers = load_default_headers()
    
    try:
        if os.path.exists(filepath):
            df = pd.read_csv(filepath, index_col=0)
            # Merge existing settings with default headers
            return {**{h: True for h in default_headers}, **df['selected'].to_dict()}
        return {h: True for h in default_headers}
    except Exception as e:
        return {h: True for h in default_headers}


def save_column_settings(main_settings, detail_settings):
    """Save settings without overwriting other columns"""
    filepath = COLUMN_SETTINGS_FILE
    try:
        # Load existing settings
        existing = pd.read_csv(filepath, index_col=0) if os.path.exists(filepath) else pd.DataFrame()
        
        # Update only the relevant sections
        updated_settings = {**existing.to_dict().get('selected', {}),
                           **main_settings,
                           **detail_settings}
        
        pd.DataFrame({'selected': updated_settings}).to_csv(filepath)
        logging.info(f"Saved column settings to {filepath}")
    except Exception as e:
        logging.error(f"Error saving column settings: {e}")


def load_filter_settings():
    """Load filter settings with empty defaults"""
    filepath = FILTER_SETTINGS_FILE
    try:
        if os.path.exists(filepath):
            df = pd.read_csv(filepath, index_col=0)
            return df.to_dict(orient='index')
        return {}
    except Exception as e:
        return {}


def save_filter_settings(main_filters, detail_filters):
    """Save filter settings without overwriting other columns"""
    filepath = FILTER_SETTINGS_FILE
    try:
        # Load existing settings
        existing = pd.read_csv(filepath, index_col=0) if os.path.exists(filepath) else pd.DataFrame()
        
        # Update only the relevant sections
        updated_filters = {**existing.to_dict(orient='index'),
                          **main_filters,
                          **detail_filters}
        
        pd.DataFrame.from_dict(updated_filters, orient='index').to_csv(filepath)
        logging.info(f"Saved filter settings to {filepath}")
    except Exception as e:
        logging.error(f"Error saving filter settings: {e}")


def load_user_defaults():
    filepath = USER_DEFAULTS_FILE
    defaults = {"emails": "", "application_numbers": ""}
    if os.path.exists(filepath):
        logging.info(f"Loading user defaults from: {filepath}")
        try:
            df = pd.read_csv(filepath, index_col=0)
            if 'value' in df.columns:
                defaults["emails"] = df.loc["emails",
                                            "value"] if "emails" in df.index else ""
                defaults["application_numbers"] = df.loc["application_numbers",
                                                         "value"] if "application_numbers" in df.index else ""
        except Exception as e:
            logging.error(f"Error loading user defaults from {filepath}: {e}")
    else:
        logging.info(
            f"User defaults file not found at: {filepath}. Creating default.")
        save_user_defaults(defaults)
    return defaults


def save_user_defaults(data):
    filepath = USER_DEFAULTS_FILE
    logging.info(f"Saving user defaults to: {filepath}")
    df = pd.DataFrame.from_dict(data, orient='index', columns=['value'])
    df.index.name = 'setting'
    df.to_csv(filepath)


def filter_dataframe(df, filters):
    for column, conditions in filters.items():
        condition = conditions['condition']
        value = conditions['value']
        if condition == "IS":
            df = df[df[column].astype(str) == value]
        elif condition == "BLANK":
            df = df[df[column].astype(str).str.strip() == ""]
        elif condition == "CONTAINS":
            df = df[df[column].astype(str).str.contains(
                value, case=False, na=False)]
        elif condition == "STARTS WITH":
            df = df[df[column].astype(str).str.startswith(value, na=False)]
        elif condition == "ENDS WITH":
            df = df[df[column].astype(str).str.endswith(value, na=False)]
    return df


def save_to_json(data, filename="patent_analysis.json"):
    with open(filename, 'w') as f:
        json.dump(data, f, indent=4)
    st.success(f"Data saved to {filename}")


def save_to_csv(data, filename="patent_analysis.csv"):
    df = pd.DataFrame([data])
    df.to_csv(filename, index=False)
    st.success(f"Data saved to {filename}")


def save_to_docx(data, filename="patent_analysis.docx"):
    document = Document()
    document.add_heading(data.get("Title", "Patent Analysis"), level=1)
    document.add_heading("Inventors", level=2)
    inventors = data.get("Inventors")
    if inventors:
        document.add_paragraph(", ".join(inventors))
    else:
        document.add_paragraph("No inventors found.")
    document.add_heading("Abstract", level=2)
    document.add_paragraph(data.get("Abstract", "No abstract found."))
    document.add_heading("Claims", level=2)
    claims = data.get("Claims")
    if claims:
        for i, claim in enumerate(claims):
            document.add_paragraph(f"{i+1}. {claim}")
    else:
        document.add_paragraph("No claims found.")
    document.add_heading("Description", level=2)
    document.add_paragraph(data.get("Description", "No description found."))
    document.save(filename)
    st.success(f"Data saved to {filename}")

# --- Streamlit App with Checkbox-Based Settings and Input Fields ---


# Set page to wide mode
st.set_page_config(layout="wide")

# Center the title
st.markdown("<h1 style='text-align: center;'>Application info retrieval</h1>",
            unsafe_allow_html=True)

# Load settings FIRST
column_settings = load_column_settings()
filter_settings = load_filter_settings()
user_defaults = load_user_defaults()

# THEN create valid headers list
valid_headers = load_default_headers() + [h for h in column_settings if h.endswith("_AD")]

# Remove duplicates and ensure proper order
valid_headers = list(dict.fromkeys(valid_headers))  # Preserves order while deduping

# FINALLY define header categories
main_headers = [h for h in valid_headers if not h.endswith("_AD")]
detail_headers = [h for h in valid_headers if h.endswith("_AD")]
all_headers = main_headers + detail_headers

# Debug output
# st.sidebar.write("Valid headers:", valid_headers)
# st.sidebar.write("Main headers:", main_headers)

# Add explicit header validation
valid_headers = load_default_headers() + [h for h in column_settings if h.endswith("_AD")]



# Enforce valid headers in settings
column_settings = {k: v for k, v in column_settings.items() if k in all_headers}
filter_settings = {k: v for k, v in filter_settings.items() if k in all_headers}

settings_changed = False

# --- Input Fields at the Top ---
st.subheader("Enter Emails and Application Numbers")
col_emails, col_app_nums = st.columns(2)

with col_emails:
    emails_input = st.text_area("Emails", key="emails_input", value=st.session_state.get(
        "emails_input", user_defaults["emails"]), height=150, help="Enter one per line or delimited by commas")
    if st.button("Clear Emails"):
        st.session_state["emails_input"] = ""
        st.rerun()

with col_app_nums:
    app_nums_input = st.text_area("Application Numbers", key="app_nums_input", value=st.session_state.get(
        "app_nums_input", user_defaults["application_numbers"]), height=150, help="Enter one per line or delimited by commas")
    if st.button("Clear App. Nos"):
        st.session_state["app_nums_input"] = ""
        st.rerun()

# Save user inputs
user_data_to_save = {"emails": st.session_state.get("emails_input", user_defaults["emails"]),
                     "application_numbers": st.session_state.get("app_nums_input", user_defaults["application_numbers"])}
save_user_defaults(user_data_to_save)

# Initialize session state for settings
if 'column_settings' not in st.session_state:
    st.session_state.column_settings = load_column_settings()
if 'filter_settings' not in st.session_state:
    st.session_state.filter_settings = load_filter_settings()

# --- Modified Settings Panels ---
main_updates = {}
detail_updates = {}
filter_main_updates = {}
filter_detail_updates = {}


# --- Filtering Panels Section ---
st.subheader("Filter Settings")

# Initialize session state keys
if 'toggle_main_settings' not in st.session_state:
    st.session_state.toggle_main_settings = False
if 'toggle_detail_settings' not in st.session_state:
    st.session_state.toggle_detail_settings = False

def handle_panel_toggles():
    """Callback to manage mutually exclusive panels"""
    # Get which button was last clicked
    if st.session_state.get('toggle_main_btn'):
        st.session_state.toggle_main_settings = st.session_state.toggle_main_btn
        st.session_state.toggle_detail_settings = False
    elif st.session_state.get('toggle_detail_btn'):
        st.session_state.toggle_detail_settings = st.session_state.toggle_detail_btn
        st.session_state.toggle_main_settings = False

# Create checkboxes with separate widget keys
col_show_main, col_show_detail, col_reset_filters = st.columns([1, 1, 1])
with col_show_main:
    st.checkbox(
        "📋 Show MAIN Settings",
        key="toggle_main_btn",
        value=st.session_state.toggle_main_settings,
        on_change=handle_panel_toggles,
        help="Display main filtering options"
    )

with col_show_detail:
    st.checkbox(
        "📑 Show DETAILS Settings",
        key="toggle_detail_btn",
        value=st.session_state.toggle_detail_settings,
        on_change=handle_panel_toggles,
        help="Display advanced detail options"
    )

# Panel display logic
if st.session_state.toggle_main_settings:
    with st.expander("MAIN FILTER SETTINGS", expanded=True):
        main_updates = {}
        filter_main_updates = {}
        
        for header in [h for h in st.session_state.column_settings if not h.endswith('_AD')]:
            with st.container():
                sanitized_key = header.replace('/', '_').replace(' ', '_').replace('&', '_')
                col1, col2, col3 = st.columns([1.2, 1, 2])
                
                with col1:
                    main_updates[header] = st.checkbox(
                        header, 
                        value=st.session_state.column_settings.get(header, False),
                        key=f"main_checkbox_{sanitized_key}"
                    )
                
                with col2:
                    condition = st.selectbox(
                        "Condition", 
                        options=["--", "IS", "BLANK", "CONTAINS", "STARTS WITH", "ENDS WITH"],
                        index=0,
                        key=f"main_condition_{sanitized_key}"
                    )
                
                with col3:
                    value = st.text_input(
                        "Value", 
                        value=st.session_state.filter_settings.get(header, {}).get('value', ''),
                        key=f"main_value_{sanitized_key}"
                    )
                
                filter_main_updates[header] = {'condition': condition, 'value': value}

elif st.session_state.toggle_detail_settings:
    with st.expander("DETAILS FILTER SETTINGS", expanded=True):
        detail_updates = {}
        filter_detail_updates = {}
        
        for header in [h for h in st.session_state.column_settings if h.endswith('_AD')]:
            with st.container():
                base_header = header.replace('_AD', '')
                sanitized_key = base_header.replace('/', '_').replace(' ', '_').replace('&', '_')
                
                col1, col2, col3 = st.columns([1.2, 1, 2])
                
                with col1:
                    detail_updates[header] = st.checkbox(
                        header, 
                        value=st.session_state.column_settings.get(header, False),
                        key=f"detail_checkbox_{sanitized_key}"
                    )
                
                with col2:
                    condition = st.selectbox(
                        "Condition", 
                        options=["--", "IS", "BLANK", "CONTAINS", "STARTS WITH", "ENDS WITH"],
                        index=0,
                        key=f"detail_condition_{sanitized_key}"
                    )
                
                with col3:
                    value = st.text_input(
                        "Value", 
                        value=st.session_state.filter_settings.get(header, {}).get('value', ''),
                        key=f"detail_value_{sanitized_key}"
                    )
                
                filter_detail_updates[header] = {'condition': condition, 'value': value}

else:
    st.write("ℹ️ Select either MAIN or DETAILS settings above to configure filters")

# Save button (appears only when a panel is active)
if st.session_state.toggle_main_settings or st.session_state.toggle_detail_settings:
    if st.button("💾 Save Current Settings", type="primary", use_container_width=True):
        # Update session state
        if st.session_state.toggle_main_settings:
            st.session_state.column_settings.update(main_updates)
            st.session_state.filter_settings.update(filter_main_updates)
        else:
            st.session_state.column_settings.update(detail_updates)
            st.session_state.filter_settings.update(filter_detail_updates)
        
        # Save to files
        pd.DataFrame.from_dict({'selected': st.session_state.column_settings}, orient='index') \
                   .to_csv(COLUMN_SETTINGS_FILE)
        pd.DataFrame.from_dict(st.session_state.filter_settings, orient='index') \
                   .to_csv(FILTER_SETTINGS_FILE)
        
        st.toast("Settings saved successfully!", icon="✅")
        st.rerun()
        
with col_reset_filters:
    if st.button(":orange[Reset Filters]", use_container_width=True):
        # Reset GUI filter entries
        for header in main_headers:
            st.session_state[f"main_condition_{header}"] = "--"
            st.session_state[f"main_value_{header}"] = ""
        for header in detail_headers:
            st.session_state[f"detail_condition_{header}"] = "--"
            st.session_state[f"detail_value_{header}"] = ""

        # Reset saved filter settings
        default_filter_settings = {
            header: {'condition': "--", 'value': ""} for header in all_headers}
        save_filter_settings(default_filter_settings)

        # Optionally, also reset column visibility to default
        default_column_settings = {header: True for header in all_headers}
        save_column_settings(default_column_settings)

        st.rerun()

# --- Add Submit Button at Bottom of Filter Panels ---
if st.button("💾 Save All Settings", use_container_width=True):
    with st.spinner("Saving settings..."):
        try:
            # Save column settings
            save_column_settings(main_updates, detail_updates)
            
            # Save filter settings
            save_filter_settings(filter_main_updates, filter_detail_updates)
            
            st.toast("Settings saved successfully!", icon="✅")
            st.rerun()
        except Exception as e:
            st.error(f"Error saving settings: {str(e)}")
            

# Progress/Log Area at the bottom
log_text_area = st.text_area("Log Messages", height=150, disabled=True)


@st.cache_resource
def setup_logging(log_text_area_key):
    logger = logging.getLogger()
    # Create a Streamlit handler that writes to the text area
    handler = StreamHandlerTextArea(st.session_state, log_text_area_key)
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    handler.setFormatter(formatter)
    logger.addHandler(handler)
    logger.setLevel(logging.INFO)
    return logger


class StreamHandlerTextArea(logging.Handler):
    def __init__(self, session_state, text_area_key):
        logging.Handler.__init__(self)
        self.session_state = session_state
        self.text_area_key = text_area_key

    def emit(self, record):
        try:
            msg = self.format(record)
            new_value = self.session_state.get(
                self.text_area_key, "") + msg + "\n"
            self.session_state[self.text_area_key] = new_value
        except Exception:
            self.handleError(record)


logger = setup_logging(log_text_area)

# About section
st.subheader("About")
st.info(
    "This app allows you to configure display and filter settings for patent data. "
    "Enter emails and application numbers (one per line or comma-separated) at the top for online retrieval. "
    "Use the blue and yellow checkboxes to show/hide the main and detailed settings panels for filtering. "
    "Click the orange 'Reset Filters' button to clear the filter entries in the GUI and revert to default filter settings. "
    "Click the 'Retrieve Data' button to initiate the online data retrieval process."
)

# --- Placeholder for your online retrieval function ---


def retrieve_online_data(identifier, data_type):
    """
    This is a placeholder function. Replace it with your actual logic
    to retrieve patent information online based on email or application number.
    """
    import time
    time.sleep(2)  # Simulate network delay
    if data_type == "email":
        return f"Simulated data retrieved for email: {identifier} at {time.strftime('%Y-%m-%d %H:%M:%S')}"
    elif data_type == "application_number":
        return f"Simulated data retrieved for application number: {identifier} at {time.strftime('%Y-%m-%d %H:%M:%S')}"
    return None
