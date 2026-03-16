import streamlit as st
import pandas as pd
import os
from pathlib import Path
import pytesseract
from PIL import Image
import pdfplumber
import datetime
from io import BytesIO
import time
import hashlib
import base64
import threading
import json

# =========================================
# CONFIG - This must be the first Streamlit command
# =========================================
st.set_page_config(page_title="Hostel Management System", layout="wide", initial_sidebar_state="collapsed")

# =========================================
# CONSTANTS - FILE PATHS
# =========================================
D_DRIVE_PATH = Path("D:/")
HOSTEL_DATA_PATH = D_DRIVE_PATH / "HostelData"
HOSTEL_PROCESSED_PATH = D_DRIVE_PATH / "HostelData_Processed"
HOSTEL_SYSTEM_PATH = D_DRIVE_PATH / "HostelSystem"

# Create directories
try:
    HOSTEL_DATA_PATH.mkdir(exist_ok=True)
    HOSTEL_PROCESSED_PATH.mkdir(exist_ok=True)
    HOSTEL_SYSTEM_PATH.mkdir(exist_ok=True)
except Exception as e:
    st.error(f"Error creating directories: {e}")

BASE_DIR = Path(__file__).parent
DATA_DIR = BASE_DIR / "data"
DATA_DIR.mkdir(exist_ok=True)

STUDENTS_FILE = DATA_DIR / "students.xlsx"
HISTORY_FILE = DATA_DIR / "history.xlsx"
FORMS_FILE = DATA_DIR / "forms.xlsx"
DEFAULTERS_FILE = DATA_DIR / "defaulters.xlsx"
PWWF_BOARDING_FILE = DATA_DIR / "pwwf_boarding.xlsx"
PROCESSED_FILES_LOG = DATA_DIR / "processed_files.txt"
ROOM_CAPACITIES_FILE = DATA_DIR / "room_capacities.json"

# Logo path
LOGO_PATH = BASE_DIR / "comsats_logo.png"
PROFILE_PIC_PATH = BASE_DIR / "pp (1).png"

# =========================================
# REMOVE DEPLOY MESSAGE CSS
# =========================================
def remove_deploy_message():
    st.markdown("""
        <style>
        /* Hide the deploy button and related messages */
        .stApp header {display: none !important;}
        .stApp footer {display: none !important;}
        .stDeployButton {display: none !important;}
        .stApp [data-testid="stToolbar"] {display: none !important;}
        .stApp [data-testid="baseButton-header"] {display: none !important;}
        .stApp [data-testid="stStatusWidget"] {display: none !important;}
        .st-emotion-cache-18ni7ap {display: none !important;}
        .st-emotion-cache-1dp5vir {display: none !important;}
        .st-emotion-cache-1mi2ry5 {display: none !important;}
        </style>
    """, unsafe_allow_html=True)

# =========================================
# CUSTOM CSS
# =========================================
def apply_custom_css():
    dark_blue = "#1E2B38"
    darker_blue = "#15232E"
    light_text = "#FFFFFF"
    bright_text = "#FFD700"
    
    st.markdown(f"""
        <style>
        /* Hide default Streamlit elements */
        .stApp header {{display: none !important;}}
        .stApp footer {{display: none !important;}}
        .stDeployButton {{display: none !important;}}
        .stApp [data-testid="stToolbar"] {{display: none !important;}}
        .stApp [data-testid="baseButton-header"] {{display: none !important;}}
        .stApp [data-testid="stStatusWidget"] {{display: none !important;}}
        .st-emotion-cache-18ni7ap {{display: none !important;}}
        .st-emotion-cache-1dp5vir {{display: none !important;}}
        .st-emotion-cache-1mi2ry5 {{display: none !important;}}
        .eczjsme11 {{display: none !important;}}
        
        /* Main app background */
        .stApp {{ background-color: {dark_blue}; }}
        .stMarkdown, .stText, p, li, span {{ color: {light_text} !important; }}
        h1, h2, h3, h4, h5, h6 {{ color: {bright_text} !important; font-weight: bold; }}
        .streamlit-expanderHeader {{ color: {bright_text} !important; background-color: {darker_blue} !important; }}
        .stTextInput label {{
            color: #FFD700 !important;
            font-weight: bold !important;
        }}
        .stSelectbox label, .stNumberInput label {{ color: {bright_text} !important; }}
        .stTextInput input, .stNumberInput input {{ color: #000000 !important; background-color: white !important; }}
        
        /* Dropdown Styling */
        .stSelectbox div[data-baseweb='select'] {{
            background-color: #1E2B38 !important;
            color: white !important;
        }}
        
        /* Fix for selectbox options visibility */
        .stSelectbox div[data-baseweb="select"] > div {{
            background-color: {dark_blue} !important;
            color: {light_text} !important;
        }}
        .stSelectbox div[data-baseweb="select"] ul {{
            background-color: {dark_blue} !important;
        }}
        .stSelectbox div[data-baseweb="select"] li {{
            color: {light_text} !important;
            background-color: {dark_blue} !important;
        }}
        .stSelectbox div[data-baseweb="select"] li:hover {{
            background-color: {darker_blue} !important;
        }}
        
        /* Table Toolbar Icons - Keep visible but restyle */
        button[data-testid='stToolbar'] {{
            display: block !important;
            background-color: {darker_blue} !important;
            color: {light_text} !important;
            border: 1px solid {bright_text} !important;
        }}
        
        /* Fix for column menu options */
        .stDataFrame div[data-testid="stDataFrameResizable"] div[role="listbox"] {{
            background-color: {dark_blue} !important;
            border: 1px solid {light_text} !important;
        }}
        .stDataFrame div[data-testid="stDataFrameResizable"] div[role="option"] {{
            color: {light_text} !important;
            background-color: {dark_blue} !important;
        }}
        .stDataFrame div[data-testid="stDataFrameResizable"] div[role="option"]:hover {{
            background-color: {darker_blue} !important;
        }}
        
        /* Fix for filter dropdowns */
        div[data-baseweb="select"] > div {{
            background-color: {dark_blue} !important;
            color: {light_text} !important;
        }}
        div[data-baseweb="select"] ul {{
            background-color: {dark_blue} !important;
        }}
        div[data-baseweb="select"] li {{
            color: {light_text} !important;
            background-color: {dark_blue} !important;
        }}
        div[data-baseweb="select"] li:hover {{
            background-color: {darker_blue} !important;
        }}
        
        .stButton > button, .stDownloadButton > button {{ 
            background-color: {dark_blue} !important; 
            color: {light_text} !important; 
            border: 1px solid {light_text} !important;
        }}
        .stButton > button:hover, .stDownloadButton > button:hover {{ 
            background-color: {darker_blue} !important; 
            border: 1px solid {bright_text} !important;
        }}
        .stDataFrame {{ background-color: white; color: #2C3E50; }}
        .stAlert {{ background-color: {darker_blue}; color: {light_text}; }}
        
        /* Style for file uploader area */
        .stFileUploader > div {{
            background-color: {dark_blue} !important;
            border: 1px dashed {light_text} !important;
            border-radius: 5px !important;
            padding: 20px !important;
        }}
        .stFileUploader > div > div {{
            color: {light_text} !important;
        }}
        .stFileUploader > div > div > small {{
            color: {light_text} !important;
        }}
        .stFileUploader > div > div > div {{
            color: {light_text} !important;
        }}
        .stFileUploader > div > button {{
            background-color: {darker_blue} !important;
            color: {light_text} !important;
            border: 1px solid {light_text} !important;
        }}
        .stFileUploader > div > button:hover {{
            background-color: {dark_blue} !important;
            border: 1px solid {bright_text} !important;
        }}
        
        /* Login page specific */
        .login-container {{
            background-color: rgba(30,43,56,0.85);
            padding: 30px;
            border-radius: 10px;
            box-shadow: 0 4px 6px rgba(0,0,0,0.3);
            backdrop-filter: blur(5px);
            margin-top: 50px;
        }}
        
        /* Profile picture styling */
        .profile-pic {{
            width: 40px;
            height: 40px;
            border-radius: 50%;
            border: 2px solid {bright_text};
            object-fit: cover;
        }}
        
        /* Main content area */
        .main .block-container {{
            padding-top: 1rem;
            padding-bottom: 0rem;
        }}
        
        /* Remove extra whitespace */
        .css-18e3th9 {{
            padding-top: 0rem;
            padding-bottom: 0rem;
        }}
        .css-1d391kg {{
            padding-top: 0rem;
        }}
        </style>
    """, unsafe_allow_html=True)

# =========================================
# ROOM CAPACITIES FUNCTIONS
# =========================================
def save_room_capacities():
    """Save room capacities to a file"""
    if 'room_capacities' in st.session_state:
        try:
            with open(ROOM_CAPACITIES_FILE, 'w') as f:
                json.dump(st.session_state.room_capacities, f)
        except Exception as e:
            st.error(f"Error saving room capacities: {e}")

def load_room_capacities():
    """Load room capacities from file"""
    if ROOM_CAPACITIES_FILE.exists():
        try:
            with open(ROOM_CAPACITIES_FILE, 'r') as f:
                st.session_state.room_capacities = json.load(f)
        except Exception as e:
            st.error(f"Error loading room capacities: {e}")
    else:
        st.session_state.room_capacities = {}

# =========================================
# PASSWORD FUNCTIONS
# =========================================
def hash_password(password):
    return hashlib.sha256(password.encode()).hexdigest()

RECOVERY_EMAIL = "umermukhtar013@gmail.com"

# =========================================
# LOGIN PAGE WITH BACKGROUND IMAGE
# =========================================
def get_login_background():
    """Load image2.png as login background"""
    possible_paths = [
        Path(__file__).parent / "image2.png",
        HOSTEL_DATA_PATH / "image2.png",
        Path("D:/image2.png")
    ]

    for path in possible_paths:
        if path.exists():
            try:
                with open(path, "rb") as f:
                    img_data = f.read()
                b64_encoded = base64.b64encode(img_data).decode()
                return f"url(data:image/png;base64,{b64_encoded})"
            except:
                continue
    return None

def login_page():
    bg_image = get_login_background()
    
    if bg_image:
        st.markdown(f"""
            <style>
            .stApp {{
                background-image: {bg_image};
                background-size: cover;
                background-position: center;
                background-repeat: no-repeat;
                background-attachment: fixed;
            }}
            </style>
        """, unsafe_allow_html=True)
    else:
        st.markdown("""
            <style>
            .stApp {
                background-color: #1E2B38;
            }
            </style>
        """, unsafe_allow_html=True)
    
    # Remove deploy message on login page too
    remove_deploy_message()
    
    col1, col2, col3 = st.columns([1,2,1])
    
    with col2:
        st.markdown('<div class="login-container">', unsafe_allow_html=True)
        
        if LOGO_PATH.exists():
            try:
                logo_image = Image.open(LOGO_PATH)
                st.image(logo_image, width=200)
            except:
                st.markdown("<h1 style='color:#FFD700; text-align:center;'>Hostel Management System</h1>", unsafe_allow_html=True)
        else:
            st.markdown("<h1 style='color:#FFD700; text-align:center;'>Hostel Management System</h1>", unsafe_allow_html=True)
        
        st.markdown("---")
        
        with st.form("login_form"):
            username = st.text_input("Username")
            password = st.text_input("Password", type="password")
            
            col_a, col_b = st.columns(2)
            with col_a:
                submit = st.form_submit_button("Login", use_container_width=True)
            with col_b:
                forgot = st.form_submit_button("Forgot Password", use_container_width=True)
            
            if submit:
                if username and password:
                    if check_login(username, password):
                        st.session_state.logged_in = True
                        st.session_state.username = username
                        st.session_state.menu = 'Dashboard'
                        load_user_profile(username)
                        st.rerun()
                    else:
                        st.error("Invalid username or password")
                else:
                    st.warning("Please enter username and password")
            
            if forgot:
                st.session_state.show_forgot = True
                st.rerun()
        
        st.markdown('</div>', unsafe_allow_html=True)

def forgot_password_page():
    bg_image = get_login_background()
    
    if bg_image:
        st.markdown(f"""
            <style>
            .stApp {{
                background-image: {bg_image};
                background-size: cover;
                background-position: center;
                background-repeat: no-repeat;
                background-attachment: fixed;
            }}
            </style>
        """, unsafe_allow_html=True)
    
    # Remove deploy message on forgot password page too
    remove_deploy_message()
    
    col1, col2, col3 = st.columns([1,2,1])
    
    with col2:
        st.markdown('<div class="login-container">', unsafe_allow_html=True)
        
        if LOGO_PATH.exists():
            try:
                logo_image = Image.open(LOGO_PATH)
                st.image(logo_image, width=200)
            except:
                st.markdown("<h1 style='color:#FFD700; text-align:center;'>Hostel Management System</h1>", unsafe_allow_html=True)
        else:
            st.markdown("<h1 style='color:#FFD700; text-align:center;'>Hostel Management System</h1>", unsafe_allow_html=True)
        
        st.markdown("<h2 style='color:#FFD700; text-align:center;'>Forgot Password</h2>", unsafe_allow_html=True)
        st.info(f"Recovery email: {RECOVERY_EMAIL}")
        
        with st.form("forgot_form"):
            email = st.text_input("Enter your email")
            new_password = st.text_input("New Password", type="password")
            confirm_password = st.text_input("Confirm Password", type="password")
            
            col_a, col_b = st.columns(2)
            with col_a:
                submit = st.form_submit_button("Reset Password", use_container_width=True)
            with col_b:
                back = st.form_submit_button("Back to Login", use_container_width=True)
            
            if submit:
                if email == RECOVERY_EMAIL:
                    if new_password == confirm_password and new_password:
                        st.success("Password reset successful! Please login with new password.")
                        time.sleep(2)
                        st.session_state.show_forgot = False
                        st.rerun()
                    else:
                        st.error("Passwords do not match or are empty")
                else:
                    st.error("Email not authorized for recovery")
            
            if back:
                st.session_state.show_forgot = False
                st.rerun()
        
        st.markdown('</div>', unsafe_allow_html=True)

def check_login(username, password):
    users_file = DATA_DIR / "users.xlsx"
    if users_file.exists():
        df = pd.read_excel(users_file)
        return any((df['username'] == username) & (df['password'] == hash_password(password)))
    return False

def load_user_profile(username):
    """Load user profile data from file"""
    profiles_file = DATA_DIR / "profiles.xlsx"
    if profiles_file.exists():
        df = pd.read_excel(profiles_file)
        user_data = df[df['username'] == username]
        if not user_data.empty:
            st.session_state.officer_name = user_data.iloc[0].get('officer_name', username)
            st.session_state.campus = user_data.iloc[0].get('campus', 'ISB')
            st.session_state.designation = user_data.iloc[0].get('designation', 'Officer')
            st.session_state.user_id = user_data.iloc[0].get('user_id', f'EMP{hash(username) % 1000:03d}')
        else:
            st.session_state.officer_name = username
            st.session_state.campus = 'ISB'
            st.session_state.designation = 'Officer'
            st.session_state.user_id = f'EMP{hash(username) % 1000:03d}'
    else:
        st.session_state.officer_name = username
        st.session_state.campus = 'ISB'
        st.session_state.designation = 'Officer'
        st.session_state.user_id = f'EMP{hash(username) % 1000:03d}'
    
    # Save profile data to ensure it persists
    save_profile_data()

# =========================================
# SEMESTER/PROGRAM DETECTION
# =========================================
def detect_semester(reg):
    reg = str(reg).upper()
    mapping = {'FA20':12,'SP21':11,'FA21':10,'SP22':9,'FA22':8,'SP23':7,
               'FA23':6,'SP24':5,'FA24':4,'SP25':3,'FA25':2,'SP26':1}
    for code, sem in mapping.items():
        if code in reg:
            return sem
    return 1

def detect_program(reg):
    reg = str(reg).upper()
    for code in ['BCS','BBA','BEE','BSCS','BSSE','MCS','MBA','MSCS']:
        if code in reg:
            return code
    parts = reg.split('-')
    return parts[1] if len(parts) >= 2 else "Unknown"

# =========================================
# Semester Format Conversion
# =========================================
def format_semester(sem):
    try:
        sem = int(sem)
        if sem == 1:
            return '1st'
        elif sem == 2:
            return '2nd'
        elif sem == 3:
            return '3rd'
        else:
            return f'{sem}th'
    except:
        return sem

# =========================================
# ROOM SORTING
# =========================================
ROOM_ORDER = [
    'A-2','A-3','A-4','A-5','A-6','A-7','A-8','A-9','A-12','A-12.1',
    'A-13','A-14','A-15','A-16','A-17','A-18','A-19','A-20','A-21','A-22','A-23','A-24',
    'B-12','B-13','B-14','B-15','B-16','B-17','B-18','B-19','B-20','B-21','B-22',
    'B-1','B-2','B-3','B-4','B-5','B-6','B-7','B-8','B-9','B-10','B-11',
    'C-1','C-2','C-3','C-4','C-5','C-6','C-7','C-8','C-9','C-10','C-11','C-12'
]
ROOM_ORDER_DICT = {room: i for i, room in enumerate(ROOM_ORDER)}

# =========================================
# DATA MANAGEMENT
# =========================================
STUDENT_COLUMNS = ["SR#","Name","Registration No","Room No","Status","Contact No","Father Contact","Blood Group","Semester","Program"]
HISTORY_COLUMNS = ["Date","Time","Student Name","Registration No","Room No","Status","Semester","Amount","Payment Method","Remarks"]
DEFAULTERS_COLUMNS = ["Student Name","Registration No","Room No","Status","Semester","Amount","Defaulter Status","Remarks"]
FORMS_COLUMNS = ["Student Name","Registration No","Room No","Status","Semester","Admission Form","PWWF Form","Consent Form","Amount"]
PWWF_BOARDING_COLUMNS = ["SR#","Student Name","Registration No","Semester","Amount","Paying Date"]

def load_data():
    # Load room capacities
    load_room_capacities()
    
    # Students
    if STUDENTS_FILE.exists():
        df = pd.read_excel(STUDENTS_FILE)
        for col in STUDENT_COLUMNS:
            if col not in df.columns:
                if col == "SR#":
                    df[col] = range(1, len(df)+1)
                elif col == "Semester":
                    df[col] = df["Registration No"].apply(detect_semester)
                elif col == "Program":
                    df[col] = df["Registration No"].apply(detect_program)
                else:
                    df[col] = ""
        df["Status"] = df["Status"].fillna("Open")
        df.loc[~df["Status"].isin(["Open", "PWWF"]), "Status"] = "Open"
        st.session_state.students = df[STUDENT_COLUMNS]
    else:
        st.session_state.students = pd.DataFrame(columns=STUDENT_COLUMNS)
    
    # History
    if HISTORY_FILE.exists():
        df = pd.read_excel(HISTORY_FILE)
        for col in df.columns:
            df[col] = df[col].astype(str)
        st.session_state.history = df
    else:
        st.session_state.history = pd.DataFrame(columns=HISTORY_COLUMNS)
    
    # Defaulters
    if DEFAULTERS_FILE.exists():
        df = pd.read_excel(DEFAULTERS_FILE)
        for col in df.columns:
            df[col] = df[col].astype(str)
        st.session_state.defaulters = df
    else:
        st.session_state.defaulters = pd.DataFrame(columns=DEFAULTERS_COLUMNS)
    
    # Forms
    if FORMS_FILE.exists():
        df = pd.read_excel(FORMS_FILE)
        for col in df.columns:
            df[col] = df[col].astype(str)
        st.session_state.forms = df
    else:
        st.session_state.forms = pd.DataFrame(columns=FORMS_COLUMNS)
    
    # PWWF Boarding
    if PWWF_BOARDING_FILE.exists():
        df = pd.read_excel(PWWF_BOARDING_FILE)
        for col in df.columns:
            df[col] = df[col].astype(str)
        if "SR#" not in df.columns:
            df.insert(0, "SR#", range(1, len(df)+1))
        st.session_state.pwwf_boarding = df
    else:
        st.session_state.pwwf_boarding = pd.DataFrame(columns=PWWF_BOARDING_COLUMNS)

def save_data():
    try:
        st.session_state.students.to_excel(STUDENTS_FILE, index=False)
        st.session_state.history.to_excel(HISTORY_FILE, index=False)
        st.session_state.defaulters.to_excel(DEFAULTERS_FILE, index=False)
        st.session_state.forms.to_excel(FORMS_FILE, index=False)
        st.session_state.pwwf_boarding.to_excel(PWWF_BOARDING_FILE, index=False)
        save_room_capacities()
    except Exception as e:
        st.error(f"Error saving data: {e}")

def save_profile_data():
    """Save user profile data"""
    profiles_file = DATA_DIR / "profiles.xlsx"
    if profiles_file.exists():
        df = pd.read_excel(profiles_file)
    else:
        df = pd.DataFrame(columns=['username', 'officer_name', 'campus', 'designation', 'user_id'])
    
    if st.session_state.username in df['username'].values:
        df.loc[df['username'] == st.session_state.username, 'officer_name'] = st.session_state.officer_name
        df.loc[df['username'] == st.session_state.username, 'campus'] = st.session_state.campus
        df.loc[df['username'] == st.session_state.username, 'designation'] = st.session_state.designation
        df.loc[df['username'] == st.session_state.username, 'user_id'] = st.session_state.user_id
    else:
        new_row = pd.DataFrame([{
            'username': st.session_state.username,
            'officer_name': st.session_state.officer_name,
            'campus': st.session_state.campus,
            'designation': st.session_state.designation,
            'user_id': st.session_state.user_id
        }])
        df = pd.concat([df, new_row], ignore_index=True)
    
    df.to_excel(profiles_file, index=False)

# =========================================
# ROOM OCCUPANCY
# =========================================
def get_room_occupancy():
    if st.session_state.students.empty:
        return pd.DataFrame(columns=["Own","S.No","Room No","Total Capacity","Occupancy","Extra/Less"])
    
    df = st.session_state.students[st.session_state.students["Room No"].notna() & (st.session_state.students["Room No"] != "")]
    if df.empty:
        return pd.DataFrame(columns=["Own","S.No","Room No","Total Capacity","Occupancy","Extra/Less"])
    
    rooms = df.groupby("Room No").size().reset_index(name='Occupancy')
    
    # Use custom capacities if available, otherwise default to 4
    def get_capacity(room):
        if 'room_capacities' in st.session_state and room in st.session_state.room_capacities:
            return st.session_state.room_capacities[room]
        return 4
    
    rooms["Total Capacity"] = rooms["Room No"].apply(get_capacity)
    rooms["Extra/Less"] = rooms.apply(
        lambda row: f"+{row['Occupancy'] - row['Total Capacity']}" 
        if row['Occupancy'] - row['Total Capacity'] > 0 
        else str(row['Occupancy'] - row['Total Capacity']), 
        axis=1
    )
    
    rooms['sort_key'] = rooms['Room No'].apply(lambda x: ROOM_ORDER_DICT.get(x, 999))
    rooms = rooms.sort_values('sort_key').drop('sort_key', axis=1).reset_index(drop=True)
    rooms.insert(0, "S.No", range(1, len(rooms)+1))
    rooms.insert(0, "Own", False)
    
    return rooms

# =========================================
# DEFAULTERS MANAGEMENT
# =========================================
def update_defaulters():
    """Automatically update defaulters list based on students not in history"""
    if st.session_state.students.empty:
        return
    
    history_regs = set()
    if not st.session_state.history.empty:
        history_regs = set(st.session_state.history["Registration No"].astype(str).values)
    
    defaulters_data = []
    for _, student in st.session_state.students.iterrows():
        reg_no = str(student["Registration No"])
        if reg_no not in history_regs:
            defaulters_data.append({
                "Student Name": student["Name"],
                "Registration No": reg_no,
                "Room No": student["Room No"],
                "Status": student["Status"],
                "Semester": format_semester(student["Semester"]),
                "Amount": "0",
                "Defaulter Status": "Yes",
                "Remarks": "No payment record found"
            })
    
    if defaulters_data:
        st.session_state.defaulters = pd.DataFrame(defaulters_data)
    else:
        st.session_state.defaulters = pd.DataFrame(columns=DEFAULTERS_COLUMNS)

# =========================================
# PROCESS UPLOADED FILE
# =========================================
def get_processed_files():
    if PROCESSED_FILES_LOG.exists():
        with open(PROCESSED_FILES_LOG, 'r') as f:
            return set(f.read().splitlines())
    return set()

def mark_file_processed(filename):
    processed = get_processed_files()
    processed.add(filename)
    with open(PROCESSED_FILES_LOG, 'w') as f:
        f.write('\n'.join(processed))

def process_upload(files):
    new_students = []
    processed_files = get_processed_files()
    
    for file in files:
        if file.name in processed_files:
            st.warning(f"{file.name} already imported - skipping")
            continue
        
        try:
            if file.name.endswith(('.xlsx','.xls')):
                df = pd.read_excel(file)
                for _, row in df.iterrows():
                    reg = str(row.iloc[1]) if len(row)>1 else ""
                    if reg and reg not in st.session_state.students['Registration No'].astype(str).values:
                        new_students.append({
                            "Name": str(row.iloc[0]) if len(row)>0 else "",
                            "Registration No": reg,
                            "Room No": str(row.iloc[2]) if len(row)>2 else "",
                            "Status": str(row.iloc[3]) if len(row)>3 and str(row.iloc[3]) in ["Open", "PWWF"] else "Open",
                            "Contact No": str(row.iloc[4]) if len(row)>4 else "",
                            "Father Contact": str(row.iloc[5]) if len(row)>5 else "",
                            "Blood Group": str(row.iloc[6]) if len(row)>6 else "",
                        })
                mark_file_processed(file.name)
        except Exception as e:
            st.error(f"Error in {file.name}: {e}")
    
    if new_students:
        new_df = pd.DataFrame(new_students)
        new_df["Semester"] = new_df["Registration No"].apply(detect_semester)
        new_df["Program"] = new_df["Registration No"].apply(detect_program)
        new_df["SR#"] = range(len(st.session_state.students)+1, len(st.session_state.students)+len(new_df)+1)
        st.session_state.students = pd.concat([st.session_state.students, new_df], ignore_index=True)
        update_defaulters()
        return len(new_students)
    return 0

# =========================================
# PAYMENT SCANNER
# =========================================
def scan_for_new_payments():
    count = 0
    
    if not HOSTEL_DATA_PATH.exists():
        st.error(f"Directory does not exist: {HOSTEL_DATA_PATH}")
        return 0
    
    payment_files = list(HOSTEL_DATA_PATH.glob("Hostel_*.csv"))
    
    if not payment_files:
        st.info("No payment files found in D:/HostelData/")
        return 0
    
    for file in payment_files:
        try:
            df = pd.read_csv(file)
            expected_cols = ["Student Name", "Registration No", "Amount"]
            if all(col in df.columns for col in expected_cols):
                
                for _, row in df.iterrows():
                    reg = str(row["Registration No"])
                    amount = str(row["Amount"])
                    
                    payment_exists = False
                    if not st.session_state.history.empty:
                        existing = st.session_state.history[
                            (st.session_state.history["Registration No"] == reg) &
                            (st.session_state.history["Amount"] == amount)
                        ]
                        if not existing.empty:
                            payment_exists = True
                            continue
                    
                    if not payment_exists:
                        student_data = st.session_state.students[
                            st.session_state.students["Registration No"] == reg
                        ]
                        
                        if not student_data.empty:
                            room_no = student_data.iloc[0]["Room No"]
                            status = student_data.iloc[0]["Status"]
                            semester = student_data.iloc[0]["Semester"]
                            student_name = student_data.iloc[0]["Name"]
                        else:
                            room_no = row.get("Room No", "")
                            status = ""
                            semester = row.get("Semester", "")
                            student_name = row["Student Name"]
                        
                        payment_date = row.get("Date", datetime.datetime.now().date())
                        
                        new_payment = pd.DataFrame([{
                            "Date": str(payment_date),
                            "Time": datetime.datetime.now().strftime("%H:%M:%S"),
                            "Student Name": str(student_name),
                            "Registration No": str(reg),
                            "Room No": str(room_no),
                            "Status": str(status),
                            "Semester": str(semester),
                            "Amount": str(amount),
                            "Payment Method": str(row.get("Payment Method", "Challan")),
                            "Remarks": f"Imported from {file.name}"
                        }])
                        
                        if st.session_state.history.empty:
                            st.session_state.history = new_payment
                        else:
                            st.session_state.history = pd.concat([st.session_state.history, new_payment], ignore_index=True)
                        
                        count += 1
                
                if count > 0:
                    dest_path = HOSTEL_PROCESSED_PATH / file.name
                    file.rename(dest_path)
            
            else:
                st.warning(f"File {file.name} missing required columns")
                
        except Exception as e:
            st.error(f"Error processing {file.name}: {str(e)}")
    
    if count > 0:
        save_data()
        update_defaulters()
        st.success(f"Successfully imported {count} new payments!")
    
    return count

# =========================================
# PROFILE PAGE with Logout Button
# =========================================
def profile_page():
    st.markdown("<h1 style='color:#FFD700;'>Profile</h1>", unsafe_allow_html=True)
    
    # Logout Button
    if st.button('Logout', use_container_width=True):
        st.session_state.logged_in = False
        st.session_state.menu = 'Dashboard'
        st.session_state.username = None
        st.rerun()
    
    st.markdown("---")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("Profile Picture")
        if PROFILE_PIC_PATH.exists():
            try:
                profile_img = Image.open(PROFILE_PIC_PATH)
                st.image(profile_img, width=200)
            except:
                st.info("Profile picture: pp (1).png")
        else:
            st.info("Profile picture: pp (1).png not found")
    
    with col2:
        st.subheader("Personal Information")
        
        officer_name = st.text_input("Officer Name", value=st.session_state.get('officer_name', st.session_state.username))
        campus = st.selectbox("Campus", 
                              ["ISB", "LHR", "ATB", "WAH", "ATK", "VEH", "SWL"],
                              index=["ISB", "LHR", "ATB", "WAH", "ATK", "VEH", "SWL"].index(st.session_state.get('campus', 'ISB')))
        user_id = st.text_input("ID", value=st.session_state.get('user_id', f'EMP{hash(st.session_state.username) % 1000:03d}'))
        designation = st.text_input("Designation", value=st.session_state.get('designation', 'Officer'))
        
        if st.button("Save Profile Changes", use_container_width=True):
            st.session_state.officer_name = officer_name
            st.session_state.campus = campus
            st.session_state.user_id = user_id
            st.session_state.designation = designation
            save_profile_data()
            st.success("Profile updated successfully!")
        
        st.markdown("---")
        st.subheader("Change Password")
        
        with st.form("change_password_form"):
            current_pwd = st.text_input("Current Password", type="password")
            new_pwd = st.text_input("New Password", type="password")
            confirm_pwd = st.text_input("Confirm New Password", type="password")
            
            if st.form_submit_button("Change Password", use_container_width=True):
                if current_pwd and new_pwd and confirm_pwd:
                    users_file = DATA_DIR / "users.xlsx"
                    df = pd.read_excel(users_file)
                    user_row = df[df['username'] == st.session_state.username]
                    
                    if not user_row.empty and user_row.iloc[0]['password'] == hash_password(current_pwd):
                        if new_pwd == confirm_pwd:
                            df.loc[df['username'] == st.session_state.username, 'password'] = hash_password(new_pwd)
                            df.to_excel(users_file, index=False)
                            st.success("Password changed successfully!")
                        else:
                            st.error("New passwords do not match")
                    else:
                        st.error("Current password is incorrect")
                else:
                    st.warning("Please fill all fields")

# =========================================
# INITIALIZE SESSION STATE
# =========================================
def init_session_state():
    """Initialize all session state variables"""
    if 'logged_in' not in st.session_state:
        st.session_state.logged_in = False
    if 'show_forgot' not in st.session_state:
        st.session_state.show_forgot = False
    if 'menu' not in st.session_state:
        st.session_state.menu = 'Dashboard'
    if 'room_capacities' not in st.session_state:
        load_room_capacities()
    if 'username' not in st.session_state:
        st.session_state.username = None
    if 'officer_name' not in st.session_state:
        st.session_state.officer_name = None
    if 'campus' not in st.session_state:
        st.session_state.campus = 'ISB'
    if 'designation' not in st.session_state:
        st.session_state.designation = 'Officer'
    if 'user_id' not in st.session_state:
        st.session_state.user_id = None
    if 'students' not in st.session_state:
        st.session_state.students = pd.DataFrame()
    if 'history' not in st.session_state:
        st.session_state.history = pd.DataFrame()
    if 'defaulters' not in st.session_state:
        st.session_state.defaulters = pd.DataFrame()
    if 'forms' not in st.session_state:
        st.session_state.forms = pd.DataFrame()
    if 'pwwf_boarding' not in st.session_state:
        st.session_state.pwwf_boarding = pd.DataFrame()

# =========================================
# MAIN APP
# =========================================
def main():
    # Initialize all session state variables first
    init_session_state()
    
    # Remove deploy message on all pages
    remove_deploy_message()
    
    # Check if user is logged in
    if not st.session_state.logged_in:
        if st.session_state.show_forgot:
            forgot_password_page()
        else:
            login_page()
        return
    
    # Load data if not already loaded (check if students dataframe is empty as a proxy)
    if st.session_state.students.empty:
        load_data()
        update_defaulters()
    
    apply_custom_css()
    
    # Top navigation
    cols = st.columns([1,1,1,1,1,1,1,1])
    
    with cols[0]:
        if LOGO_PATH.exists():
            try:
                logo_image = Image.open(LOGO_PATH)
                st.image(logo_image, width=80)
            except:
                st.markdown("<span style='color:#FFD700; font-weight:bold;'>HMS</span>", unsafe_allow_html=True)
        else:
            st.markdown("<span style='color:#FFD700; font-weight:bold;'>HMS</span>", unsafe_allow_html=True)
    
    menu_items = ["Dashboard", "Students", "Upload", "History", "Forms", "PWWF Boarding", "Profile"]
    for i, item in enumerate(menu_items):
        with cols[i+1]:
            if st.button(item, use_container_width=True):
                st.session_state.menu = item
                st.rerun()
    
    st.markdown("---")
    
    menu = st.session_state.get('menu', 'Dashboard')
    
    # ========== DASHBOARD ==========
    if menu == "Dashboard":
        st.markdown("<h1 style='color:#FFD700;'>Dashboard</h1>", unsafe_allow_html=True)
        
        # Calculate metrics
        total_students = len(st.session_state.students)
        open_students = len(st.session_state.students[st.session_state.students["Status"]=="Open"]) if not st.session_state.students.empty else 0
        pwwf = len(st.session_state.students[st.session_state.students["Status"]=="PWWF"]) if not st.session_state.students.empty else 0
        rooms_used = len(st.session_state.students["Room No"].unique()) if not st.session_state.students.empty else 0
        
        # Calculate total capacity using room occupancy dataframe
        rooms_df = get_room_occupancy()
        total_capacity = rooms_df['Total Capacity'].sum() if not rooms_df.empty else 0
        
        # Display metrics in 6 columns
        cols = st.columns(6)
        
        cols[0].metric('Total Students', total_students)
        cols[1].metric('Open Merit', open_students)
        cols[2].metric('PWWF', pwwf)
        cols[3].metric('Rooms Used', rooms_used)
        cols[4].metric('Total Capacity', total_capacity)
        cols[5].metric('Occupancy %', '0%')
    
    # ========== STUDENTS ==========
    elif menu == "Students":
        st.markdown("<h1 style='color:#FFD700;'>Student Management</h1>", unsafe_allow_html=True)
        
        with st.expander("Room Overview", expanded=True):
            rooms_df = get_room_occupancy()
            
            if not rooms_df.empty:
                edited_rooms = st.data_editor(
                    rooms_df,
                    use_container_width=True,
                    key="room_editor",
                    column_config={
                        "Own": st.column_config.CheckboxColumn("Own", help="Select rooms"),
                        "S.No": st.column_config.NumberColumn("S.No", disabled=True),
                        "Room No": st.column_config.TextColumn("Room No", disabled=True),
                        "Total Capacity": st.column_config.NumberColumn("Capacity", min_value=1, max_value=10),
                        "Occupancy": st.column_config.NumberColumn("Occupancy", disabled=True),
                        "Extra/Less": st.column_config.TextColumn("Extra/Less", disabled=True)
                    }
                )
                
                col_save1, col_save2, col_save3 = st.columns([1, 1, 4])
                with col_save1:
                    if st.button("Save Capacity Changes", use_container_width=True, key="save_room_changes"):
                        # Create a mapping of room to new capacity
                        for _, row in edited_rooms.iterrows():
                            room_no = row["Room No"]
                            new_capacity = row["Total Capacity"]
                            
                            if 'room_capacities' not in st.session_state:
                                st.session_state.room_capacities = {}
                            
                            st.session_state.room_capacities[room_no] = new_capacity
                        
                        save_room_capacities()
                        st.success(f"Room capacities updated successfully for {len(edited_rooms)} rooms!")
                        st.rerun()
                
                with col_save2:
                    if st.button("Reset to Default", use_container_width=True):
                        if 'room_capacities' in st.session_state:
                            st.session_state.room_capacities = {}
                            save_room_capacities()
                        st.success("Reset to default capacities (4 per room)")
                        st.rerun()
                
                # Show current custom capacities if any
                if 'room_capacities' in st.session_state and st.session_state.room_capacities:
                    with st.expander("Current Custom Capacities", expanded=False):
                        custom_df = pd.DataFrame([
                            {"Room No": room, "Custom Capacity": cap} 
                            for room, cap in st.session_state.room_capacities.items()
                        ])
                        st.dataframe(custom_df, use_container_width=True)
                
                to_delete = edited_rooms[edited_rooms["Own"] == True]["Room No"].tolist()
                if to_delete and st.button("Delete Selected Rooms", use_container_width=True):
                    for room in to_delete:
                        st.session_state.students.loc[st.session_state.students["Room No"] == room, "Room No"] = ""
                    save_data()
                    update_defaulters()
                    st.success(f"Deleted {len(to_delete)} rooms")
                    st.rerun()
            else:
                st.info("No rooms with students yet")
        
        with st.expander("Manual Student Entry", expanded=False):
            st.markdown("<h3 style='color: #FFD700;'>Add New Student Manually</h3>", unsafe_allow_html=True)
            
            col1, col2 = st.columns(2)
            
            with col1:
                new_name = st.text_input("Student Name", key="new_name")
                new_reg = st.text_input("Registration No", key="new_reg")
                new_room = st.text_input("Room No", key="new_room")
                new_status = st.selectbox("Status", ["Open", "PWWF"], key="new_status", index=0)
            
            with col2:
                new_contact = st.text_input("Contact No", key="new_contact")
                new_father = st.text_input("Father Contact", key="new_father")
                new_blood = st.selectbox("Blood Group", ["A+","A-","B+","B-","AB+","AB-","O+","O-",""], key="new_blood")
            
            if st.button("Add Student", use_container_width=True):
                if new_name and new_reg:
                    if new_reg not in st.session_state.students["Registration No"].astype(str).values:
                        semester = detect_semester(new_reg)
                        program = detect_program(new_reg)
                        next_sr = len(st.session_state.students) + 1 if not st.session_state.students.empty else 1
                        
                        new_student = pd.DataFrame([{
                            "SR#": next_sr,
                            "Name": new_name,
                            "Registration No": new_reg,
                            "Room No": new_room,
                            "Status": new_status,
                            "Contact No": str(new_contact),
                            "Father Contact": str(new_father),
                            "Blood Group": new_blood,
                            "Semester": format_semester(semester),
                            "Program": program
                        }])
                        
                        if st.session_state.students.empty:
                            st.session_state.students = new_student
                        else:
                            st.session_state.students = pd.concat([st.session_state.students, new_student], ignore_index=True)
                        
                        update_defaulters()
                        st.success(f"Student {new_name} added successfully!")
                        st.rerun()
                    else:
                        st.error("Registration number already exists")
                else:
                    st.warning("Name and Registration No are required")
        
        st.markdown("<h2 style='color: #FFD700;'>Student List</h2>", unsafe_allow_html=True)
        
        if not st.session_state.students.empty:
            st.session_state.students["Status"] = st.session_state.students["Status"].fillna("Open")
            st.session_state.students.loc[~st.session_state.students["Status"].isin(["Open", "PWWF"]), "Status"] = "Open"
            
            st.session_state.students['sort_key'] = st.session_state.students['Room No'].apply(
                lambda x: ROOM_ORDER_DICT.get(x, 999) if pd.notna(x) and x and x != "" else 999
            )
            st.session_state.students = st.session_state.students.sort_values('sort_key').drop('sort_key', axis=1).reset_index(drop=True)
            st.session_state.students["SR#"] = range(1, len(st.session_state.students)+1)
        
        search = st.text_input("Search by Name or Registration No", placeholder="Type to search...")
        
        display_df = st.session_state.students.copy()
        if not display_df.empty:
            display_df["Contact No"] = display_df["Contact No"].astype(str)
            display_df["Father Contact"] = display_df["Father Contact"].astype(str)
            display_df["Room No"] = display_df["Room No"].astype(str)
            display_df["Status"] = display_df["Status"].astype(str)
            display_df["Semester"] = display_df["Semester"].astype(str)
            display_df["Program"] = display_df["Program"].astype(str)
            
            # Apply semester formatting to displayed data
            display_df['Semester'] = display_df['Semester'].apply(format_semester)
        
        if search:
            display_df = display_df[
                display_df["Name"].str.contains(search, case=False, na=False) |
                display_df["Registration No"].str.contains(search, case=False, na=False)
            ]
        
        edited_df = st.data_editor(
            display_df,
            num_rows="dynamic",
            use_container_width=True,
            key="student_editor",
            column_config={
                "SR#": st.column_config.NumberColumn("SR#", step=1, disabled=True),
                "Name": st.column_config.TextColumn("Name", required=True),
                "Registration No": st.column_config.TextColumn("Reg No", required=True),
                "Room No": st.column_config.TextColumn("Room No"),
                "Status": st.column_config.SelectboxColumn("Status", options=["Open", "PWWF"], required=True, default="Open"),
                "Contact No": st.column_config.TextColumn("Contact"),
                "Father Contact": st.column_config.TextColumn("Father"),
                "Blood Group": st.column_config.SelectboxColumn("Blood", options=["A+","A-","B+","B-","AB+","AB-","O+","O-",""]),
                "Semester": st.column_config.TextColumn("Semester", disabled=True),
                "Program": st.column_config.TextColumn("Program", disabled=True)
            }
        )
        
        col1, col2, col3 = st.columns([1,1,2])
        with col1:
            if st.button("Save Student Changes", use_container_width=True):
                if not edited_df.empty:
                    for idx in edited_df.index:
                        if idx < len(st.session_state.students):
                            for col in edited_df.columns:
                                st.session_state.students.loc[idx, col] = edited_df.loc[idx, col]
                    save_data()
                    update_defaulters()
                    st.success("Student list saved successfully!")
                    st.rerun()
        
        with col2:
            if not st.session_state.students.empty:
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as w:
                    st.session_state.students.to_excel(w, index=False, sheet_name='Students')
                st.download_button(
                    "Download Excel",
                    output.getvalue(),
                    f"students_{datetime.datetime.now().strftime('%Y%m%d')}.xlsx",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
    
    # ========== UPLOAD ==========
    elif menu == "Upload":
        st.markdown("<h1 style='color:#FFD700;'>Upload Students</h1>", unsafe_allow_html=True)
        st.info("Upload files containing: Name, Registration No, Room No, Status, Contact No, Father Contact, Blood Group")
        
        files = st.file_uploader(
            "Choose files", 
            type=["xlsx","xls","pdf","png","jpg","jpeg"], 
            accept_multiple_files=True,
            help="Upload Excel, PDF, or Image files"
        )
        
        if files and st.button("Import All Files", use_container_width=True):
            count = process_upload(files)
            if count:
                st.success(f"Imported {count} new students")
                st.rerun()
            else:
                st.info("No new students to import")
    
    # ========== HISTORY ==========
    elif menu == "History":
        st.markdown("<h1 style='color:#FFD700;'>Payment History & Defaulters</h1>", unsafe_allow_html=True)
        
        # D: Drive Integration
        with st.expander("D: Drive Integration", expanded=True):
            col1, col2, col3 = st.columns(3)
            with col1:
                st.write(f"**Source:** {HOSTEL_DATA_PATH}")
                if HOSTEL_DATA_PATH.exists():
                    st.success("Directory exists")
                else:
                    st.error("Directory does not exist!")
            
            with col2:
                st.write(f"**Archive:** {HOSTEL_PROCESSED_PATH}")
                if HOSTEL_PROCESSED_PATH.exists():
                    st.success("Directory exists")
                else:
                    st.error("Directory does not exist!")
            
            with col3:
                if st.button("Scan for Payments", use_container_width=True):
                    with st.spinner("Scanning for payments..."):
                        count = scan_for_new_payments()
                        if count:
                            st.rerun()
            
            st.markdown("---")
            pending = list(HOSTEL_DATA_PATH.glob("Hostel_*.csv"))
            st.write(f"**Pending Files:** {len(pending)}")
            if pending:
                for file in pending:
                    file_size = file.stat().st_size / 1024
                    modified = datetime.datetime.fromtimestamp(file.stat().st_mtime)
                    st.text(f"• {file.name} ({file_size:.1f} KB) - {modified.strftime('%Y-%m-%d %H:%M')}")
            else:
                st.info("No pending payment files")
        
        # Defaulters Section
        with st.expander('Defaulters', expanded=False):
            paid_regs = set(st.session_state.history['Registration No']) if not st.session_state.history.empty else set()
            defaulters = st.session_state.students[
                ~st.session_state.students['Registration No'].isin(paid_regs)
            ] if not st.session_state.students.empty else pd.DataFrame()
            
            if not defaulters.empty:
                # Apply semester formatting
                if 'Semester' in defaulters.columns:
                    defaulters['Semester'] = defaulters['Semester'].apply(format_semester)
                st.dataframe(defaulters, use_container_width=True)
            else:
                st.success('No defaulters found')
        
        # Filter Section for Payment Records
        st.markdown("---")
        st.markdown("<h2 style='color: #FFD700;'>Filter Payment Records</h2>", unsafe_allow_html=True)
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            if not st.session_state.history.empty and 'Semester' in st.session_state.history.columns:
                semesters = ['All'] + sorted(st.session_state.history['Semester'].unique().tolist())
                selected_semester = st.selectbox("Filter by Semester", semesters)
            else:
                selected_semester = 'All'
                st.selectbox("Filter by Semester", ['All'], disabled=True)
        
        with col2:
            if not st.session_state.history.empty and 'Date' in st.session_state.history.columns:
                months = ['All', 'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
                selected_month = st.selectbox("Filter by Month", months)
            else:
                selected_month = 'All'
                st.selectbox("Filter by Month", ['All'], disabled=True)
        
        with col3:
            search_type = st.selectbox("Search by", ["Name", "Registration No", "Room No"])
            search_value = st.text_input(f"Enter {search_type}", placeholder=f"Type to search by {search_type}...")
        
        filtered_history = st.session_state.history.copy()
        
        if not filtered_history.empty:
            for col in filtered_history.columns:
                filtered_history[col] = filtered_history[col].astype(str)
            
            if selected_semester != 'All':
                filtered_history = filtered_history[filtered_history['Semester'] == selected_semester]
            
            if selected_month != 'All' and 'Date' in filtered_history.columns:
                month_map = {'Jan':'01', 'Feb':'02', 'Mar':'03', 'Apr':'04', 'May':'05', 'Jun':'06',
                            'Jul':'07', 'Aug':'08', 'Sep':'09', 'Oct':'10', 'Nov':'11', 'Dec':'12'}
                month_num = month_map.get(selected_month, '')
                if month_num:
                    filtered_history = filtered_history[
                        filtered_history['Date'].str.contains(f"-{month_num}-", na=False)
                    ]
            
            if search_value:
                if search_type == "Name":
                    filtered_history = filtered_history[filtered_history['Student Name'].str.contains(search_value, case=False, na=False)]
                elif search_type == "Registration No":
                    filtered_history = filtered_history[filtered_history['Registration No'].str.contains(search_value, case=False, na=False)]
                elif search_type == "Room No":
                    filtered_history = filtered_history[filtered_history['Room No'].str.contains(search_value, case=False, na=False)]
        
        st.markdown("<h2 style='color: #FFD700;'>Payment Records</h2>", unsafe_allow_html=True)
        st.write(f"**Showing {len(filtered_history)} records**")
        
        edited = st.data_editor(
            filtered_history,
            num_rows="dynamic",
            use_container_width=True,
            key="history_editor",
            column_config={
                "Date": st.column_config.TextColumn("Date"),
                "Time": st.column_config.TextColumn("Time"),
                "Student Name": st.column_config.TextColumn("Student Name"),
                "Registration No": st.column_config.TextColumn("Registration No"),
                "Room No": st.column_config.TextColumn("Room No"),
                "Status": st.column_config.SelectboxColumn("Status", options=["Open", "PWWF", ""]),
                "Semester": st.column_config.TextColumn("Semester"),
                "Amount": st.column_config.TextColumn("Amount"),
                "Payment Method": st.column_config.SelectboxColumn("Payment Method", options=["Cash", "Bank Transfer", "Cheque", "Online", "Challan", "Other"]),
                "Remarks": st.column_config.TextColumn("Remarks")
            }
        )
        
        col1, col2 = st.columns(2)
        with col1:
            if st.button("Save History Changes", use_container_width=True):
                st.session_state.history = edited
                save_data()
                update_defaulters()
                st.success("History saved successfully!")
                st.rerun()
        
        with col2:
            if not st.session_state.history.empty:
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as w:
                    st.session_state.history.to_excel(w, index=False, sheet_name='History')
                st.download_button(
                    "Download History as Excel",
                    output.getvalue(),
                    f"history_{datetime.datetime.now().strftime('%Y%m%d')}.xlsx",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
    
    # ========== FORMS ==========
    elif menu == "Forms":
        st.markdown("<h1 style='color:#FFD700;'>Forms Management</h1>", unsafe_allow_html=True)
        
        if st.session_state.forms.empty and not st.session_state.students.empty:
            data = []
            for _, row in st.session_state.students.iterrows():
                data.append({
                    "Student Name": row["Name"],
                    "Registration No": row["Registration No"],
                    "Room No": row["Room No"],
                    "Status": row["Status"],
                    "Semester": format_semester(row["Semester"]),
                    "Admission Form": "Not Submitted",
                    "PWWF Form": "Not Submitted" if row["Status"] == "PWWF" else "N/A",
                    "Consent Form": "Not Submitted",
                    "Amount": "N/A" if row["Status"] == "PWWF" else "0"
                })
            st.session_state.forms = pd.DataFrame(data)
            save_data()
        
        forms_display = st.session_state.forms.copy()
        if not forms_display.empty:
            for col in forms_display.columns:
                forms_display[col] = forms_display[col].astype(str)
            
            # Apply semester formatting
            if 'Semester' in forms_display.columns:
                forms_display['Semester'] = forms_display['Semester'].apply(format_semester)
        
        tabs = st.tabs(["Admission Form", "PWWF Verification", "Consent Form"])
        
        with tabs[0]:
            st.markdown("<h2 style='color: #FFD700;'>Admission Form</h2>", unsafe_allow_html=True)
            if not forms_display.empty:
                df = forms_display[["Student Name","Registration No","Room No","Status","Semester","Admission Form"]].copy()
                df.insert(0, "SR#", range(1, len(df)+1))
                
                # Make SR# editable by setting disabled=False
                edited = st.data_editor(
                    df, 
                    use_container_width=True, 
                    key="admission_editor",
                    column_config={
                        "SR#": st.column_config.NumberColumn("SR#", step=1, disabled=False),
                        "Student Name": st.column_config.TextColumn("Student Name", disabled=True),
                        "Registration No": st.column_config.TextColumn("Registration No", disabled=True),
                        "Room No": st.column_config.TextColumn("Room No", disabled=True),
                        "Status": st.column_config.TextColumn("Status", disabled=True),
                        "Semester": st.column_config.TextColumn("Semester", disabled=True),
                        "Admission Form": st.column_config.SelectboxColumn("Admission Form", options=["Submitted", "Not Submitted"])
                    }
                )
                
                if st.button("Save Admission Form Changes", use_container_width=True, key="save_admission"):
                    for _, row in edited.iterrows():
                        mask = st.session_state.forms["Registration No"] == row["Registration No"]
                        if mask.any():
                            st.session_state.forms.loc[mask, "Admission Form"] = row["Admission Form"]
                    save_data()
                    st.success("Admission form changes saved!")
                    st.rerun()
        
        with tabs[1]:
            st.markdown("<h2 style='color: #FFD700;'>PWWF Verification Form</h2>", unsafe_allow_html=True)
            pwwf = forms_display[forms_display["Status"] == "PWWF"]
            if not pwwf.empty:
                df = pwwf[["Student Name","Registration No","Room No","Status","Semester","PWWF Form"]].copy()
                df.insert(0, "SR#", range(1, len(df)+1))
                
                edited = st.data_editor(
                    df, 
                    use_container_width=True, 
                    key="pwwf_editor",
                    column_config={
                        "SR#": st.column_config.NumberColumn("SR#", step=1, disabled=True),
                        "Student Name": st.column_config.TextColumn("Student Name", disabled=True),
                        "Registration No": st.column_config.TextColumn("Registration No", disabled=True),
                        "Room No": st.column_config.TextColumn("Room No", disabled=True),
                        "Status": st.column_config.TextColumn("Status", disabled=True),
                        "Semester": st.column_config.TextColumn("Semester", disabled=True),
                        "PWWF Form": st.column_config.SelectboxColumn("PWWF Form", options=["Submitted", "Not Submitted"])
                    }
                )
                
                if st.button("Save PWWF Form Changes", use_container_width=True, key="save_pwwf"):
                    for _, row in edited.iterrows():
                        mask = st.session_state.forms["Registration No"] == row["Registration No"]
                        if mask.any():
                            st.session_state.forms.loc[mask, "PWWF Form"] = row["PWWF Form"]
                    save_data()
                    st.success("PWWF form changes saved!")
                    st.rerun()
            else:
                st.info("No PWWF students found")
        
        with tabs[2]:
            st.markdown("<h2 style='color: #FFD700;'>Consent Form</h2>", unsafe_allow_html=True)
            if not forms_display.empty:
                df = forms_display[["Student Name","Registration No","Room No","Status","Semester","Consent Form","Amount"]].copy()
                df.insert(0, "SR#", range(1, len(df)+1))
                
                edited = st.data_editor(
                    df, 
                    use_container_width=True, 
                    key="consent_editor",
                    column_config={
                        "SR#": st.column_config.NumberColumn("SR#", step=1, disabled=True),
                        "Student Name": st.column_config.TextColumn("Student Name", disabled=True),
                        "Registration No": st.column_config.TextColumn("Registration No", disabled=True),
                        "Room No": st.column_config.TextColumn("Room No", disabled=True),
                        "Status": st.column_config.TextColumn("Status", disabled=True),
                        "Semester": st.column_config.TextColumn("Semester", disabled=True),
                        "Consent Form": st.column_config.SelectboxColumn("Consent Form", options=["Submitted", "Not Submitted"]),
                        "Amount": st.column_config.TextColumn("Amount")
                    }
                )
                
                if st.button("Save Consent Form Changes", use_container_width=True, key="save_consent"):
                    for _, row in edited.iterrows():
                        mask = st.session_state.forms["Registration No"] == row["Registration No"]
                        if mask.any():
                            st.session_state.forms.loc[mask, "Consent Form"] = row["Consent Form"]
                            if row["Status"] == "Open":
                                st.session_state.forms.loc[mask, "Amount"] = row["Amount"]
                    save_data()
                    st.success("Consent form changes saved!")
                    st.rerun()
        
        if not st.session_state.forms.empty:
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as w:
                st.session_state.forms.to_excel(w, index=False, sheet_name='Forms')
            st.download_button(
                "Download Forms as Excel",
                output.getvalue(),
                f"forms_{datetime.datetime.now().strftime('%Y%m%d')}.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
    
    # ========== PWWF BOARDING ==========
    elif menu == "PWWF Boarding":
        st.markdown("<h1 style='color:#FFD700;'>PWWF Boarding</h1>", unsafe_allow_html=True)
        
        # Upload Section
        st.markdown("<h2 style='color: #FFD700;'>Upload PWWF Boarding Data</h2>", unsafe_allow_html=True)
        uploaded_file = st.file_uploader(
            "Choose Excel file",
            type=["xlsx", "xls"],
            key="pwwf_upload",
            help="Upload Excel file with columns: Student Name, Registration No, Semester, Amount, Paying Date"
        )
        
        if uploaded_file is not None:
            try:
                df_upload = pd.read_excel(uploaded_file)
                # Convert semester to ordinal if it's a number
                if 'Semester' in df_upload.columns:
                    df_upload['Semester'] = df_upload['Semester'].apply(format_semester)
                # Add SR# column
                df_upload.insert(0, "SR#", range(1, len(df_upload)+1))
                st.session_state.pwwf_boarding = df_upload
                save_data()
                st.success("PWWF Boarding data uploaded successfully!")
                st.rerun()
            except Exception as e:
                st.error(f"Error uploading file: {e}")
        
        st.markdown("---")
        
        # Display and Edit PWWF Boarding Data
        st.markdown("<h2 style='color: #FFD700;'>PWWF Boarding List</h2>", unsafe_allow_html=True)
        
        if not st.session_state.pwwf_boarding.empty:
            # Convert semester to ordinal if needed
            if 'Semester' in st.session_state.pwwf_boarding.columns:
                st.session_state.pwwf_boarding['Semester'] = st.session_state.pwwf_boarding['Semester'].apply(format_semester)
            
            edited_pwwf = st.data_editor(
                st.session_state.pwwf_boarding,
                num_rows="dynamic",
                use_container_width=True,
                key="pwwf_boarding_editor",
                column_config={
                    "SR#": st.column_config.NumberColumn("SR#", disabled=True),
                    "Student Name": st.column_config.TextColumn("Student Name", required=True),
                    "Registration No": st.column_config.TextColumn("Registration No", required=True),
                    "Semester": st.column_config.TextColumn("Semester"),
                    "Amount": st.column_config.TextColumn("Amount"),
                    "Paying Date": st.column_config.DateColumn("Paying Date", format="YYYY-MM-DD")
                }
            )
            
            col1, col2 = st.columns(2)
            with col1:
                if st.button("Save PWWF Boarding Changes", use_container_width=True):
                    # Update SR# after deletions
                    edited_pwwf["SR#"] = range(1, len(edited_pwwf)+1)
                    st.session_state.pwwf_boarding = edited_pwwf
                    save_data()
                    st.success("PWWF Boarding data saved successfully!")
                    st.rerun()
            
            with col2:
                # Download
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as w:
                    st.session_state.pwwf_boarding.to_excel(w, index=False, sheet_name='PWWF Boarding')
                st.download_button(
                    "Download PWWF Boarding Data",
                    output.getvalue(),
                    f"pwwf_boarding_{datetime.datetime.now().strftime('%Y%m%d')}.xlsx",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
        else:
            st.info("No PWWF Boarding data available. Please upload an Excel file.")
    
    # ========== PROFILE ==========
    elif menu == "Profile":
        profile_page()
    
    # =========================================
    # FOOTER
    # =========================================
    st.markdown("---")
    st.markdown("📍 1.5 KM Defence Road, Off Raiwind Road, Lahore Pakistan")

# =========================================
# INITIALIZE USERS
# =========================================
def init_users():
    users_file = DATA_DIR / "users.xlsx"
    if not users_file.exists():
        df = pd.DataFrame([{
            'username': 'admin',
            'password': hash_password('admin123'),
            'name': 'Administrator'
        }])
        df.to_excel(users_file, index=False)

init_users()

# =========================================
# RUN
# =========================================
if __name__ == "__main__":
    main()
