import streamlit as st
import subprocess
import sys
import os
import time
from datetime import datetime
import glob
import traceback
import uuid

# Page configuration
st.set_page_config(
    page_title="Milestone Report Generator",
    page_icon="📈",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Enhanced Custom CSS for a professional, modern interface
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Roboto:wght@300;400;500;700&display=swap');

    .stApp {
        background: linear-gradient(145deg, #2c3e50 0%, #4a69bd 100%);
        font-family: 'Roboto', sans-serif;
        color: #333333;
    }

    /* Hide Streamlit default elements */
    #MainMenu, footer, div[data-testid="stToolbar"], .stDeployButton, div[data-testid="stDecoration"] {
        display: none;
    }

    .main-container {
        background: linear-gradient(145deg, #2c3e50 0%, #4a69bd 100%);
        font-family: 'Roboto', sans-serif;
        color: #333333;
        padding: 2rem;
    }

    .main-title {
        text-align: center;
        font-size: 2.5rem;
        font-weight: 700;
        color: #ffffff;
        margin-bottom: 0.5rem;
        letter-spacing: -0.5px;
    }

    .subtitle {
        text-align: center;
        font-size: 1.1rem;
        color: #d3d3d3;
        margin-bottom: 2rem;
        font-weight: 400;
    }

    .chat-message {
        padding: 1.2rem;
        border-radius: 12px;
        margin-bottom: 1rem;
        display: flex;
        align-items: flex-start;
        transition: all 0.3s ease;
        background: #ffffff;
        box-shadow: 0 2px 8px rgba(0, 0, 0, 0.05);
    }

    .chat-message.bot {
        background: #f6faff;
        margin-right: 1rem;
        border-left: 3px solid #4a69bd;
    }

    .chat-message.user {
        background: #f0fdf4;
        margin-left: 1rem;
        flex-direction: row-reverse;
        border-right: 3px solid #2ca02c;
    }

    .chat-message .avatar {
        width: 2.8rem;
        height: 2.8rem;
        border-radius: 50%;
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 1.3rem;
        margin: 0 0.8rem;
        background: #4a69bd;
        color: white;
    }

    .user .avatar {
        background: #2ca02c;
    }

    .chat-message .message {
        flex: 1;
        font-size: 1rem;
        line-height: 1.5;
        color: #333333;
    }

    .project-selection {
        background: rgba(255, 255, 255, 0.1);
        border-radius: 12px;
        padding: 2rem;
        margin: 1.5rem 0;
        box-shadow: 0 4px 16px rgba(0, 0, 0, 0.05);
    }

    .project-selection h3 {
        text-align: center;
        color: #ffffff;
        font-size: 1.6rem;
        font-weight: 600;
        margin-bottom: 1.5rem;
    }

    .project-buttons {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
        gap: 1rem;
    }

    .stButton > button {
        background: linear-gradient(145deg, #6b7280, #4a69bd) !important;
        color: #ffffff !important;
        border: none !important;
        padding: 1rem !important;
        border-radius: 10px !important;
        font-size: 1rem !important;
        font-weight: 500 !important;
        transition: all 0.3s ease !important;
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.1) !important;
        width: 100% !important;
    }

    .stButton > button:hover {
        transform: translateY(-2px) !important;
        box-shadow: 0 6px 16px rgba(0, 0, 0, 0.15) !important;
        background: linear-gradient(145deg, #4a69bd, #6b7280) !important;
    }

    .status-container, .success-container, .error-container {
        border-radius: 12px;
        padding: 1.5rem;
        margin: 1.5rem 0;
        text-align: center;
        box-shadow: 0 4px 16px rgba(0, 0, 0, 0.05);
    }

    .status-container {
        background: #fef9e7;
        border: 1px solid #facc15;
    }

    .success-container {
        background: #f0fdf4;
        border: 1px solid #2ca02c;
    }

    .error-container {
        background: #fef2f2;
        border: 1px solid #dc2626;
    }

    .status-container h4, .success-container h4, .error-container h4 {
        font-size: 1.3rem;
        font-weight: 600;
        margin-bottom: 0.8rem;
    }

    .status-container p, .success-container p, .error-container p {
        font-size: 1rem;
        margin-bottom: 0;
    }

    .stDownloadButton > button {
        background: linear-gradient(145deg, #2ca02c, #22c55e) !important;
        color: white !important;
        border: none !important;
        padding: 1.2rem !important;
        border-radius: 10px !important;
        font-size: 1.1rem !important;
        font-weight: 500 !important;
        transition: all 0.3s ease !important;
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.1) !important;
        width: 100% !important;
    }

    .stDownloadButton > button:hover {
        transform: translateY(-2px) !important;
        box-shadow: 0 6px 16px rgba(0, 0, 0, 0.15) !important;
        background: linear-gradient(145deg, #22c55e, #2ca02c) !important;
    }

    .stProgress > div > div > div > div {
        background: #4a69bd !important;
        border-radius: 8px !important;
    }

    .stProgress > div > div > div {
        background-color: #e5e7eb !important;
        border-radius: 8px !important;
        height: 12px !important;
    }

    .footer {
        text-align: center;
        color: rgba(255, 255, 255, 0.9);
        font-size: 0.9rem;
        margin-top: 2rem;
        padding: 1.5rem;
        background: rgba(0, 0, 0, 0.05);
        border-radius: 12px;
    }

    hr {
        border: none;
        height: 1px;
        background: linear-gradient(145deg, #4a69bd, #6b7280);
        margin: 1.5rem 0;
    }

    ::-webkit-scrollbar {
        width: 6px;
    }

    ::-webkit-scrollbar-track {
        background: rgba(255, 255, 255, 0.1);
        border-radius: 8px;
    }

    ::-webkit-scrollbar-thumb {
        background: #4a69bd;
        border-radius: 8px;
    }

    @media (max-width: 768px) {
        .main-title {
            font-size: 1.8rem;
        }

        .main-container {
            margin: 1rem;
            padding: 1.5rem;
        }

        .project-buttons {
            grid-template-columns: 1fr;
        }
    }

    .debug-info {
        background: #f8fafc;
        border: 1px solid #e2e8f0;
        border-radius: 8px;
        padding: 1rem;
        margin: 1rem 0;
        font-family: 'Courier New', monospace;
        font-size: 0.85rem;
        color: #1f2937;
    }
</style>
""", unsafe_allow_html=True)

# Initialize session state
if 'messages' not in st.session_state:
    st.session_state.messages = []
    st.session_state.stage = 'welcome'
    st.session_state.selected_project = None
    st.session_state.report_file = None

# Project configurations
PROJECTS = {
    'Veridia': {
        'script': 'veridia.py',
        'display_name': 'VERIDIA',
        'patterns': [
            'Time_Delivery_Milestones_Report_*.xlsx',
            '*Veridia*.xlsx',
            'Veridia_*.xlsx',
            '*veridia*.xlsx'
        ]
    },
    'Eligo': {
        'script': 'eligo.py',
        'display_name': 'ELIGO',
        'patterns': [
            '*Eligo*.xlsx',
            'Eligo_*.xlsx',
            '*eligo*.xlsx'
        ]
    },
    'EWS-LIG': {
        'script': 'ews-lig.py',
        'display_name': 'EWS-LIG',
        'patterns': [
            '*EWS*LIG*.xlsx',
            '*EWS-LIG*.xlsx',
            'EWS_LIG_*.xlsx',
            '*ews*lig*.xlsx'
        ]
    },
    'WaveCityClub': {
        'script': 'wavecityclub.py',
        'display_name': 'WAVECITY CLUB',
        'patterns': [
            'Wave_City_Club_Report_*.xlsx',
            '*WaveCityClub*.xlsx',
            '*Wave*City*Club*.xlsx',
            '*wavecityclub*.xlsx'
        ]
    },
    'Eden': {
        'script': 'eden.py',
        'display_name': 'EDEN',
        'patterns': [
            'Eden_KRA_Milestone_Report_*.xlsx',
            '*Eden*.xlsx',
            'Eden_*.xlsx',
            '*eden*.xlsx'
        ]
    }
}

def add_message(role, content):
    st.session_state.messages.append({
        'role': role,
        'content': content,
        'timestamp': datetime.now()
    })

def display_chat_message(message):
    role = message['role']
    content = message['content']
    
    if role == 'bot':
        st.markdown(f"""
        <div class="chat-message bot">
            <div class="avatar">🤖</div>
            <div class="message">{content}</div>
        </div>
        """, unsafe_allow_html=True)
    else:
        st.markdown(f"""
        <div class="chat-message user">
            <div class="avatar">👤</div>
            <div class="message">{content}</div>
        </div>
        """, unsafe_allow_html=True)

def find_generated_file(project_config, project_name):
    patterns = project_config['patterns']
    files_before = set(glob.glob("*.xlsx"))
    
    for pattern in patterns:
        matches = glob.glob(pattern)
        if matches:
            latest_file = max(matches, key=os.path.getctime)
            file_time = os.path.getctime(latest_file)
            if (time.time() - file_time) < 600:
                return latest_file
    
    files_after = set(glob.glob("*.xlsx"))
    new_files = files_after - files_before
    
    if new_files:
        return max(new_files, key=os.path.getctime)
    
    return None

def run_project_script(project_name):
    try:
        project_config = PROJECTS[project_name]
        script_path = project_config['script']
        
        if not os.path.exists(script_path):
            available_files = [f for f in os.listdir('.') if f.endswith('.py')]
            return False, f"Script file '{script_path}' not found. Available Python files: {available_files}"

        files_before = set(glob.glob("*.xlsx"))
        timeout_duration = 600 if project_name == 'Veridia' else 300
        env = os.environ.copy()
        env['PYTHONUNBUFFERED'] = '1'
        env['MPLBACKEND'] = 'Agg'

        result = subprocess.run(
            [sys.executable, script_path],
            capture_output=True,
            text=True,
            timeout=timeout_duration,
            cwd=os.getcwd(),
            env=env
        )

        if result.returncode != 0:
            return False, f"Script failed with return code {result.returncode}. Error: {result.stderr}"

        generated_file = find_generated_file(project_config, project_name)
        if generated_file and os.path.exists(generated_file):
            return True, generated_file
        
        return False, f"No report file found. Patterns: {project_config['patterns']}\nOutput: {result.stdout}\nError: {result.stderr}"

    except subprocess.TimeoutExpired:
        return False, f"Script timed out after {timeout_duration//60} minutes."
    except Exception as e:
        return False, f"Error: {str(e)}\nTraceback: {traceback.format_exc()}"

def main():
    st.markdown('<div class="main-container">', unsafe_allow_html=True)

    st.markdown("""
    <div class="main-title">Milestone Report Generator</div>
    <div class="subtitle">Professional report generation for project milestones</div>
    """, unsafe_allow_html=True)
    st.markdown("---")

    if not st.session_state.messages:
        add_message('bot', "Welcome to the Milestone Report Generator!")
        add_message('bot', "Please select a project to generate its milestone report.")

    for msg in st.session_state.messages:
        display_chat_message(msg)

    if st.session_state.stage == 'welcome' and not st.session_state.selected_project:
        st.markdown('<div class="project-selection"><h3>Select Project</h3></div>', unsafe_allow_html=True)
        
        st.markdown('<div class="project-buttons">', unsafe_allow_html=True)
        for key, info in PROJECTS.items():
            if st.button(f"{info['display_name']}", key=str(uuid.uuid4())):
                st.session_state.selected_project = key
                st.session_state.stage = 'processing'
                add_message('user', f"Generating report for {info['display_name']}.")
                add_message('bot', f"Generating {info['display_name']} report...")
                st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)

    elif st.session_state.stage == 'processing':
        proj = st.session_state.selected_project
        info = PROJECTS[proj]
        
        st.markdown(f"""
        <div class="status-container">
            <h4>Processing {info['display_name']}...</h4>
            <p>Generating your report. This may take a few minutes.</p>
        </div>
        """, unsafe_allow_html=True)

        progress_bar = st.progress(0)
        status_text = st.empty()
        progress_steps = [
            (0.2, "Initializing..."),
            (0.4, "Loading data..."),
            (0.6, "Processing calculations..."),
            (0.8, "Generating report..."),
            (1.0, "Finalizing...")
        ]
        
        for progress, step_text in progress_steps:
            status_text.text(step_text)
            progress_bar.progress(progress)
            time.sleep(0.5)
        
        progress_bar.empty()
        status_text.empty()
        
        with st.expander("Debug Information"):
            success, result = run_project_script(proj)
        
        if success:
            st.session_state.report_file = result
            st.session_state.stage = 'completed'
            add_message('bot', f"Report for {info['display_name']} generated successfully!")
        else:
            st.session_state.stage = 'error'
            st.session_state.error_message = result
            add_message('bot', f"Error generating {info['display_name']} report.")
        
        st.rerun()

    elif st.session_state.stage == 'completed':
        proj = st.session_state.selected_project
        info = PROJECTS[proj]
        
        st.markdown(f"""
        <div class="success-container">
            <h4>Report Generated</h4>
            <p>Your {info['display_name']} report is ready for download.</p>
        </div>
        """, unsafe_allow_html=True)

        if st.session_state.report_file and os.path.exists(st.session_state.report_file):
            file_size = os.path.getsize(st.session_state.report_file)
            file_size_mb = file_size / (1024 * 1024)
            st.info(f"File: {os.path.basename(st.session_state.report_file)} ({file_size_mb:.2f} MB)")
            
            with open(st.session_state.report_file, "rb") as f:
                file_data = f.read()
            
            st.download_button(
                label=f"Download {info['display_name']} Report",
                data=file_data,
                file_name=os.path.basename(st.session_state.report_file),
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        else:
            st.error("Report file not found.")

        if st.button("Generate Another Report", use_container_width=True):
            st.session_state.messages = []
            st.session_state.stage = 'welcome'
            st.session_state.selected_project = None
            st.session_state.report_file = None
            st.rerun()

    elif st.session_state.stage == 'error':
        proj = st.session_state.selected_project or ""
        info = PROJECTS.get(proj, {'display_name': 'Unknown'})
        
        st.markdown(f"""
        <div class="error-container">
            <h4>Error Generating Report</h4>
            <p>An issue occurred while generating the {info['display_name']} report.</p>
        </div>
        """, unsafe_allow_html=True)

        if hasattr(st.session_state, 'error_message'):
            with st.expander("Error Details", expanded=True):
                st.markdown(f"```\n{st.session_state.error_message}\n```")

        with st.expander("Troubleshooting Tips"):
            st.markdown(f"""
            **Common Issues:**
            - Missing script file: Ensure `{info.get('script', 'unknown.py')}` exists.
            - Missing dependencies: Check Python package installations.
            - Data file issues: Verify input data files are accessible.
            - Permissions: Ensure write permissions in the directory.
            """)

        col1, col2 = st.columns(2)
        with col1:
            if st.button("Try Again", use_container_width=True):
                st.session_state.stage = 'processing'
                add_message('bot', f"Retrying {info['display_name']} report...")
                st.rerun()
        with col2:
            if st.button("Start Over", use_container_width=True):
                st.session_state.messages = []
                st.session_state.stage = 'welcome'
                st.session_state.selected_project = None
                st.session_state.report_file = None
                if hasattr(st.session_state, 'error_message'):
                    delattr(st.session_state, 'error_message')
                st.rerun()

    st.markdown("---")
    st.markdown("""
    <div class="footer">
        <div>Milestone Report Generator</div>
        <div>Supported Projects: Veridia • Eligo • EWS-LIG • WaveCityClub • Eden</div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown('</div>', unsafe_allow_html=True)

if __name__ == "__main__":
    main()
