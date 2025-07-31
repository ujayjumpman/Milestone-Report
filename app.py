import streamlit as st
import subprocess
import sys
import os
import time
from datetime import datetime
import glob

# Page configuration
st.set_page_config(
    page_title="Milestone Report Generator",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Enhanced Custom CSS for modern, visually appealing interface
st.markdown("""
<style>
    /* Import Google Fonts */
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
    
    /* Global styles */
    .stApp {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        font-family: 'Inter', sans-serif;
    }
    
    /* Hide Streamlit default elements - but keep the page title tab */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    div[data-testid="stToolbar"] {visibility: hidden;}
    .stDeployButton {display: none;}
    div[data-testid="stDecoration"] {display: none;}
    
    /* Main container styling */
    .main-container {
        background: rgba(255, 255, 255, 0.95);
        backdrop-filter: blur(10px);
        border-radius: 20px;
        padding: 2rem;
        margin: 2rem auto;
        max-width: 1200px;
        box-shadow: 0 20px 40px rgba(0, 0, 0, 0.1);
        border: 1px solid rgba(255, 255, 255, 0.3);
    }
    
    /* Title styling */
    .main-title {
        text-align: center;
        font-size: 3rem;
        font-weight: 900;
        color: #000000 !important;
        margin-bottom: 0.5rem;
        text-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
    }
    
    .subtitle {
        text-align: center;
        font-size: 1.2rem;
        color: #000000 !important;
        margin-bottom: 2rem;
        font-weight: 700;
    }
    
    /* Chat message styling */
    .chat-message {
        padding: 1.5rem;
        border-radius: 15px;
        margin-bottom: 1.5rem;
        display: flex;
        align-items: flex-start;
        animation: fadeInUp 0.3s ease-out;
        box-shadow: 0 4px 15px rgba(0, 0, 0, 0.08);
    }
    
    @keyframes fadeInUp {
        from {
            opacity: 0;
            transform: translateY(20px);
        }
        to {
            opacity: 1;
            transform: translateY(0);
        }
    }
    
    .chat-message.bot {
        background: linear-gradient(135deg, #f8f9ff, #e8f0fe);
        margin-right: 2rem;
        border-left: 4px solid #667eea;
    }
    
    .chat-message.user {
        background: linear-gradient(135deg, #e3f2fd, #f1f8e9);
        margin-left: 2rem;
        flex-direction: row-reverse;
        border-right: 4px solid #4caf50;
    }
    
    .chat-message .avatar {
        width: 3.5rem;
        height: 3.5rem;
        border-radius: 50%;
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 1.5rem;
        margin: 0 1rem;
        background: linear-gradient(135deg, #667eea, #764ba2);
        color: white;
        box-shadow: 0 4px 15px rgba(102, 126, 234, 0.3);
    }
    
    .user .avatar {
        background: linear-gradient(135deg, #4caf50, #45a049);
        box-shadow: 0 4px 15px rgba(76, 175, 80, 0.3);
    }
    
    .chat-message .message {
        flex: 1;
        padding: 0 0.5rem;
        font-size: 1.1rem;
        line-height: 1.6;
        color: #000000 !important;
        font-weight: 600;
    }
    
    /* Project selection section */
    .project-selection {
        background: linear-gradient(135deg, #f8f9ff, #fff);
        border-radius: 20px;
        padding: 2rem;
        margin: 2rem 0;
        box-shadow: 0 10px 30px rgba(0, 0, 0, 0.08);
        border: 1px solid rgba(102, 126, 234, 0.1);
    }
    
    .project-selection h3 {
        text-align: center;
        color: #000000 !important;
        font-size: 1.8rem;
        font-weight: 800;
        margin-bottom: 2rem;
    }
    
    /* Project buttons */
    .project-buttons {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
        gap: 1.5rem;
        margin: 2rem 0;
    }
    
    .stButton > button {
        background: linear-gradient(135deg, #667eea, #764ba2) !important;
        color: white !important;
        border: none !important;
        padding: 1.2rem 2rem !important;
        border-radius: 15px !important;
        font-size: 1.1rem !important;
        font-weight: 600 !important;
        transition: all 0.3s ease !important;
        box-shadow: 0 8px 25px rgba(102, 126, 234, 0.3) !important;
        width: 100% !important;
        height: 70px !important;
        text-transform: uppercase !important;
        letter-spacing: 0.5px !important;
    }
    
    .stButton > button:hover {
        transform: translateY(-5px) !important;
        box-shadow: 0 15px 35px rgba(102, 126, 234, 0.4) !important;
        background: linear-gradient(135deg, #764ba2, #667eea) !important;
    }
    
    .stButton > button:active {
        transform: translateY(-2px) !important;
    }
    
    /* Status containers */
    .status-container {
        background: linear-gradient(135deg, #fff3cd, #ffeaa7);
        border: 2px solid #f39c12;
        border-radius: 15px;
        padding: 2rem;
        margin: 2rem 0;
        text-align: center;
        box-shadow: 0 8px 25px rgba(243, 156, 18, 0.2);
    }
    
    .status-container h4 {
        color: #d68910;
        font-size: 1.5rem;
        font-weight: 600;
        margin-bottom: 1rem;
    }
    
    .status-container p {
        color: #7d6608;
        font-size: 1.1rem;
        margin-bottom: 0;
    }
    
    .success-container {
        background: linear-gradient(135deg, #d4edda, #c3e6cb);
        border: 2px solid #27ae60;
        border-radius: 15px;
        padding: 2rem;
        margin: 2rem 0;
        text-align: center;
        box-shadow: 0 8px 25px rgba(39, 174, 96, 0.2);
    }
    
    .success-container h4 {
        color: #27ae60;
        font-size: 1.5rem;
        font-weight: 600;
        margin-bottom: 1rem;
    }
    
    .success-container p {
        color: #155724;
        font-size: 1.1rem;
        margin-bottom: 0;
    }
    
    .error-container {
        background: linear-gradient(135deg, #f8d7da, #f5c6cb);
        border: 2px solid #e74c3c;
        border-radius: 15px;
        padding: 2rem;
        margin: 2rem 0;
        text-align: center;
        box-shadow: 0 8px 25px rgba(231, 76, 60, 0.2);
    }
    
    .error-container h4 {
        color: #c0392b;
        font-size: 1.5rem;
        font-weight: 600;
        margin-bottom: 1rem;
    }
    
    .error-container p {
        color: #721c24;
        font-size: 1.1rem;
        margin-bottom: 0;
    }
    
    /* Download button styling */
    .stDownloadButton > button {
        background: linear-gradient(135deg, #27ae60, #2ecc71) !important;
        color: white !important;
        border: none !important;
        padding: 1.5rem 3rem !important;
        border-radius: 15px !important;
        font-size: 1.2rem !important;
        font-weight: 600 !important;
        transition: all 0.3s ease !important;
        box-shadow: 0 8px 25px rgba(39, 174, 96, 0.3) !important;
        width: 100% !important;
        height: 80px !important;
        text-transform: uppercase !important;
        letter-spacing: 0.5px !important;
    }
    
    .stDownloadButton > button:hover {
        transform: translateY(-3px) !important;
        box-shadow: 0 12px 30px rgba(39, 174, 96, 0.4) !important;
        background: linear-gradient(135deg, #2ecc71, #27ae60) !important;
    }
    
    /* Progress bar styling */
    .stProgress > div > div > div > div {
        background: linear-gradient(135deg, #667eea, #764ba2) !important;
        border-radius: 10px !important;
    }
    
    .stProgress > div > div > div {
        background-color: rgba(102, 126, 234, 0.1) !important;
        border-radius: 10px !important;
        height: 15px !important;
    }
    
    /* Footer styling */
    .footer {
        text-align: center;
        color: rgba(255, 255, 255, 0.8);
        font-size: 1rem;
        margin-top: 3rem;
        padding: 2rem;
        background: rgba(255, 255, 255, 0.1);
        border-radius: 15px;
        backdrop-filter: blur(10px);
    }
    
    /* Divider */
    hr {
        border: none;
        height: 2px;
        background: linear-gradient(135deg, #667eea, #764ba2);
        margin: 2rem 0;
        border-radius: 2px;
    }
    
    /* Custom scrollbar */
    ::-webkit-scrollbar {
        width: 8px;
    }
    
    ::-webkit-scrollbar-track {
        background: rgba(255, 255, 255, 0.1);
        border-radius: 10px;
    }
    
    ::-webkit-scrollbar-thumb {
        background: linear-gradient(135deg, #667eea, #764ba2);
        border-radius: 10px;
    }
    
    ::-webkit-scrollbar-thumb:hover {
        background: linear-gradient(135deg, #764ba2, #667eea);
    }
    
    /* Responsive design */
    @media (max-width: 768px) {
        .main-title {
            font-size: 2rem;
        }
        
        .main-container {
            margin: 1rem;
            padding: 1rem;
        }
        
        .chat-message {
            margin: 0.5rem 0;
        }
        
        .chat-message.bot {
            margin-right: 0.5rem;
        }
        
        .chat-message.user {
            margin-left: 0.5rem;
        }
    }
</style>
""", unsafe_allow_html=True)

# Initialize session state
if 'messages' not in st.session_state:
    st.session_state.messages = []
    st.session_state.stage = 'welcome'
    st.session_state.selected_project = None
    st.session_state.report_file = None

# Project configurations - Fixed script names and patterns
PROJECTS = {
    'Veridia': {
        'script': 'veridia.py',
        'display_name': 'Veridia',
        'icon': '🌿',
        'pattern': 'Time_Delivery_Milestones_Report_*.xlsx'  # Based on original code
    },
    'Eligo': {
        'script': 'eligo.py', 
        'display_name': 'Eligo',
        'icon': '⚡',
        'pattern': '*Eligo*.xlsx'
    },
    'EWS-LIG': {
        'script': 'ews-lig.py',  # Fixed script name
        'display_name': 'EWS-LIG',
        'icon': '🔍',
        'pattern': '*EWS*LIG*.xlsx'
    },
    'WaveCityClub': {  # Fixed key name
        'script': 'wavecityclub.py',
        'display_name': 'WaveCityClub',
        'icon': '🌊',
        'pattern': 'Wave_City_Club_Report_*.xlsx'  # Based on original code
    },
    'Eden': {
        'script': 'eden.py',
        'display_name': 'Eden',
        'icon': '🏡',
        'pattern': 'Eden_KRA_Milestone_Report_*.xlsx'  # Based on Eden script
    }
}

def add_message(role, content):
    """Add a message to the chat history"""
    st.session_state.messages.append({
        'role': role,
        'content': content,
        'timestamp': datetime.now()
    })

def display_chat_message(message):
    """Display a single chat message"""
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

def run_project_script(project_name):
    """Run the project script and return the generated file path"""
    try:
        project_config = PROJECTS[project_name]
        script_path = project_config['script']
        
        # Check if script file exists
        if not os.path.exists(script_path):
            return False, f"Script file '{script_path}' not found. Please ensure the file exists in the current directory."
        
        # Add logging for debugging
        st.write(f"🔄 Running script: {script_path}")
        
        # Execute the script
        result = subprocess.run(
            [sys.executable, script_path],
            capture_output=True,
            text=True,
            timeout=300  # 5 minutes timeout
        )
        
        # Check execution result
        if result.returncode != 0:
            error_msg = result.stderr or result.stdout or "Unknown error occurred"
            st.error(f"Script execution failed: {error_msg[:500]}...")
            return False, f"Script execution failed:\n{error_msg[:500]}..."
        
        # Log success
        if result.stdout:
            st.write(f"✅ Script output: {result.stdout[:200]}...")
        
        # Look for generated file using project-specific pattern
        pattern = project_config['pattern']
        st.write(f"🔍 Looking for files matching pattern: {pattern}")
        
        matches = glob.glob(pattern)
        
        if not matches:
            # Fallback patterns
            fallback_patterns = [
                f"*{project_config['display_name']}*.xlsx",
                f"*{project_name}*.xlsx",
                "*.xlsx"
            ]
            
            for fallback_pattern in fallback_patterns:
                st.write(f"🔍 Trying fallback pattern: {fallback_pattern}")
                matches = glob.glob(fallback_pattern)
                if matches:
                    break
        
        if not matches:
            # List all Excel files for debugging
            all_excel = glob.glob("*.xlsx")
            st.write(f"📁 All Excel files found: {all_excel}")
            return False, f"Report was generated but no matching file found. Pattern used: {pattern}. Available Excel files: {all_excel}"
        
        # Get the most recent file
        latest_file = max(matches, key=os.path.getctime)
        st.write(f"📊 Found report file: {latest_file}")
        
        return True, latest_file

    except subprocess.TimeoutExpired:
        return False, "Script execution timed out (5 minutes). Please try again or check if the script is stuck."
    except FileNotFoundError:
        return False, f"Python interpreter not found. Please ensure Python is installed and accessible."
    except Exception as e:
        return False, f"Unexpected error running script: {str(e)}"

def main():
    st.markdown('<div class="main-container">', unsafe_allow_html=True)

    # Title
    st.markdown("""
    <div class="main-title">📊 Milestone Report Generator</div>
    <div class="subtitle">Generate comprehensive milestone reports with just one click</div>
    """, unsafe_allow_html=True)
    st.markdown("---")

    # Intro messages
    if not st.session_state.messages:
        add_message('bot', "Hello! 👋 Welcome to the Milestone Report Generator.")
        add_message('bot', "Which project would you like to generate a milestone report for?")

    # Display chat history
    for msg in st.session_state.messages:
        display_chat_message(msg)

    # Welcome / Project selection
    if st.session_state.stage == 'welcome' and not st.session_state.selected_project:
        st.markdown('<div class="project-selection"><h3>🚀 Select Your Project</h3></div>', unsafe_allow_html=True)
        
        # Display project buttons in a grid
        cols = st.columns(len(PROJECTS))
        for idx, (key, info) in enumerate(PROJECTS.items()):
            with cols[idx]:
                if st.button(f"{info['icon']} {info['display_name']}", key=key):
                    st.session_state.selected_project = key
                    st.session_state.stage = 'processing'
                    add_message('user', f"I want to generate a milestone report for {info['display_name']}.")
                    add_message('bot', f"Excellent choice! I'll generate the {info['display_name']} report now. Please wait...")
                    st.rerun()

    # Processing stage
    elif st.session_state.stage == 'processing':
        proj = st.session_state.selected_project
        info = PROJECTS[proj]
        
        st.markdown(f"""
        <div class="status-container">
          <h4>{info['icon']} Processing {info['display_name']}...</h4>
          <p>Please wait while I generate your report. This may take a few minutes.</p>
        </div>
        """, unsafe_allow_html=True)

        # Progress bar animation
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
            time.sleep(0.8)
        
        # Clear progress indicators
        progress_bar.empty()
        status_text.empty()
        
        # Run the actual script
        success, result = run_project_script(proj)
        
        if success:
            st.session_state.report_file = result
            st.session_state.stage = 'completed'
            add_message('bot', f"✅ Your {info['display_name']} report has been generated successfully!")
        else:
            st.session_state.stage = 'error'
            st.session_state.error_message = result
            add_message('bot', f"❌ There was an error generating the {info['display_name']} report.")
        
        st.rerun()

    # Completed stage
    elif st.session_state.stage == 'completed':
        proj = st.session_state.selected_project
        info = PROJECTS[proj]
        
        st.markdown(f"""
        <div class="success-container">
          <h4>{info['icon']} Report Generated Successfully!</h4>
          <p>Your {info['display_name']} milestone report is ready to download.</p>
        </div>
        """, unsafe_allow_html=True)

        # Display file info
        if st.session_state.report_file and os.path.exists(st.session_state.report_file):
            file_size = os.path.getsize(st.session_state.report_file)
            file_size_mb = file_size / (1024 * 1024)
            st.info(f"📄 File: {os.path.basename(st.session_state.report_file)} ({file_size_mb:.2f} MB)")
            
            # Download button
            with open(st.session_state.report_file, "rb") as f:
                file_data = f.read()
            
            st.download_button(
                label=f"📥 Download {info['display_name']} Report",
                data=file_data,
                file_name=os.path.basename(st.session_state.report_file),
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        else:
            st.error("Report file not found or was deleted.")

        # Generate another report button
        if st.button("🔄 Generate Another Report", use_container_width=True):
            st.session_state.messages = []
            st.session_state.stage = 'welcome'
            st.session_state.selected_project = None
            st.session_state.report_file = None
            st.rerun()

    # Error stage
    elif st.session_state.stage == 'error':
        proj = st.session_state.selected_project or ""
        info = PROJECTS.get(proj, {'display_name': 'Unknown', 'icon': '❌'})
        
        st.markdown(f"""
        <div class="error-container">
          <h4>{info['icon']} Error Generating {info['display_name']} Report</h4>
          <p>There was an issue generating your report. Please check the details below and try again.</p>
        </div>
        """, unsafe_allow_html=True)

        # Show error details
        if hasattr(st.session_state, 'error_message'):
            with st.expander("🔍 Error Details"):
                st.code(st.session_state.error_message)

        # Action buttons
        col1, col2 = st.columns(2)
        with col1:
            if st.button("🔄 Try Again", use_container_width=True):
                st.session_state.stage = 'processing'
                add_message('bot', f"Retrying the {info['display_name']} report generation...")
                st.rerun()
        with col2:
            if st.button("🏠 Start Over", use_container_width=True):
                st.session_state.messages = []
                st.session_state.stage = 'welcome'
                st.session_state.selected_project = None
                st.session_state.report_file = None
                if hasattr(st.session_state, 'error_message'):
                    delattr(st.session_state, 'error_message')
                st.rerun()

    # Footer
    st.markdown("---")
    st.markdown("""
    <div class="footer">
      <div style="font-size:1.2rem;">📊 Milestone Report Generator</div>
      <div>Automated report generation for project milestones</div>
      <div style="margin-top:1rem; font-size:0.9rem;">
        Supported Projects: Veridia • Eligo • EWS-LIG • WaveCityClub • Eden
      </div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown('</div>', unsafe_allow_html=True)

if __name__ == "__main__":
    main()