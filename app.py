import streamlit as st
import subprocess
import sys
import os
import time
import gc
import psutil
from datetime import datetime
import glob
import traceback

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
    
    /* Clear memory button styling */
    .clear-memory-btn {
        background: linear-gradient(135deg, #e74c3c, #c0392b) !important;
        color: white !important;
        border: none !important;
        padding: 0.8rem 1.5rem !important;
        border-radius: 10px !important;
        font-size: 0.9rem !important;
        font-weight: 600 !important;
        transition: all 0.3s ease !important;
        box-shadow: 0 4px 15px rgba(231, 76, 60, 0.3) !important;
        width: 100% !important;
        height: 50px !important;
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
    
    /* Debug info styling */
    .debug-info {
        background: linear-gradient(135deg, #f0f4f8, #e2e8f0);
        border: 2px solid #718096;
        border-radius: 10px;
        padding: 1rem;
        margin: 1rem 0;
        font-family: monospace;
        font-size: 0.9rem;
        color: #2d3748;
    }
    
    /* System info styling */
    .system-info {
        background: linear-gradient(135deg, #e8f5e8, #f0f8f0);
        border: 2px solid #27ae60;
        border-radius: 10px;
        padding: 1rem;
        margin: 1rem 0;
        font-size: 0.9rem;
        color: #155724;
    }
</style>
""", unsafe_allow_html=True)

# Initialize session state
if 'messages' not in st.session_state:
    st.session_state.messages = []
    st.session_state.stage = 'welcome'
    st.session_state.selected_project = None
    st.session_state.report_file = None

# Project configurations - Enhanced with better debugging
PROJECTS = {
    'Veridia': {
        'script': 'veridia.py',
        'display_name': 'Veridia',
        'icon': '🌿',
        'patterns': [
            'Time_Delivery_Milestones_Report_*.xlsx',
            '*Veridia*.xlsx',
            'Veridia_*.xlsx',
            '*veridia*.xlsx'
        ]
    },
    'Eligo': {
        'script': 'eligo.py', 
        'display_name': 'Eligo',
        'icon': '⚡',
        'patterns': [
            '*Eligo*.xlsx',
            'Eligo_*.xlsx',
            '*eligo*.xlsx'
        ]
    },
    'EWS-LIG': {
        'script': 'ews-lig.py',
        'display_name': 'EWS-LIG',
        'icon': '🔍',
        'patterns': [
            '*EWS*LIG*.xlsx',
            '*EWS-LIG*.xlsx',
            'EWS_LIG_*.xlsx',
            '*ews*lig*.xlsx'
        ]
    },
    'WaveCityClub': {
        'script': 'wavecityclub.py',
        'display_name': 'WaveCityClub',
        'icon': '🌊',
        'patterns': [
            'Wave_City_Club_Report_*.xlsx',
            '*WaveCityClub*.xlsx',
            '*Wave*City*Club*.xlsx',
            '*wavecityclub*.xlsx'
        ]
    },
    'Eden': {
        'script': 'eden.py',
        'display_name': 'Eden',
        'icon': '🏡',
        'patterns': [
            'Eden_KRA_Milestone_Report_*.xlsx',
            '*Eden*.xlsx',
            'Eden_*.xlsx',
            '*eden*.xlsx'
        ]
    }
}

def cleanup_resources():
    """Clean up system resources between script executions"""
    try:
        st.write("🧹 **Cleaning up system resources...**")
        
        # Force garbage collection
        gc.collect()
        
        # Kill any orphaned Python processes (be careful with this)
        current_pid = os.getpid()
        killed_processes = 0
        
        for proc in psutil.process_iter(['pid', 'name', 'cmdline']):
            try:
                if proc.info['name'] in ['python', 'python.exe']:
                    # Check if it's a subprocess of our scripts
                    cmdline = proc.info['cmdline'] or []
                    script_names = ['veridia.py', 'eligo.py', 'ews-lig.py', 'wavecityclub.py', 'eden.py']
                    if any(script in ' '.join(cmdline) for script in script_names):
                        if proc.info['pid'] != current_pid:
                            proc.terminate()
                            proc.wait(timeout=3)
                            killed_processes += 1
            except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.TimeoutExpired):
                continue
        
        if killed_processes > 0:
            st.write(f"🔄 Terminated {killed_processes} orphaned script processes")
        
        # Clear any temporary files that might be locked
        temp_patterns = ['~$*.xlsx', '*.tmp', '.~lock.*', '__pycache__']
        cleaned_files = 0
        
        for pattern in temp_patterns:
            for file in glob.glob(pattern):
                try:
                    if os.path.isfile(file):
                        os.remove(file)
                        cleaned_files += 1
                    elif os.path.isdir(file):
                        import shutil
                        shutil.rmtree(file)
                        cleaned_files += 1
                except:
                    pass
        
        if cleaned_files > 0:
            st.write(f"🗑️ Cleaned {cleaned_files} temporary files")
        
        # Brief pause to let system settle
        time.sleep(2)
        
        # Show memory status after cleanup
        memory_info = psutil.virtual_memory()
        st.write(f"💾 **Memory after cleanup:** {memory_info.percent:.1f}% used ({memory_info.available / (1024**3):.1f} GB available)")
        st.success("✅ Resource cleanup completed!")
        
    except Exception as e:
        st.write(f"⚠️ Resource cleanup warning: {e}")

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

def find_generated_file(project_config, project_name):
    """Find the generated report file using multiple patterns"""
    patterns = project_config['patterns']
    
    for pattern in patterns:
        st.write(f"🔍 Searching with pattern: {pattern}")
        matches = glob.glob(pattern)
        if matches:
            # Get the most recent file
            latest_file = max(matches, key=os.path.getctime)
            file_time = os.path.getctime(latest_file)
            current_time = time.time()
            
            # Check if file was created recently (within last 10 minutes)
            if (current_time - file_time) < 600:  # 10 minutes
                st.write(f"✅ Found recent file: {latest_file}")
                return latest_file
            else:
                st.write(f"⏰ File found but too old: {latest_file}")
    
    # Check for any new Excel files created
    all_excel = glob.glob("*.xlsx")
    if all_excel:
        latest_new_file = max(all_excel, key=os.path.getctime)
        file_time = os.path.getctime(latest_new_file)
        current_time = time.time()
        
        if (current_time - file_time) < 600:  # 10 minutes
            st.write(f"📄 Found recent Excel file: {latest_new_file}")
            return latest_new_file
    
    return None

def monitor_memory_during_execution():
    """Monitor and display current system status"""
    try:
        memory_info = psutil.virtual_memory()
        cpu_percent = psutil.cpu_percent(interval=1)
        
        st.markdown(f"""
        <div class="system-info">
            <strong>💻 System Status:</strong><br>
            🧠 Memory: {memory_info.percent:.1f}% used ({memory_info.available / (1024**3):.1f} GB available)<br>
            ⚡ CPU: {cpu_percent:.1f}% usage<br>
            🔧 Active Python processes: {len([p for p in psutil.process_iter() if 'python' in p.name().lower()])}
        </div>
        """, unsafe_allow_html=True)
        
        return memory_info.percent
        
    except Exception as e:
        st.write(f"⚠️ Could not monitor system status: {e}")
        return 0

def run_project_script(project_name):
    """Enhanced script execution with proper resource management"""
    try:
        project_config = PROJECTS[project_name]
        script_path = project_config['script']
        
        # Show system status before execution
        memory_before = monitor_memory_during_execution()
        
        # Enhanced debugging information
        st.write(f"🔧 **Debug Information for {project_name}:**")
        st.write(f"📝 Script path: {script_path}")
        st.write(f"📁 Current directory: {os.getcwd()}")
        st.write(f"🐍 Python executable: {sys.executable}")
        
        # Check if script file exists
        if not os.path.exists(script_path):
            available_files = [f for f in os.listdir('.') if f.endswith('.py')]
            return False, f"❌ Script file '{script_path}' not found in current directory.\n\nAvailable Python files: {available_files}"
        
        st.write(f"✅ Script file found: {script_path}")
        
        # Store existing Excel files before execution
        files_before = set(glob.glob("*.xlsx"))
        st.write(f"📊 Excel files before execution: {len(files_before)}")
        
        # Enhanced timeout settings
        timeout_settings = {
            'Veridia': 900,      # 15 minutes for Veridia
            'Eligo': 600,        # 10 minutes for Eligo  
            'EWS-LIG': 600,      # 10 minutes
            'WaveCityClub': 450, # 7.5 minutes
            'Eden': 450          # 7.5 minutes
        }
        timeout_duration = timeout_settings.get(project_name, 300)
        
        st.write(f"🚀 Executing script: {script_path} (timeout: {timeout_duration//60} minutes)")
        
        # Create enhanced environment
        env = os.environ.copy()
        env.update({
            'PYTHONUNBUFFERED': '1',
            'MPLBACKEND': 'Agg',
            'OPENBLAS_NUM_THREADS': '1',  # Limit BLAS threads
            'MKL_NUM_THREADS': '1',       # Limit MKL threads
            'NUMEXPR_NUM_THREADS': '1',   # Limit NumExpr threads
            'OMP_NUM_THREADS': '1',       # Limit OpenMP threads
            'PYTHONDONTWRITEBYTECODE': '1',  # Don't create .pyc files
            'PYTHONHASHSEED': '0',           # Consistent hashing
        })
        
        start_time = time.time()
        progress_placeholder = st.empty()
        progress_placeholder.info(f"⏱️ Running {project_name} script... (Max wait: {timeout_duration//60} minutes)")
        
        # Use Popen for better process control
        process = subprocess.Popen(
            [sys.executable, script_path],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            cwd=os.getcwd(),
            env=env,
            bufsize=1,
            universal_newlines=True
        )
        
        try:
            # Wait for completion with timeout
            stdout, stderr = process.communicate(timeout=timeout_duration)
            elapsed_time = time.time() - start_time
            
            if process.returncode == 0:
                progress_placeholder.success(f"✅ Script completed in {elapsed_time:.1f} seconds")
            else:
                progress_placeholder.error(f"❌ Script failed with return code {process.returncode}")
                
        except subprocess.TimeoutExpired:
            # Handle timeout more gracefully
            elapsed_time = time.time() - start_time
            progress_placeholder.error(f"⏱️ Script timed out after {elapsed_time:.1f} seconds")
            
            # Terminate the process
            process.terminate()
            try:
                process.wait(timeout=10)
            except subprocess.TimeoutExpired:
                process.kill()
                process.wait()
            
            # Clean up any remaining processes
            st.write("🧹 **Performing emergency cleanup...**")
            cleanup_resources()
            
            timeout_msg = f"""
⏱️ **{project_name} script timed out after {timeout_duration//60} minutes.**

**This could indicate:**
1. **Large dataset processing** - The script is processing very large files
2. **Memory issues** - System running out of memory (was {memory_before:.1f}% before execution)
3. **Infinite loop** - Bug in the script causing it to loop
4. **Resource contention** - Previous script execution interfering

**Recommended actions:**
1. Click "Clear Memory" button before trying again
2. Check system resources (memory/CPU usage)
3. Try running `{script_path}` manually to identify the bottleneck
4. Consider restarting the Streamlit app to clear all resources
5. Break large input files into smaller chunks if applicable

**To debug manually:**
```bash
cd {os.getcwd()}
python {script_path}
```

This will show you exactly where the script stops or gets stuck.
            """
            return False, timeout_msg
        
        # Show memory status after execution
        memory_after = monitor_memory_during_execution()
        memory_change = memory_after - memory_before
        if memory_change > 10:
            st.warning(f"⚠️ Significant memory increase: +{memory_change:.1f}%")
        
        # Enhanced result logging
        st.write(f"📤 Script execution completed with return code: {process.returncode}")
        
        if stdout:
            st.write("📄 **Script Output (stdout):**")
            st.code(stdout[:1000] + ("..." if len(stdout) > 1000 else ""))
        
        if stderr:
            st.write("⚠️ **Script Errors (stderr):**")
            st.code(stderr[:1000] + ("..." if len(stderr) > 1000 else ""))
        
        # Check execution result
        if process.returncode != 0:
            error_details = f"""
Return Code: {process.returncode}
STDOUT: {stdout}
STDERR: {stderr}
            """
            return False, f"❌ Script execution failed with return code {process.returncode}.\n\nDetails:\n{error_details}"
        
        # Check for new files after execution
        files_after = set(glob.glob("*.xlsx"))
        new_files = files_after - files_before
        st.write(f"📊 Excel files after execution: {len(files_after)} (New: {len(new_files)})")
        
        if new_files:
            st.write(f"🆕 New Excel files created: {list(new_files)}")
        
        # Look for generated file
        generated_file = find_generated_file(project_config, project_name)
        
        if generated_file and os.path.exists(generated_file):
            file_size = os.path.getsize(generated_file)
            st.write(f"✅ **Report file found:** {generated_file} ({file_size:,} bytes)")
            return True, generated_file
        
        # Diagnostic information if file not found
        all_excel = glob.glob("*.xlsx")
        error_msg = f"""
❌ **Report file not found after script execution.**

**Diagnostics:**
- Script executed successfully (return code: {process.returncode})
- Patterns searched: {project_config['patterns']}
- All Excel files in directory: {all_excel}
- New files created: {list(new_files) if new_files else 'None'}

**Script Output:**
STDOUT: {stdout[:500]}...
STDERR: {stderr[:500]}...
        """
        return False, error_msg

    except Exception as e:
        cleanup_resources()  # Clean up on any error
        error_details = f"""
Exception Type: {type(e).__name__}
Exception Message: {str(e)}
Traceback: {traceback.format_exc()}
        """
        return False, f"❌ Unexpected error occurred:\n{error_details}"

def main():
    st.markdown('<div class="main-container">', unsafe_allow_html=True)

    # Title
    st.markdown("""
    <div class="main-title">📊 Milestone Report Generator</div>
    <div class="subtitle">Generate comprehensive milestone reports with just one click</div>
    """, unsafe_allow_html=True)
    
    # Add Clear Memory button if there's a selected project
    if st.session_state.get('selected_project') or st.session_state.stage != 'welcome':
        col1, col2 = st.columns([4, 1])
        with col2:
            if st.button("🧹 Clear Memory", help="Clear system resources and memory", key="clear_memory_btn"):
                with st.spinner("Clearing system resources..."):
                    cleanup_resources()
                    time.sleep(1)
                st.rerun()
    
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
        
        # Create a debug expander for detailed logging
        with st.expander("🔍 Debug Information", expanded=True):
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
            with st.expander("🔍 Detailed Error Information", expanded=True):
                st.code(st.session_state.error_message)

        # Troubleshooting tips
        with st.expander("💡 Troubleshooting Tips"):
            st.markdown(f"""
            **Common issues and solutions:**
            
            1. **Script file missing**: Ensure `{info.get('script', 'unknown.py')}` exists in the same directory as this Streamlit app.
            
            2. **Import errors**: Check if all required Python packages are installed.
            
            3. **Data file missing**: Ensure any required input data files are in the correct location.
            
            4. **Permission issues**: Check if the script has permission to write files to the current directory.
            
            5. **Path issues**: Verify that all file paths in the script are correct.
            
            6. **Memory/Resource issues**: Try clicking "Clear Memory" button and then retry.
            
            **Next steps:**
            - Try running `{info.get('script', 'unknown.py')}` manually from the command line
            - Check the script's dependencies and requirements
            - Verify input data files are present and accessible
            - Clear memory and restart if needed
            """)

        # Action buttons
        col1, col2, col3 = st.columns(3)
        with col1:
            if st.button("🧹 Clear & Retry", use_container_width=True, help="Clear memory and try again"):
                with st.spinner("Clearing resources..."):
                    cleanup_resources()
                st.session_state.stage = 'processing'
                add_message('bot', f"🔄 Cleared memory and retrying the {info['display_name']} report generation...")
                st.rerun()
        with col2:
            if st.button("🔄 Try Again", use_container_width=True):
                st.session_state.stage = 'processing'
                add_message('bot', f"Retrying the {info['display_name']} report generation...")
                st.rerun()
        with col3:
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
      <div style="margin-top:0.5rem; font-size:0.8rem; color: rgba(255,255,255,0.6);">
        💡 Tip: Use "Clear Memory" between reports for optimal performance
      </div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown('</div>', unsafe_allow_html=True)

if __name__ == "__main__":
    main()
