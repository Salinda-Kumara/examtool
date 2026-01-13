import subprocess
import sys
import os
import webbrowser
import time
import socket

def find_free_port():
    """Find a free port on localhost"""
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.bind(('', 0))
        return s.getsockname()[1]

def main():
    # Get the directory where the exe is located
    if getattr(sys, 'frozen', False):
        app_dir = os.path.dirname(sys.executable)
    else:
        app_dir = os.path.dirname(os.path.abspath(__file__))
    
    app_path = os.path.join(app_dir, 'app.py')
    port = find_free_port()
    
    print(f"Starting SAB Campus Excel Analyzer on port {port}...")
    print("Please wait, the browser will open automatically...")
    
    # Start streamlit
    process = subprocess.Popen([
        sys.executable, '-m', 'streamlit', 'run', app_path,
        '--server.port', str(port),
        '--server.headless', 'true',
        '--browser.gatherUsageStats', 'false'
    ], cwd=app_dir)
    
    # Wait a bit then open browser
    time.sleep(3)
    webbrowser.open(f'http://localhost:{port}')
    
    print(f"\nApp running at: http://localhost:{port}")
    print("Press Ctrl+C to stop the server...")
    
    try:
        process.wait()
    except KeyboardInterrupt:
        process.terminate()
        print("\nServer stopped.")

if __name__ == '__main__':
    main()
