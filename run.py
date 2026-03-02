"""
Grid — Double-click this file to start.
Opens your browser automatically.
"""
import subprocess
import sys
import os

def check_flask():
    try:
        import flask
        return True
    except ImportError:
        return False

def install_flask():
    print("First-time setup: installing Flask...")
    subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'flask', '-q'])
    print("Done!\n")

def main():
    os.chdir(os.path.dirname(os.path.abspath(__file__)))
    
    if not check_flask():
        install_flask()
    
    from app import main as run_app
    run_app()

if __name__ == '__main__':
    main()
