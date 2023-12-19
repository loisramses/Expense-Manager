import os
import subprocess

def create_venv():
    if not os.path.exists('.venv'):
        subprocess.run(['python', '-m', 'venv', '.venv'])
    
        freeze_output = subprocess.run([os.path.join('.venv', 'Scripts', 'python'), '-m', 'pip', 'freeze'], capture_output=True, text=True)
        with open('requirements.txt', 'r') as req_file:
            requirements_content = req_file.read()

        if freeze_output.stdout.strip() != requirements_content.strip():
            install_requirements()

def run_in_venv(command):
    subprocess.run([os.path.join('.venv', 'Scripts', 'python')] + command, shell=True)

def install_requirements():
    run_in_venv(['-m', 'pip', 'install', '-r', 'requirements.txt'])

def main():
    create_venv()
    run_in_venv(['Main.py'])

if __name__ == "__main__":
    main()
