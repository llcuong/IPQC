import subprocess
from datetime import datetime

def git_commit():
    timestamp = datetime.now().strftime("%y%m%d-%H%M")
    print(timestamp)
    commit_message = f'{timestamp}'
    try:
        subprocess.run(["git", "add", "."], check=True)
        subprocess.run(["git", "commit", "-m", commit_message], check=True)
        subprocess.run(["git", "push"], check=True)
    except subprocess.CalledProcessError as e:
        print("Error during git operations:", e)
if __name__ == "__main__":
    git_commit()