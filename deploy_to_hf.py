import os
import subprocess
import sys
import getpass
import time

def run_command(command, cwd=None, capture_output=True):
    try:
        result = subprocess.run(
            command,
            cwd=cwd,
            shell=True,
            check=True,
            text=True,
            capture_output=capture_output
        )
        return result.stdout.strip()
    except subprocess.CalledProcessError as e:
        print(f"Error running command: {command}")
        print(f"Error output: {e.stderr}")
        return None

def main():
    print("\n🚀 Hugging Face Spaces Auto-Deployer")
    print("=====================================")
    print("This script will deploy your current code directly to a new Hugging Face Space.\n")

    # 1. Get Prerequisites
    username = input("Enter your Hugging Face Username: ").strip()
    if not username:
        print("Username is required.")
        return

    space_name = input(f"Enter a name for your new Space (e.g., ai-financial-modeler): ").strip()
    if not space_name:
        space_name = "ai-financial-modeler"
        print(f"Using default name: {space_name}")

    print("\n🔑 Authentication")
    print("Please paste your Hugging Face Access Token.")
    print("Get it here: https://huggingface.co/settings/tokens (Role: WRITE)")
    token = getpass.getpass("Token (hidden input): ").strip()
    
    if not token.startswith("hf_"):
        print("⚠️  Warning: Token usually starts with 'hf_'. Proceeding anyway...")

    repo_url = f"https://{username}:{token}@huggingface.co/spaces/{username}/{space_name}"
    
    print(f"\n⚙️  Configuring deployment to: spaces/{username}/{space_name}...")

    # 2. Check Git Status
    if not os.path.exists(".git"):
        print("❌ Not a git repository. Please run inside the project root.")
        return

    # 3. Configure Remote
    # Remove existing 'space' remote if it exists to ensure freshness
    try:
        subprocess.run("git remote remove space", shell=True, check=False, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    except:
        pass
    
    print("🔗 Adding 'space' remote...")
    # Add remote with token embedded
    cmd_add_remote = f"git remote add space {repo_url}"
    if run_command(cmd_add_remote) is None:
        return

    # 4. Push to Space
    print("\n🚀 Deploying... (This pushes your code to Hugging Face)")
    print("   This might take a minute. Please wait.")
    
    # Force push to main branch on the space
    push_cmd = "git push --force space main"
    
    try:
        result = subprocess.run(
            push_cmd,
            shell=True,
            check=True,
            text=True,
            capture_output=True # Capture to hide token in case of simple print, but verify error
        )
        print("\n✅ Deployment Command Sent Successfully!")
        print("--------------------------------------------------")
        print(f"👉 Your App is Building here: https://huggingface.co/spaces/{username}/{space_name}")
        print("--------------------------------------------------")
        print("Note: If the Space didn't exist, this push might have created it,")
        print("or you might need to create an empty Space with 'Docker' SDK first on the website if basic push fails.")
        
    except subprocess.CalledProcessError as e:
        print("\n❌ Deployment Failed.")
        print("Possible reasons:")
        print("1. The Space 'spaces/" + username + "/" + space_name + "' does not exist. Please create it first!")
        print("   -> Go to: https://huggingface.co/new-space")
        print("   -> Name: " + space_name)
        print("   -> SDK: Docker")
        print("   -> Then run this script again.")
        print("2. Invalid Token.")
        print(f"\nGit Error: {e.stderr}")

if __name__ == "__main__":
    main()
