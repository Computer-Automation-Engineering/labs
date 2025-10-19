#!/usr/bin/env python3
"""
GIT PRACTICE WITH CLAUDE
=========================
Interactive practice session where you can practice Git commands
and get instant feedback from Claude.

Features:
- Command validation
- Safe execution in sandbox
- Instant feedback
- Best practices suggestions
- Real-time error detection

Author: Created for Jason's Git learning journey
Date: 2025-10-19
"""

import subprocess
import os
import sys
import shutil

class GitPractice:
    """Interactive Git practice with Claude's guidance"""

    def __init__(self):
        self.sandbox_dir = os.path.join(os.getcwd(), "git_practice_sandbox")
        self.setup_sandbox()

    def setup_sandbox(self):
        """Set up a safe practice environment"""
        if os.path.exists(self.sandbox_dir):
            response = input(f"\nSandbox directory exists. Reset it? (y/n): ").lower()
            if response == 'y':
                shutil.rmtree(self.sandbox_dir)
                os.makedirs(self.sandbox_dir)
                print("✓ Sandbox reset")
        else:
            os.makedirs(self.sandbox_dir)
            print("✓ Sandbox created")

    def run_git_command(self, command):
        """Execute a git command in the sandbox"""
        try:
            result = subprocess.run(
                command,
                shell=True,
                capture_output=True,
                text=True,
                cwd=self.sandbox_dir
            )
            return result.stdout, result.stderr, result.returncode
        except Exception as e:
            return "", f"Error: {str(e)}", 1

    def analyze_command(self, command):
        """Analyze the command and provide feedback"""
        feedback = []

        # Check for good practices
        if "git commit" in command and "-m" in command:
            feedback.append("✓ Good: Using -m flag for commit message")

        if "git add ." in command:
            feedback.append("⚠ Caution: 'git add .' stages all files. Make sure that's what you want.")

        if "git reset --hard" in command:
            feedback.append("⚠ WARNING: --hard will permanently delete changes!")

        if "git push -f" in command or "git push --force" in command:
            feedback.append("⚠ DANGER: Force push can overwrite remote history!")

        # Check for common mistakes
        if "git commit" in command and "-m" not in command and "--amend" not in command:
            feedback.append("ℹ Tip: Consider using -m flag: git commit -m \"message\"")

        return feedback

    def free_practice(self):
        """Free practice mode - execute any git command"""
        print("\n" + "="*60)
        print("FREE PRACTICE MODE")
        print("="*60)
        print(f"\nSandbox directory: {self.sandbox_dir}")
        print("\nType Git commands to practice.")
        print("Type 'help' for suggestions, 'status' for current state, 'exit' to quit.\n")

        while True:
            command = input("git> ").strip()

            if command == 'exit':
                break
            elif command == 'help':
                self.show_command_suggestions()
                continue
            elif command == 'status':
                command = 'git status'
            elif not command:
                continue

            # Add 'git' prefix if not present
            if not command.startswith('git'):
                command = f'git {command}'

            # Analyze command before execution
            feedback = self.analyze_command(command)
            if feedback:
                for item in feedback:
                    print(item)

            # Execute command
            stdout, stderr, returncode = self.run_git_command(command)

            # Display output
            if stdout:
                print(stdout)
            if stderr:
                print(stderr)

            # Provide additional feedback
            if returncode == 0:
                if "Initialized empty Git repository" in stdout:
                    print("✓ Repository initialized! Try: git status")
                elif "nothing to commit" in stdout:
                    print("ℹ Your working directory is clean!")
                elif "Changes to be committed" in stdout:
                    print("✓ You have staged changes ready to commit!")
            else:
                print(f"✗ Command exited with code {returncode}")
                self.suggest_fix(command, stderr)

    def show_command_suggestions(self):
        """Show common Git commands"""
        print("\n" + "="*60)
        print("COMMON GIT COMMANDS")
        print("="*60)
        print("""
Repository Setup:
  git init                 - Initialize repository
  git status               - Check current status

Staging & Committing:
  git add <file>           - Stage a file
  git add .                - Stage all files
  git commit -m "message"  - Commit changes

Viewing History:
  git log                  - View commit history
  git log --oneline        - Compact history
  git diff                 - View changes

Branching:
  git branch               - List branches
  git branch <name>        - Create branch
  git checkout <name>      - Switch branch
  git checkout -b <name>   - Create and switch
  git merge <name>         - Merge branch

Undoing:
  git checkout -- <file>   - Discard changes
  git reset HEAD <file>    - Unstage file
""")

    def suggest_fix(self, command, error):
        """Suggest fixes for common errors"""
        suggestions = []

        if "not a git repository" in error.lower():
            suggestions.append("ℹ Tip: Initialize a repository first with 'git init'")

        if "nothing added to commit" in error.lower():
            suggestions.append("ℹ Tip: Stage files first with 'git add <file>'")

        if "no changes added to commit" in error.lower():
            suggestions.append("ℹ Tip: No files are staged. Use 'git add .' to stage all files")

        if "pathspec" in error.lower() and "did not match" in error.lower():
            suggestions.append("ℹ Tip: The file doesn't exist. Check the filename.")

        if suggestions:
            print("\nSuggestions:")
            for suggestion in suggestions:
                print(suggestion)

    def guided_challenge(self, challenge_num):
        """Guided challenges with step-by-step instructions"""
        challenges = {
            1: {
                "title": "Challenge 1: Create Your First Commit",
                "description": """
Steps:
1. Initialize a Git repository (git init)
2. Create a file called 'hello.txt' with some text
3. Check the status (git status)
4. Stage the file (git add hello.txt)
5. Commit with message "Add hello.txt" (git commit -m "Add hello.txt")
6. View the log (git log)
""",
                "steps": [
                    ("git init", "Repository initialized"),
                    ("echo 'Hello World' > hello.txt", "File created"),
                    ("git status", "Check untracked files"),
                    ("git add hello.txt", "File staged"),
                    ("git commit -m 'Add hello.txt'", "Committed"),
                    ("git log", "View history")
                ]
            },
            2: {
                "title": "Challenge 2: Work with Branches",
                "description": """
Steps:
1. Create a new branch called 'feature' (git checkout -b feature)
2. Create a file 'feature.txt'
3. Commit the file
4. Switch back to main (git checkout main)
5. Merge the feature branch (git merge feature)
6. Delete the feature branch (git branch -d feature)
""",
                "steps": [
                    ("git checkout -b feature", "Branch created"),
                    ("echo 'Feature code' > feature.txt", "File created"),
                    ("git add feature.txt && git commit -m 'Add feature'", "Committed"),
                    ("git checkout main", "Switched to main"),
                    ("git merge feature", "Merged"),
                    ("git branch -d feature", "Branch deleted")
                ]
            },
            3: {
                "title": "Challenge 3: Fix a Mistake",
                "description": """
Steps:
1. Create and commit a file
2. Make a change to the file (don't stage it)
3. View the change with git diff
4. Decide you don't want the change
5. Discard the change (git checkout -- <file>)
6. Verify it's gone
""",
                "steps": [
                    ("echo 'Original' > test.txt && git add test.txt && git commit -m 'Add test'", "Setup"),
                    ("echo 'Mistake' >> test.txt", "Make change"),
                    ("git diff", "View change"),
                    ("git checkout -- test.txt", "Discard change"),
                    ("cat test.txt", "Verify")
                ]
            },
            4: {
                "title": "Challenge 4: Multiple Commits",
                "description": """
Steps:
1. Create 3 different files
2. Commit each file separately with descriptive messages
3. View the log with git log --oneline
4. See the history you've built!
""",
                "steps": [
                    ("echo 'File 1' > file1.txt && git add file1.txt && git commit -m 'Add file1'", "First commit"),
                    ("echo 'File 2' > file2.txt && git add file2.txt && git commit -m 'Add file2'", "Second commit"),
                    ("echo 'File 3' > file3.txt && git add file3.txt && git commit -m 'Add file3'", "Third commit"),
                    ("git log --oneline", "View history")
                ]
            }
        }

        if challenge_num not in challenges:
            print("Invalid challenge number")
            return

        challenge = challenges[challenge_num]
        print("\n" + "="*60)
        print(challenge["title"])
        print("="*60)
        print(challenge["description"])

        input("\nPress Enter to start the challenge...")

        print("\nTry to complete the challenge on your own!")
        print("Type 'hint' for the next step, 'solution' to see all steps, or 'exit' to quit.\n")

        step_index = 0
        while True:
            user_input = input("git> ").strip()

            if user_input == 'exit':
                break
            elif user_input == 'hint':
                if step_index < len(challenge["steps"]):
                    cmd, desc = challenge["steps"][step_index]
                    print(f"\nHint: {desc}")
                    print(f"Try: {cmd}")
                    step_index += 1
                else:
                    print("\nNo more hints - you've seen all steps!")
            elif user_input == 'solution':
                print("\nComplete solution:")
                for cmd, desc in challenge["steps"]:
                    print(f"{desc}:")
                    print(f"  {cmd}\n")
                break
            else:
                # Execute user command
                if not user_input.startswith('git') and not user_input.startswith('echo') and not user_input.startswith('cat'):
                    user_input = f'git {user_input}'

                stdout, stderr, returncode = self.run_git_command(user_input)
                if stdout:
                    print(stdout)
                if stderr:
                    print(stderr)

    def main_menu(self):
        """Main practice menu"""
        while True:
            print("\n" + "="*60)
            print("     GIT PRACTICE WITH CLAUDE")
            print("="*60)
            print(f"\nSandbox: {self.sandbox_dir}\n")
            print("PRACTICE MODES:")
            print("1. Free Practice - Execute any Git commands")
            print("2. Challenge 1 - Create Your First Commit")
            print("3. Challenge 2 - Work with Branches")
            print("4. Challenge 3 - Fix a Mistake")
            print("5. Challenge 4 - Multiple Commits")
            print("6. Reset Sandbox")
            print("0. Exit")

            choice = input("\nChoose an option (0-6): ").strip()

            if choice == '0':
                print("\nGreat practice! Keep learning Git!")
                break
            elif choice == '1':
                self.free_practice()
            elif choice in ['2', '3', '4', '5']:
                self.guided_challenge(int(choice) - 1)
            elif choice == '6':
                self.setup_sandbox()
            else:
                print("Invalid choice!")

if __name__ == "__main__":
    print("="*60)
    print("     GIT PRACTICE WITH CLAUDE")
    print("="*60)
    print("\nWelcome to interactive Git practice!")
    print("You'll practice real Git commands in a safe sandbox environment.\n")

    # Check if git is installed
    try:
        result = subprocess.run(['git', '--version'], capture_output=True)
        if result.returncode != 0:
            print("ERROR: Git is not installed!")
            print("Install it with: sudo apt install git")
            sys.exit(1)
    except FileNotFoundError:
        print("ERROR: Git is not installed!")
        print("Install it with: sudo apt install git")
        sys.exit(1)

    practice = GitPractice()
    practice.main_menu()
