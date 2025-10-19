#!/usr/bin/env python3
"""
INTERACTIVE GIT LESSON
======================
Enhanced interactive Git learning experience with hands-on practice.
Features live command execution, guided challenges, and instant feedback.

Author: Created for Jason's Git learning journey
Date: 2025-10-19
"""

import subprocess
import os
import sys

class GitLesson:
    """Interactive Git lesson with command execution and feedback"""

    def __init__(self):
        self.lesson_progress = {i: False for i in range(1, 8)}
        self.practice_dir = os.path.join(os.getcwd(), "git_practice_sandbox")

    def clear_screen(self):
        """Clear terminal screen"""
        os.system('clear' if os.name == 'posix' else 'cls')

    def run_command(self, command, cwd=None):
        """Execute a shell command and return output"""
        try:
            result = subprocess.run(
                command,
                shell=True,
                capture_output=True,
                text=True,
                cwd=cwd
            )
            return result.stdout + result.stderr, result.returncode
        except Exception as e:
            return f"Error: {str(e)}", 1

    def check_git_installed(self):
        """Check if Git is installed"""
        output, returncode = self.run_command("git --version")
        if returncode != 0:
            print("ERROR: Git is not installed!")
            print("Install it with: sudo apt install git")
            return False
        print(f"✓ {output.strip()}")
        return True

    def lesson_1_intro(self):
        """Lesson 1: Introduction to Git"""
        self.clear_screen()
        print("="*60)
        print("LESSON 1: INTRODUCTION TO GIT")
        print("="*60)
        print("""
What is Git?
Git is a distributed version control system that tracks changes
in your code over time.

Key Benefits:
✓ Track all changes to your files
✓ Collaborate with others
✓ Revert to previous versions
✓ Work on features independently
✓ Backup code in the cloud

Basic Git Workflow:
  Working Directory → Staging Area → Repository → Remote

  1. Modify files (working directory)
  2. Stage changes (git add)
  3. Commit changes (git commit)
  4. Push to remote (git push)
""")

        input("\nPress Enter to check your Git installation...")

        if self.check_git_installed():
            print("\n✓ Git is ready to use!")
            self.lesson_progress[1] = True
        else:
            print("\n✗ Please install Git before continuing")

        input("\nPress Enter to continue...")

    def lesson_2_setup(self):
        """Lesson 2: Git Configuration"""
        self.clear_screen()
        print("="*60)
        print("LESSON 2: GIT CONFIGURATION")
        print("="*60)
        print("""
Before using Git, you need to configure your identity.
This information appears in every commit you make.

Commands:
  git config --global user.name "Your Name"
  git config --global user.email "your@email.com"
  git config --list
""")

        print("\nLet's check your current Git configuration:")
        output, _ = self.run_command("git config --list | grep user")
        if output:
            print(output)
            print("✓ Git is already configured!")
        else:
            print("Git is not configured yet.")
            print("\nRun these commands in your terminal:")
            print('  git config --global user.name "Your Name"')
            print('  git config --global user.email "your@email.com"')

        self.lesson_progress[2] = True
        input("\nPress Enter to continue...")

    def lesson_3_init_repo(self):
        """Lesson 3: Creating a Repository"""
        self.clear_screen()
        print("="*60)
        print("LESSON 3: CREATING A REPOSITORY")
        print("="*60)
        print("""
A repository (repo) is a project folder tracked by Git.

Commands:
  git init          - Initialize a new repo
  git status        - Check repo status
  git clone <url>   - Clone existing repo

Let's create a practice repository!
""")

        # Create practice directory
        if not os.path.exists(self.practice_dir):
            os.makedirs(self.practice_dir)
            print(f"✓ Created directory: {self.practice_dir}")

        print(f"\nInitializing Git repository in {self.practice_dir}...")
        output, returncode = self.run_command("git init", cwd=self.practice_dir)

        if returncode == 0:
            print("✓ Repository initialized successfully!")
            print(output)

            # Check status
            print("\nRunning 'git status'...")
            output, _ = self.run_command("git status", cwd=self.practice_dir)
            print(output)

            self.lesson_progress[3] = True
        else:
            print("✗ Failed to initialize repository")
            print(output)

        input("\nPress Enter to continue...")

    def lesson_4_staging_committing(self):
        """Lesson 4: Staging and Committing"""
        self.clear_screen()
        print("="*60)
        print("LESSON 4: STAGING AND COMMITTING")
        print("="*60)
        print("""
The Git Workflow:
  Modified → Staged → Committed

Commands:
  git add <file>           - Stage a file
  git add .                - Stage all files
  git commit -m "message"  - Commit staged changes
  git log                  - View commit history

Let's practice!
""")

        # Create a sample file
        readme_path = os.path.join(self.practice_dir, "README.md")
        with open(readme_path, "w") as f:
            f.write("# Git Practice Repository\n\nLearning Git with Claude!\n")

        print("✓ Created README.md")

        print("\nChecking status...")
        output, _ = self.run_command("git status", cwd=self.practice_dir)
        print(output)

        input("Press Enter to stage the file...")

        print("\nStaging README.md...")
        output, _ = self.run_command("git add README.md", cwd=self.practice_dir)
        print("✓ File staged")

        print("\nChecking status again...")
        output, _ = self.run_command("git status", cwd=self.practice_dir)
        print(output)

        input("Press Enter to commit...")

        print("\nCommitting changes...")
        output, _ = self.run_command(
            'git commit -m "Initial commit with README"',
            cwd=self.practice_dir
        )
        print(output)

        print("\nViewing commit history...")
        output, _ = self.run_command("git log --oneline", cwd=self.practice_dir)
        print(output)

        self.lesson_progress[4] = True
        input("\nPress Enter to continue...")

    def lesson_5_branching(self):
        """Lesson 5: Branching"""
        self.clear_screen()
        print("="*60)
        print("LESSON 5: BRANCHING")
        print("="*60)
        print("""
Branches let you work on features independently.

Commands:
  git branch              - List branches
  git branch <name>       - Create branch
  git checkout <name>     - Switch branch
  git checkout -b <name>  - Create and switch
  git merge <name>        - Merge branch

Let's create a feature branch!
""")

        print("\nCurrent branches:")
        output, _ = self.run_command("git branch", cwd=self.practice_dir)
        print(output if output.strip() else "* main")

        input("Press Enter to create a new branch...")

        print("\nCreating branch 'feature-hello'...")
        output, _ = self.run_command(
            "git checkout -b feature-hello",
            cwd=self.practice_dir
        )
        print(output)

        # Create a file in the feature branch
        hello_path = os.path.join(self.practice_dir, "hello.py")
        with open(hello_path, "w") as f:
            f.write('print("Hello from feature branch!")\n')

        print("✓ Created hello.py")

        print("\nStaging and committing...")
        self.run_command("git add hello.py", cwd=self.practice_dir)
        output, _ = self.run_command(
            'git commit -m "Add hello.py"',
            cwd=self.practice_dir
        )
        print(output)

        input("Press Enter to merge into main...")

        print("\nSwitching to main branch...")
        output, _ = self.run_command("git checkout main", cwd=self.practice_dir)
        print(output)

        print("\nMerging feature-hello...")
        output, _ = self.run_command("git merge feature-hello", cwd=self.practice_dir)
        print(output)

        print("\n✓ Feature branch merged successfully!")
        self.lesson_progress[5] = True

        input("\nPress Enter to continue...")

    def lesson_6_viewing_changes(self):
        """Lesson 6: Viewing Changes and History"""
        self.clear_screen()
        print("="*60)
        print("LESSON 6: VIEWING CHANGES AND HISTORY")
        print("="*60)
        print("""
View what's changed in your repository.

Commands:
  git status            - Current status
  git diff              - Unstaged changes
  git diff --staged     - Staged changes
  git log               - Commit history
  git log --oneline     - Compact history
  git log --graph       - Visual graph

Let's explore your repository!
""")

        print("\nCurrent status:")
        output, _ = self.run_command("git status", cwd=self.practice_dir)
        print(output)

        print("\nCommit history:")
        output, _ = self.run_command("git log --oneline --graph", cwd=self.practice_dir)
        print(output)

        # Make a change to demonstrate diff
        readme_path = os.path.join(self.practice_dir, "README.md")
        with open(readme_path, "a") as f:
            f.write("\n## Progress\nLearning Git step by step!\n")

        print("\nMade changes to README.md")
        print("\nUnstaged changes (git diff):")
        output, _ = self.run_command("git diff", cwd=self.practice_dir)
        print(output[:500] if output else "No output")

        self.lesson_progress[6] = True
        input("\nPress Enter to continue...")

    def lesson_7_undoing_changes(self):
        """Lesson 7: Undoing Changes"""
        self.clear_screen()
        print("="*60)
        print("LESSON 7: UNDOING CHANGES")
        print("="*60)
        print("""
Mistakes happen! Git helps you fix them.

Commands:
  git checkout -- <file>    - Discard unstaged changes
  git reset HEAD <file>     - Unstage file
  git reset --soft HEAD~1   - Undo commit (keep changes)
  git reset --hard HEAD~1   - Undo commit (delete changes)

CAUTION: --hard deletes changes permanently!
""")

        print("\nCurrent status:")
        output, _ = self.run_command("git status", cwd=self.practice_dir)
        print(output)

        if "modified" in output.lower():
            print("\nYou have unstaged changes.")
            response = input("Discard them? (y/n): ").lower()

            if response == 'y':
                print("\nDiscarding changes...")
                output, _ = self.run_command(
                    "git checkout -- .",
                    cwd=self.practice_dir
                )
                print("✓ Changes discarded")

                print("\nNew status:")
                output, _ = self.run_command("git status", cwd=self.practice_dir)
                print(output)

        self.lesson_progress[7] = True
        input("\nPress Enter to continue...")

    def guided_challenge(self):
        """Guided coding challenge"""
        self.clear_screen()
        print("="*60)
        print("GUIDED CHALLENGE: BUILD A PROJECT")
        print("="*60)
        print("""
Challenge: Create a calculator project with proper Git workflow

Steps:
1. Create a new directory 'calculator'
2. Initialize Git
3. Create main.py with a calculator function
4. Commit: "Initial calculator setup"
5. Create branch 'feature-add'
6. Add an addition function
7. Commit and merge to main
8. Create branch 'feature-subtract'
9. Add a subtraction function
10. Commit and merge to main

Try it yourself! I'll guide you if needed.
""")

        input("Press Enter when you want to start or skip...")

    def show_progress(self):
        """Show learning progress"""
        self.clear_screen()
        print("="*60)
        print("YOUR LEARNING PROGRESS")
        print("="*60)

        lessons = [
            "Introduction to Git",
            "Git Configuration",
            "Creating a Repository",
            "Staging and Committing",
            "Branching",
            "Viewing Changes and History",
            "Undoing Changes"
        ]

        for i, lesson in enumerate(lessons, 1):
            status = "✓" if self.lesson_progress[i] else "☐"
            print(f"{status} Lesson {i}: {lesson}")

        completed = sum(self.lesson_progress.values())
        total = len(self.lesson_progress)
        percentage = (completed / total) * 100

        print(f"\nCompleted: {completed}/{total} ({percentage:.0f}%)")
        input("\nPress Enter to continue...")

    def main_menu(self):
        """Main interactive menu"""
        while True:
            self.clear_screen()
            print("="*60)
            print("     INTERACTIVE GIT LESSON")
            print("="*60)
            print("\nLearn Git with hands-on practice!\n")

            print("LESSONS:")
            print("1. Introduction to Git")
            print("2. Git Configuration")
            print("3. Creating a Repository")
            print("4. Staging and Committing")
            print("5. Branching")
            print("6. Viewing Changes and History")
            print("7. Undoing Changes")
            print()
            print("8. Guided Challenge")
            print("9. Show Progress")
            print("0. Exit")

            choice = input("\nChoose an option (0-9): ").strip()

            if choice == '0':
                print("\nGreat work! Keep practicing Git!")
                break
            elif choice == '1':
                self.lesson_1_intro()
            elif choice == '2':
                self.lesson_2_setup()
            elif choice == '3':
                self.lesson_3_init_repo()
            elif choice == '4':
                self.lesson_4_staging_committing()
            elif choice == '5':
                self.lesson_5_branching()
            elif choice == '6':
                self.lesson_6_viewing_changes()
            elif choice == '7':
                self.lesson_7_undoing_changes()
            elif choice == '8':
                self.guided_challenge()
            elif choice == '9':
                self.show_progress()
            else:
                print("Invalid choice!")
                input("Press Enter to continue...")

if __name__ == "__main__":
    lesson = GitLesson()
    lesson.main_menu()
