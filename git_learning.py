#!/usr/bin/env python3
"""
GIT LEARNING TUTORIAL
=====================
A beginner-friendly guide to learning Git from scratch.
Complete with 7 lessons and hands-on exercises.

Author: Created for Jason's Git learning journey
Date: 2025-10-19
"""

def lesson_1_intro():
    """Lesson 1: What is Git?"""
    print("\n" + "="*60)
    print("LESSON 1: WHAT IS GIT?")
    print("="*60)

    print("""
Git is a version control system that tracks changes in your code.

KEY CONCEPTS:
- Repository (repo): A project folder tracked by Git
- Commit: A snapshot of your code at a specific point in time
- Branch: A separate line of development
- Remote: A version of your repo hosted online (like GitHub)

WHY USE GIT?
✓ Track all changes to your code
✓ Collaborate with others
✓ Go back to previous versions
✓ Work on features without breaking main code
✓ Backup your work online

BASIC WORKFLOW:
1. Make changes to files
2. Stage changes (prepare them for commit)
3. Commit changes (save snapshot)
4. Push to remote (backup online)
""")

    print("\nEXERCISE 1.1:")
    print("Think about why version control is useful.")
    print("Imagine you break your code - with Git, you can go back!")
    input("\nPress Enter when ready to continue...")

def lesson_2_setup():
    """Lesson 2: Git Setup and Configuration"""
    print("\n" + "="*60)
    print("LESSON 2: GIT SETUP AND CONFIGURATION")
    print("="*60)

    print("""
INSTALLING GIT:
- Linux: sudo apt install git
- Mac: brew install git
- Windows: Download from git-scm.com

INITIAL CONFIGURATION:
These commands set your identity for all commits.

Command:
  git config --global user.name "Your Name"
  git config --global user.email "your.email@example.com"

CHECK YOUR CONFIG:
  git config --list

EXAMPLE:
  $ git config --global user.name "Jason"
  $ git config --global user.email "jason@example.com"
  $ git config --list
  user.name=Jason
  user.email=jason@example.com
""")

    print("\nEXERCISE 2.1:")
    print("Run these commands in your terminal:")
    print("  git config --global user.name \"Your Name\"")
    print("  git config --global user.email \"your@email.com\"")
    print("  git config --list")
    input("\nPress Enter when you've completed this...")

def lesson_3_basic_commands():
    """Lesson 3: Basic Git Commands"""
    print("\n" + "="*60)
    print("LESSON 3: BASIC GIT COMMANDS")
    print("="*60)

    print("""
CREATING A REPOSITORY:

1. INITIALIZE A NEW REPO:
   git init
   - Creates a new Git repository in current folder
   - Creates hidden .git folder

2. CLONE AN EXISTING REPO:
   git clone <url>
   - Downloads a copy of a remote repository
   - Example: git clone https://github.com/user/repo.git

CHECKING STATUS:
   git status
   - Shows which files are modified, staged, or untracked
   - Use this command ALL THE TIME!

EXAMPLE WORKFLOW:
  $ mkdir my_project
  $ cd my_project
  $ git init
  Initialized empty Git repository in /path/to/my_project/.git/

  $ git status
  On branch main
  No commits yet
  nothing to commit (create/copy files and use "git add")
""")

    print("\nEXERCISE 3.1:")
    print("Try these commands:")
    print("  mkdir test_repo")
    print("  cd test_repo")
    print("  git init")
    print("  git status")
    input("\nPress Enter when done...")

def lesson_4_staging_committing():
    """Lesson 4: Staging and Committing"""
    print("\n" + "="*60)
    print("LESSON 4: STAGING AND COMMITTING")
    print("="*60)

    print("""
THE THREE STATES OF FILES:
1. Modified: Changed but not staged
2. Staged: Ready to be committed
3. Committed: Safely stored in Git history

STAGING FILES:
  git add <filename>      # Stage a specific file
  git add .               # Stage all changes
  git add *.py            # Stage all .py files

COMMITTING FILES:
  git commit -m "Your message here"
  - Saves staged changes with a descriptive message
  - Message should describe WHAT and WHY

GOOD COMMIT MESSAGES:
  ✓ "Add user login feature"
  ✓ "Fix bug in password validation"
  ✓ "Update README with installation steps"

  ✗ "fixed stuff"
  ✗ "changes"
  ✗ "asdf"

EXAMPLE WORKFLOW:
  $ echo "print('Hello Git')" > test.py
  $ git status
  Untracked files:
    test.py

  $ git add test.py
  $ git status
  Changes to be committed:
    new file: test.py

  $ git commit -m "Add initial test.py file"
  [main 1a2b3c4] Add initial test.py file
   1 file changed, 1 insertion(+)
""")

    print("\nEXERCISE 4.1:")
    print("In your test_repo:")
    print("  1. Create a file: echo 'Hello' > hello.txt")
    print("  2. Check status: git status")
    print("  3. Stage it: git add hello.txt")
    print("  4. Commit it: git commit -m \"Add hello.txt\"")
    input("\nPress Enter when done...")

def lesson_5_viewing_history():
    """Lesson 5: Viewing History and Changes"""
    print("\n" + "="*60)
    print("LESSON 5: VIEWING HISTORY AND CHANGES")
    print("="*60)

    print("""
VIEW COMMIT HISTORY:
  git log
  - Shows all commits from newest to oldest
  - Press 'q' to quit the log view

  git log --oneline
  - Compact view (one line per commit)

  git log --graph --oneline --all
  - Visual representation of branches

VIEW CHANGES:
  git diff
  - Shows unstaged changes

  git diff --staged
  - Shows staged changes

  git show <commit-hash>
  - Shows details of a specific commit

EXAMPLE:
  $ git log --oneline
  1a2b3c4 Add hello.txt
  5e6f7g8 Initial commit

  $ git diff
  diff --git a/hello.txt b/hello.txt
  index 802992c..d95f3ad 100644
  --- a/hello.txt
  +++ b/hello.txt
  @@ -1 +1 @@
  -Hello
  +Hello Git!
""")

    print("\nEXERCISE 5.1:")
    print("Try these commands:")
    print("  git log")
    print("  git log --oneline")
    print("  git diff")
    input("\nPress Enter when done...")

def lesson_6_branching():
    """Lesson 6: Branching and Merging"""
    print("\n" + "="*60)
    print("LESSON 6: BRANCHING AND MERGING")
    print("="*60)

    print("""
WHAT IS A BRANCH?
A branch is a separate line of development.
Main branch = main or master (your primary code)
Feature branches = new features you're working on

BRANCH COMMANDS:
  git branch
  - List all branches (* shows current branch)

  git branch <branch-name>
  - Create a new branch

  git checkout <branch-name>
  - Switch to a branch

  git checkout -b <branch-name>
  - Create AND switch to new branch (shortcut!)

  git merge <branch-name>
  - Merge a branch into current branch

  git branch -d <branch-name>
  - Delete a branch

EXAMPLE WORKFLOW:
  $ git branch
  * main

  $ git checkout -b feature-login
  Switched to a new branch 'feature-login'

  $ git branch
    main
  * feature-login

  # ... make changes and commit ...

  $ git checkout main
  $ git merge feature-login
  Updating 1a2b3c4..5e6f7g8
  Fast-forward
   login.py | 10 ++++++++++

  $ git branch -d feature-login
  Deleted branch feature-login
""")

    print("\nEXERCISE 6.1:")
    print("Practice branching:")
    print("  git checkout -b test-branch")
    print("  git branch  # See all branches")
    print("  git checkout main  # Switch back")
    print("  git branch -d test-branch  # Delete it")
    input("\nPress Enter when done...")

def lesson_7_remote_repos():
    """Lesson 7: Working with Remote Repositories"""
    print("\n" + "="*60)
    print("LESSON 7: WORKING WITH REMOTE REPOSITORIES")
    print("="*60)

    print("""
WHAT IS A REMOTE?
A remote is a version of your repository hosted online
(GitHub, GitLab, Bitbucket, etc.)

REMOTE COMMANDS:
  git remote add origin <url>
  - Connect your local repo to a remote

  git remote -v
  - View all remotes

  git push origin <branch-name>
  - Upload your commits to remote

  git push -u origin main
  - Upload and set upstream (first time)

  git pull
  - Download changes from remote

  git fetch
  - Download changes but don't merge

TYPICAL WORKFLOW:
  # First time setup
  $ git remote add origin https://github.com/user/repo.git
  $ git push -u origin main

  # Daily workflow
  $ git add .
  $ git commit -m "Add new feature"
  $ git push

  # Get latest changes
  $ git pull

GITHUB WORKFLOW:
  1. Create repo on GitHub
  2. Copy the URL
  3. git remote add origin <url>
  4. git push -u origin main
  5. Now you can push/pull anytime!
""")

    print("\nEXERCISE 7.1:")
    print("Understanding remotes:")
    print("  git remote -v  # See your remotes")
    print("  # If you have a GitHub account, try:")
    print("  # 1. Create a repo on GitHub")
    print("  # 2. git remote add origin <url>")
    print("  # 3. git push -u origin main")
    input("\nPress Enter when done...")

def review_quiz():
    """Quick review quiz"""
    print("\n" + "="*60)
    print("QUICK REVIEW QUIZ")
    print("="*60)

    questions = [
        {
            "question": "What command initializes a new Git repository?",
            "answer": "git init"
        },
        {
            "question": "What command stages all files?",
            "answer": "git add ."
        },
        {
            "question": "What command creates a commit?",
            "answer": "git commit -m"
        },
        {
            "question": "What command shows the status of your repo?",
            "answer": "git status"
        },
        {
            "question": "What command creates a new branch and switches to it?",
            "answer": "git checkout -b"
        }
    ]

    print("\nAnswer these questions (just think about them):\n")
    for i, q in enumerate(questions, 1):
        print(f"{i}. {q['question']}")
        input("   Press Enter to see answer...")
        print(f"   Answer: {q['answer']}\n")

def main():
    """Main tutorial function"""
    print("="*60)
    print("        WELCOME TO GIT LEARNING TUTORIAL")
    print("="*60)
    print("\nThis tutorial will teach you Git basics from scratch.")
    print("You'll learn 7 essential lessons about version control.\n")

    lessons = [
        ("What is Git?", lesson_1_intro),
        ("Git Setup and Configuration", lesson_2_setup),
        ("Basic Git Commands", lesson_3_basic_commands),
        ("Staging and Committing", lesson_4_staging_committing),
        ("Viewing History and Changes", lesson_5_viewing_history),
        ("Branching and Merging", lesson_6_branching),
        ("Working with Remote Repositories", lesson_7_remote_repos),
    ]

    while True:
        print("\n" + "="*60)
        print("LESSONS:")
        print("="*60)
        for i, (title, _) in enumerate(lessons, 1):
            print(f"{i}. {title}")
        print("8. Review Quiz")
        print("0. Exit")

        choice = input("\nChoose a lesson (0-8): ").strip()

        if choice == '0':
            print("\nGreat work! Keep practicing Git!")
            print("Remember: git status is your best friend!")
            break
        elif choice == '8':
            review_quiz()
        elif choice.isdigit() and 1 <= int(choice) <= len(lessons):
            _, lesson_func = lessons[int(choice) - 1]
            lesson_func()
        else:
            print("Invalid choice. Please enter 0-8.")

if __name__ == "__main__":
    main()
