#!/usr/bin/env python3
"""
GIT PRACTICE EXERCISES
======================
6 hands-on projects to practice your Git skills.
Complete these in order to master Git fundamentals.

Author: Created for Jason's Git learning journey
Date: 2025-10-19
"""

import json
import os
from datetime import datetime

def exercise_1():
    """Exercise 1: Create Your First Repository"""
    print("\n" + "="*60)
    print("EXERCISE 1: CREATE YOUR FIRST REPOSITORY")
    print("="*60)
    print("""
GOAL: Create a new Git repository and make your first commit.

STEPS:
1. Create a new folder called 'my_first_repo'
2. Navigate into it
3. Initialize a Git repository
4. Create a README.md file with your name
5. Stage the file
6. Commit with message "Initial commit with README"
7. Check your git log

COMMANDS TO USE:
  mkdir my_first_repo
  cd my_first_repo
  git init
  echo "# My First Repository" > README.md
  echo "Created by: [Your Name]" >> README.md
  git add README.md
  git commit -m "Initial commit with README"
  git log

EXPECTED RESULT:
- A new Git repository
- One commit in your history
- A README.md file tracked by Git

CHALLENGE:
Add a second file called 'hello.txt' and commit it separately.
""")

def exercise_2():
    """Exercise 2: Practice Staging and Committing"""
    print("\n" + "="*60)
    print("EXERCISE 2: PRACTICE STAGING AND COMMITTING")
    print("="*60)
    print("""
GOAL: Practice making multiple commits with good messages.

STEPS:
1. In your my_first_repo folder, create 3 new files:
   - main.py (with print("Hello World"))
   - utils.py (with def helper(): pass)
   - config.py (with API_KEY = "test")

2. Stage and commit them ONE AT A TIME with descriptive messages

3. Make a change to main.py (add another print statement)

4. View your changes with git diff

5. Stage and commit the change

6. View your commit history

COMMANDS TO USE:
  echo 'print("Hello World")' > main.py
  git add main.py
  git commit -m "Add main.py with hello world"

  echo 'def helper(): pass' > utils.py
  git add utils.py
  git commit -m "Add utils.py with helper function"

  echo 'API_KEY = "test"' > config.py
  git add config.py
  git commit -m "Add config.py with API key"

  echo 'print("Goodbye")' >> main.py
  git diff
  git add main.py
  git commit -m "Add goodbye message to main.py"

  git log --oneline

EXPECTED RESULT:
- At least 4 commits
- Each commit has a descriptive message
- Files are properly tracked

CHALLENGE:
Practice using 'git add .' to stage multiple files at once.
""")

def exercise_3():
    """Exercise 3: Branching Basics"""
    print("\n" + "="*60)
    print("EXERCISE 3: BRANCHING BASICS")
    print("="*60)
    print("""
GOAL: Create branches, make changes, and merge them.

SCENARIO:
You're working on a project and want to add a new feature
without affecting your main code.

STEPS:
1. Create and switch to a new branch called 'feature-add-function'

2. In this branch, add a new file 'calculator.py' with:
   def add(a, b):
       return a + b

3. Commit this change

4. Switch back to main branch

5. Notice calculator.py is not there!

6. Merge the feature branch into main

7. Now calculator.py should be in main

8. Delete the feature branch

COMMANDS TO USE:
  git checkout -b feature-add-function
  echo 'def add(a, b):' > calculator.py
  echo '    return a + b' >> calculator.py
  git add calculator.py
  git commit -m "Add calculator with add function"

  git checkout main
  ls  # calculator.py not here!

  git merge feature-add-function
  ls  # calculator.py is here now!

  git branch -d feature-add-function
  git branch  # should only show main

EXPECTED RESULT:
- Feature branch created and merged
- calculator.py exists in main
- Feature branch deleted
- Clean commit history

CHALLENGE:
Create another branch 'feature-subtract', add a subtract
function, and merge it into main.
""")

def exercise_4():
    """Exercise 4: Fixing Mistakes"""
    print("\n" + "="*60)
    print("EXERCISE 4: FIXING MISTAKES")
    print("="*60)
    print("""
GOAL: Learn to undo changes and fix mistakes.

SCENARIO:
You made changes you don't want to keep.

STEPS:
1. Make a change to README.md (add some text)

2. View the change with git diff

3. Decide you don't want this change

4. Discard the change (unstaged changes)

5. Make another change and stage it

6. Decide you don't want this staged change either

7. Unstage it

COMMANDS TO USE:
  echo "Unwanted change" >> README.md
  git diff
  git checkout -- README.md  # Discard changes
  cat README.md  # Change is gone!

  echo "Another unwanted change" >> README.md
  git add README.md
  git status  # It's staged
  git reset HEAD README.md  # Unstage it
  git status  # It's unstaged now
  git checkout -- README.md  # Discard it

IMPORTANT COMMANDS:
  git checkout -- <file>     # Discard unstaged changes
  git reset HEAD <file>      # Unstage changes
  git reset --soft HEAD~1    # Undo last commit (keep changes)
  git reset --hard HEAD~1    # Undo last commit (DELETE changes)

EXPECTED RESULT:
- Understand how to undo changes
- Know the difference between staged and unstaged
- Feel confident fixing mistakes

CHALLENGE:
Make a commit with a typo, then use git reset --soft HEAD~1
to undo it, fix the message, and commit again.
""")

def exercise_5():
    """Exercise 5: Working with Remote Repositories"""
    print("\n" + "="*60)
    print("EXERCISE 5: WORKING WITH REMOTE REPOSITORIES")
    print("="*60)
    print("""
GOAL: Connect your local repo to GitHub and push your code.

PREREQUISITES:
- You need a GitHub account (free at github.com)

STEPS:
1. Go to github.com and create a new repository
   - Name it 'git-practice'
   - Don't initialize with README (we already have one)

2. Copy the repository URL

3. Add the remote to your local repo

4. Push your code to GitHub

5. Refresh GitHub page and see your code!

6. Make a change locally and push again

COMMANDS TO USE:
  # After creating repo on GitHub and copying URL:
  git remote add origin https://github.com/yourusername/git-practice.git
  git remote -v  # Verify remote was added

  git branch -M main  # Rename to main if needed
  git push -u origin main  # Push to GitHub

  # Make a change
  echo "Updated from local" >> README.md
  git add README.md
  git commit -m "Update README from local"
  git push  # Push again

EXPECTED RESULT:
- Your code is on GitHub
- You can push and pull changes
- Your backup is in the cloud!

CHALLENGE:
Clone your repository to a different folder to simulate
working from another computer.
""")

def exercise_6():
    """Exercise 6: Complete Workflow Project"""
    print("\n" + "="*60)
    print("EXERCISE 6: COMPLETE WORKFLOW PROJECT")
    print("="*60)
    print("""
GOAL: Build a small project using proper Git workflow.

PROJECT: Create a simple Python calculator with Git

REQUIREMENTS:
1. Create a new repo called 'calculator-project'
2. Initialize Git
3. Create main.py with basic structure
4. Commit: "Initial calculator structure"
5. Create branch 'feature-addition'
6. Add addition function
7. Commit: "Add addition function"
8. Merge to main
9. Create branch 'feature-subtraction'
10. Add subtraction function
11. Commit: "Add subtraction function"
12. Merge to main
13. Update README.md with usage instructions
14. Commit: "Add documentation"
15. Push everything to GitHub

COMPLETE WORKFLOW:
  mkdir calculator-project
  cd calculator-project
  git init

  # Create initial file
  cat > main.py << 'EOF'
def calculator():
    print("Simple Calculator")

if __name__ == "__main__":
    calculator()
EOF

  git add main.py
  git commit -m "Initial calculator structure"

  # Add addition feature
  git checkout -b feature-addition
  # Edit main.py to add addition function
  git add main.py
  git commit -m "Add addition function"
  git checkout main
  git merge feature-addition
  git branch -d feature-addition

  # Add subtraction feature
  git checkout -b feature-subtraction
  # Edit main.py to add subtraction function
  git add main.py
  git commit -m "Add subtraction function"
  git checkout main
  git merge feature-subtraction
  git branch -d feature-subtraction

  # Add README
  echo "# Calculator Project" > README.md
  echo "Simple calculator with add/subtract" >> README.md
  git add README.md
  git commit -m "Add documentation"

  # Push to GitHub
  # (Create repo on GitHub first)
  git remote add origin <your-github-url>
  git push -u origin main

EXPECTED RESULT:
- Complete project with multiple features
- Clean commit history
- Proper branch usage
- Documentation
- Code on GitHub

CHALLENGE:
Add multiplication and division on separate branches,
then merge them all into main.
""")

def load_or_create_user():
    """Load existing user or create new user"""
    progress_file = "git_exercises_progress.json"
    user_name = ""
    exercises_completed = []

    if os.path.exists(progress_file):
        try:
            with open(progress_file, 'r') as f:
                data = json.load(f)
                user_name = data.get('user_name', '')
                exercises_completed = data.get('exercises_completed', [])
                print(f"\nWelcome back, {user_name}!")
                print(f"Exercises completed: {len(exercises_completed)}/6")
                input("\nPress Enter to continue...")
        except:
            user_name = create_new_user()
    else:
        user_name = create_new_user()

    return user_name, exercises_completed, progress_file

def create_new_user():
    """Create new user profile"""
    print("\n" + "="*60)
    print("WELCOME TO GIT PRACTICE EXERCISES!")
    print("="*60)
    user_name = input("\nPlease enter your name: ").strip()
    if not user_name:
        user_name = "Student"
    print(f"\nHello, {user_name}! Let's practice Git!")
    input("\nPress Enter to begin...")
    return user_name

def save_progress(user_name, exercises_completed, progress_file):
    """Save user progress"""
    try:
        data = {
            'user_name': user_name,
            'exercises_completed': exercises_completed,
            'last_session': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        }
        with open(progress_file, 'w') as f:
            json.dump(data, f, indent=2)
        print(f"\n✓ Progress saved for {user_name}!")
    except Exception as e:
        print(f"\n✗ Error saving progress: {e}")

def clear_progress(progress_file):
    """Clear saved progress"""
    try:
        if os.path.exists(progress_file):
            os.remove(progress_file)
            print("\n✓ Progress cleared!")
    except Exception as e:
        print(f"\n✗ Error clearing progress: {e}")

def show_progress(user_name, exercises_completed):
    """Show exercise progress"""
    print("\n" + "="*60)
    print(f"YOUR PROGRESS - {user_name}")
    print("="*60)

    exercises = [
        "Exercise 1: Create Your First Repository",
        "Exercise 2: Practice Staging and Committing",
        "Exercise 3: Branching Basics",
        "Exercise 4: Fixing Mistakes",
        "Exercise 5: Working with Remote Repositories",
        "Exercise 6: Complete Workflow Project"
    ]

    print()
    for i, title in enumerate(exercises, 1):
        status = "✓" if i in exercises_completed else "☐"
        print(f"{status} {title}")

    completed = len(exercises_completed)
    total = len(exercises)
    percentage = (completed / total) * 100 if total > 0 else 0

    print(f"\nCompleted: {completed}/{total} ({percentage:.0f}%)")
    input("\nPress Enter to continue...")

def main():
    """Main exercise menu"""
    print("="*60)
    print("        GIT PRACTICE EXERCISES")
    print("="*60)
    print("\nComplete these exercises to master Git!")
    print("Work through them in order.\n")

    user_name, exercises_completed, progress_file = load_or_create_user()

    exercises = [
        ("Create Your First Repository", exercise_1),
        ("Practice Staging and Committing", exercise_2),
        ("Branching Basics", exercise_3),
        ("Fixing Mistakes", exercise_4),
        ("Working with Remote Repositories", exercise_5),
        ("Complete Workflow Project", exercise_6),
    ]

    while True:
        print("\n" + "="*60)
        print(f"EXERCISES - {user_name}")
        print("="*60)
        for i, (title, _) in enumerate(exercises, 1):
            completed = "✓" if i in exercises_completed else " "
            print(f"[{completed}] {i}. {title}")
        print("7. Show Progress")
        print("8. Mark Exercise as Complete")
        print("0. Exit")

        choice = input("\nChoose an exercise (0-8): ").strip()

        if choice == '0':
            print("\n" + "="*60)
            print("SAVE PROGRESS")
            print("="*60)
            save_choice = input("\nDo you want to save your progress? (yes/no): ").strip().lower()

            if save_choice in ['yes', 'y']:
                save_progress(user_name, exercises_completed, progress_file)
                print(f"\nKeep practicing, {user_name}! Git gets easier with use!")
            else:
                clear_progress(progress_file)
                print(f"\nProgress cleared. See you next time, {user_name}!")
            break
        elif choice == '7':
            show_progress(user_name, exercises_completed)
        elif choice == '8':
            ex_num = input("Which exercise did you complete? (1-6): ").strip()
            if ex_num.isdigit() and 1 <= int(ex_num) <= 6:
                ex_num = int(ex_num)
                if ex_num not in exercises_completed:
                    exercises_completed.append(ex_num)
                    print(f"\n✓ Exercise {ex_num} marked as complete!")
                else:
                    print(f"\nExercise {ex_num} was already completed!")
            else:
                print("Invalid exercise number!")
            input("\nPress Enter to continue...")
        elif choice.isdigit() and 1 <= int(choice) <= len(exercises):
            _, exercise_func = exercises[int(choice) - 1]
            exercise_func()
        else:
            print("Invalid choice. Please enter 0-8.")

if __name__ == "__main__":
    main()
