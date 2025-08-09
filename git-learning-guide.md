# Git Learning Guide - Progressive Practice Module

## Learning Progress Tracker
**Started:** Today  
**Current Level:** Beginner  
**Goal:** Master Git fundamentals from local to remote operations

## Module Structure Overview
This guide is broken into 6 progressive levels, each building on the previous:

### ✅ Planning Phase - COMPLETED
- [x] Analyze all-git-commands.txt reference file
- [x] Design progressive learning structure
- [x] Create tracking system

### 🔲 Level 1: Repository Basics (PENDING)
**Commands to learn:** `init`, `status`
**Skills:** Creating repositories, checking status

### 🔲 Level 2: File Management (PENDING)  
**Commands to learn:** `add`, `commit`, `rm`, `mv`
**Skills:** Staging changes, making commits, file operations

### 🔲 Level 3: History & Inspection (PENDING)
**Commands to learn:** `log`, `show`, `diff`, `blame`
**Skills:** Viewing commit history, examining changes

### 🔲 Level 4: Branching & Navigation (PENDING)
**Commands to learn:** `branch`, `checkout`, `switch`
**Skills:** Creating branches, switching between branches

### 🔲 Level 5: Merging & Advanced Local (PENDING)
**Commands to learn:** `merge`, `rebase`, `stash`, `reset`
**Skills:** Combining branches, managing work-in-progress

### 🔲 Level 6: Remote Operations (PENDING)
**Commands to learn:** `clone`, `push`, `pull`, `fetch`, `remote`
**Skills:** Working with remote repositories, collaboration

---

## Reference Commands from all-git-commands.txt

### Main Porcelain Commands (Essential for Learning)
From lines 3-47 of all-git-commands.txt, organized by learning priority:

- **Repository Management**
  - `init` (line 24) - Create an empty Git repository
  - `clone` (line 14) - Clone a repository into a new directory
  - `status` (line 43) - Show the working tree status

- **File Operations**  
  - `add` (line 4) - Add file contents to the index
  - `commit` (line 15) - Record changes to the repository
  - `rm` (line 37) - Remove files from working tree and index
  - `mv` (line 28) - Move or rename files

- **History & Inspection**
  - `log` (line 25) - Show commit logs
  - `show` (line 40) - Show various types of objects
  - `diff` (line 17) - Show changes between commits, working tree, etc

- **Branching**
  - `branch` (line 8) - List, create, or delete branches  
  - `checkout` (line 10) - Switch branches or restore files
  - `switch` (line 45) - Switch branches
  - `merge` (line 27) - Join development histories

- **Remote Operations**
  - `fetch` (line 18) - Download objects and refs from another repository
  - `pull` (line 30) - Fetch from and integrate with another repository
  - `push` (line 31) - Update remote refs along with associated objects

- **Advanced Local Operations**
  - `stash` (line 42) - Stash changes in dirty working directory
  - `reset` (line 34) - Reset current HEAD to specified state
  - `rebase` (line 33) - Reapply commits on top of another base

### Ancillary Commands (Lines 49-61)
- `remote` (line 58) - Manage set of tracked repositories

---

## Detailed Level Breakdown

### Level 1: Repository Basics
**Goal:** Understand what Git is and how to start using it

**Commands to Master:**
1. `git init` - Initialize a new repository
2. `git status` - Check current repository state

**Learning Objectives:**
- [ ] Understand what a Git repository is
- [ ] Know how to create a new repository
- [ ] Learn to check the status of files
- [ ] Understand working directory vs staging area concepts

### Level 2: File Management  
**Goal:** Learn to track and commit changes

**Commands to Master:**
1. `git add` - Stage files for commit
2. `git commit` - Save changes to repository
3. `git rm` - Remove files from Git tracking
4. `git mv` - Rename/move files in Git

**Learning Objectives:**
- [ ] Stage individual files and multiple files
- [ ] Write meaningful commit messages
- [ ] Understand the staging area concept
- [ ] Manage file deletions and renames

### Level 3: History & Inspection
**Goal:** Navigate and understand project history

**Commands to Master:**
1. `git log` - View commit history
2. `git show` - Examine specific commits
3. `git diff` - Compare changes
4. `git blame` - See who changed what

**Learning Objectives:**
- [ ] Navigate commit history effectively
- [ ] Compare different versions of files
- [ ] Understand commit information
- [ ] Track changes by author

### Level 4: Branching & Navigation  
**Goal:** Work with multiple versions simultaneously

**Commands to Master:**
1. `git branch` - Create and list branches
2. `git checkout` - Switch between branches
3. `git switch` - Modern way to switch branches

**Learning Objectives:**
- [ ] Understand branching concepts
- [ ] Create feature branches
- [ ] Switch between different branches
- [ ] Manage multiple lines of development

### Level 5: Merging & Advanced Local Operations
**Goal:** Combine work and handle complex scenarios

**Commands to Master:**
1. `git merge` - Combine branches
2. `git stash` - Temporarily save work
3. `git reset` - Undo changes
4. `git rebase` - Reapply commits (advanced)

**Learning Objectives:**
- [ ] Merge branches successfully
- [ ] Handle merge conflicts
- [ ] Temporarily stash work in progress
- [ ] Undo mistakes safely

### Level 6: Remote Operations & Collaboration
**Goal:** Work with remote repositories and collaborate with others

**Commands to Master:**
1. `git clone` - Copy a repository from remote
2. `git remote` - Manage remote repositories
3. `git fetch` - Download changes from remote
4. `git pull` - Fetch and merge remote changes
5. `git push` - Upload local changes to remote

**Learning Objectives:**
- [ ] Clone existing repositories
- [ ] Set up and manage remote connections
- [ ] Understand fetch vs pull differences
- [ ] Push local changes to remote repositories
- [ ] Handle basic collaboration scenarios

---

## Practice Lab Setup

### Lab Environment Requirements
- Terminal/Command line access
- Text editor (nano, vim, or VS Code)
- Git installed and configured
- Practice directory: `/home/jbragdon/git/cae/labs/git-practice/`
- GitHub account (for Level 6 remote operations)

### Next Steps
1. ✅ Created comprehensive learning guide structure
2. 🔲 Build Level 1 exercises and hands-on labs
3. 🔲 Create sample files and scenarios for each level
4. 🔲 Add detailed command reference with examples
5. 🔲 Design remote repository exercises for Level 6
6. 🔲 Test all exercises for clarity and effectiveness

## Notes Section
- This guide progresses from local Git operations to remote collaboration
- Each level builds on previous knowledge systematically
- Hands-on exercises will be included for every level
- Real-world scenarios and best practices included
- Level 6 introduces collaboration and remote workflow concepts

---

*Generated from all-git-commands.txt reference - Lines referenced throughout for accurate command descriptions*