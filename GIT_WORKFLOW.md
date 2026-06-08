# Git Workflow Guide for Educational Automation Library

## Initial Setup (One Time Only)

### Step 1: Clone Your GitHub Repository to WSL

```bash
# Navigate to your projects directory
cd ~/projects

# Clone your repository (replace YOUR_USERNAME and REPO_NAME)
git clone https://github.com/YOUR_USERNAME/edu-automation-scripts.git

# Navigate into the repository
cd edu-automation-scripts
```

**Alternative if using SSH:**
```bash
git clone git@github.com:YOUR_USERNAME/edu-automation-scripts.git
```

### Step 2: Set Up Git Identity (If First Time Using Git)

```bash
# Set your name (will appear in commits)
git config --global user.name "Your Name"

# Set your email (should match GitHub account)
git config --global user.email "your.email@example.com"

# Optional: Set default editor
git config --global core.editor "nano"
```

### Step 3: Install Python Dependencies

```bash
# Using system packages (easiest)
sudo apt install python3-pandas python3-openpyxl

# OR using virtual environment (best practice)
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

---

## Daily Workflow: Making Changes and Pushing to GitHub

### Scenario 1: You Modified a Script Locally

```bash
# Navigate to your repository
cd ~/projects/edu-automation-scripts

# Check what files changed
git status

# See exactly what changed
git diff grading/grade_ace_writing_interactive.py

# Add the changed file to staging
git add grading/grade_ace_writing_interactive.py

# OR add all changed files
git add .

# Commit with a descriptive message
git commit -m "Fix: Handle compound names in grading script"

# Push to GitHub
git push
```

### Scenario 2: You Created a New Script

```bash
# Create your new script
nano grading/new_script.py

# Add it to git
git add grading/new_script.py

# Commit
git commit -m "Add: New script for processing attendance data"

# Push to GitHub
git push
```

### Scenario 3: You Want to Download Latest Scripts from Another Computer

```bash
# Navigate to repository
cd ~/projects/edu-automation-scripts

# Pull latest changes from GitHub
git pull

# Now you have the latest version!
```

---

## Typical Teacher Workflow

### Monday Morning: Update Scripts for New Week

```bash
cd ~/projects/edu-automation-scripts

# Pull any changes from GitHub (if working on multiple computers)
git pull

# Make your changes
nano grading/grade_ace_writing_interactive.py

# Test your changes
python3 grading/grade_ace_writing_interactive.py

# If it works, commit and push
git add .
git commit -m "Update: Week 2 assignment prompts"
git push
```

### Sharing Improvements with Other Teachers

```bash
# After making improvements
git add .
git commit -m "Add: Better error handling for missing files"
git push

# Share the GitHub URL with colleagues
# They can now pull your improvements!
```

---

## Useful Git Commands Reference

### Checking Status

```bash
# See what files changed
git status

# See detailed changes
git diff

# See commit history
git log

# See recent commits in one line each
git log --oneline
```

### Undoing Mistakes

```bash
# Undo changes to a file (before committing)
git checkout -- filename.py

# Undo last commit (keep changes in files)
git reset --soft HEAD~1

# Undo last commit (discard all changes)
git reset --hard HEAD~1

# See what a file looked like in a previous commit
git show HEAD~1:grading/script.py
```

### Working with Branches (Advanced)

```bash
# Create a new branch to test changes
git checkout -b experimental-feature

# Make changes, commit them
git add .
git commit -m "Experimental: Try new grading algorithm"

# If it works, merge back to main
git checkout main
git merge experimental-feature

# If it doesn't work, just switch back to main
git checkout main
```

---

## Best Practices for Teachers

### ✅ DO:

1. **Commit often with clear messages**
   ```bash
   git commit -m "Fix: Handle students with no middle name"
   ```

2. **Pull before you push** (if working on multiple computers)
   ```bash
   git pull
   # make changes
   git push
   ```

3. **Test scripts before committing**
   - Run the script with sample data
   - Make sure it works
   - Then commit and push

4. **Use descriptive commit messages**
   - Good: "Add: Interactive prompts to grading script"
   - Bad: "Update"

5. **Keep sensitive data out of git**
   - Never commit real student rosters
   - Never commit CSV files with student data
   - The .gitignore file protects you from this

### ❌ DON'T:

1. **Don't commit without testing**
   - Broken code affects everyone who pulls

2. **Don't commit large files**
   - Keep actual student data files separate
   - Only commit the scripts themselves

3. **Don't force push (git push -f)**
   - This can overwrite others' work
   - Only use if you really know what you're doing

---

## Quick Command Cheat Sheet

```bash
# The "I just want to push my changes" workflow
git add .
git commit -m "Your message here"
git push

# The "I want the latest version" workflow
git pull

# The "What did I change?" workflow
git status
git diff

# The "I messed up, start over" workflow
git reset --hard HEAD
git pull

# The "I want to save work but not commit yet" workflow
git stash              # Save work temporarily
git pull               # Get latest
git stash pop          # Restore your work
```

---

## Troubleshooting Common Issues

### "Permission denied (publickey)"

**Option 1: Use HTTPS instead**
```bash
# Change remote URL to HTTPS
git remote set-url origin https://github.com/YOUR_USERNAME/REPO_NAME.git
```

**Option 2: Set up SSH keys** (see GITHUB_SETUP.md)

### "Your branch is behind 'origin/main'"

```bash
# Pull the latest changes
git pull

# If there are conflicts, git will tell you which files
# Edit those files to resolve conflicts, then:
git add .
git commit -m "Merge remote changes"
```

### "Changes not staged for commit"

```bash
# You edited files but didn't add them
git add .
git commit -m "Your message"
```

### Accidentally Committed Student Data

```bash
# Remove from repository but keep local file
git rm --cached sensitive_file.csv
git commit -m "Remove sensitive file from repo"
git push

# Add to .gitignore to prevent future accidents
echo "sensitive_file.csv" >> .gitignore
git add .gitignore
git commit -m "Update gitignore"
git push
```

---

## Directory Structure in Your Git Repo

```
~/projects/edu-automation-scripts/    (Your local repository)
├── README.md                          (Main documentation)
├── GITHUB_SETUP.md                    (Setup instructions)
├── QUICK_REFERENCE.md                 (Quick guide)
├── requirements.txt                   (Python dependencies)
├── .gitignore                         (Files to never commit)
│
├── roster_management/
│   ├── create_master_roster.py
│   └── create_master_roster_interactive.py
│
├── grading/
│   ├── grade_ace_writing.py
│   ├── grade_ace_writing_interactive.py
│   ├── match_form_submissions.py
│   └── match_form_submissions_interactive.py
│
└── examples/
    └── sample_data/
```

**Keep your actual work files separate:**
```
~/teaching/                            (NOT in git)
├── rosters/
│   ├── Period_1_Roster.csv           (Real student data)
│   └── Period_3_Roster.csv
├── assignments/
│   └── Week_1_Responses.csv
└── output/
    └── Graded_Assignments.xlsx
```

---

## Getting Help

- **Git documentation**: https://git-scm.com/doc
- **GitHub guides**: https://guides.github.com
- **Quick help**: `git help <command>`
  - Example: `git help commit`

---

**Remember**: Git seems complicated at first, but 90% of the time you'll only use:
```bash
git pull       # Get latest
git add .      # Stage changes
git commit -m "Message"    # Commit
git push       # Upload to GitHub
```

Everything else is just for special situations!
