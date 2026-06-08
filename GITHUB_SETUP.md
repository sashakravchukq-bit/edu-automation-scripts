# GitHub Repository Setup Guide

## Step-by-Step Instructions

### 1. Create GitHub Account (if needed)
- Go to https://github.com
- Click "Sign up"
- Choose a username (suggestion: use your teacher name or school)
- Complete registration

### 2. Create New Repository

**Option A: Through GitHub Website (Easiest)**
1. Log into GitHub
2. Click the "+" icon in top right → "New repository"
3. Repository name: `edu-automation-scripts` (or your preferred name)
4. Description: "Python scripts to automate educational tasks"
5. Choose "Public" (so other teachers can find it) or "Private" (only you can see)
6. Check "Add a README file"
7. Add .gitignore: Choose "Python"
8. Choose license: "MIT License" (allows others to use/modify)
9. Click "Create repository"

**Option B: Through Command Line (If comfortable with Git)**
```bash
# Navigate to your folder
cd /path/to/edu-automation-library

# Initialize git
git init

# Add files
git add .

# Commit
git commit -m "Initial commit: Educational automation scripts"

# Connect to GitHub (replace USERNAME and REPO_NAME)
git remote add origin https://github.com/USERNAME/REPO_NAME.git

# Push to GitHub
git branch -M main
git push -u origin main
```

### 3. Upload Files to GitHub

**Option A: Web Upload (Easiest)**
1. In your repository, click "Add file" → "Upload files"
2. Drag and drop all files from the `edu-automation-library` folder
   - README.md
   - roster_management/ folder with scripts
   - grading/ folder with scripts
   - gradebook/ folder with scripts
   - examples/ folder
3. Add commit message: "Add initial automation scripts"
4. Click "Commit changes"

**Option B: GitHub Desktop (User-Friendly)**
1. Download GitHub Desktop from https://desktop.github.com
2. Clone your repository
3. Copy all files from `edu-automation-library` to the cloned folder
4. In GitHub Desktop, you'll see all changes
5. Add commit message and push

**Option C: Command Line**
```bash
git add .
git commit -m "Add automation scripts with documentation"
git push
```

### 4. Organize Repository Structure

Your repository should look like this:
```
edu-automation-scripts/
├── README.md                          (main documentation)
├── .gitignore                         (Python gitignore)
├── LICENSE                            (MIT or your choice)
│
├── roster_management/
│   ├── create_master_roster.py       (combine period rosters)
│   └── README.md                      (specific instructions)
│
├── grading/
│   ├── grade_ace_writing.py          (ACE rubric grader)
│   ├── match_form_submissions.py     (completion reports)
│   └── README.md
│
├── gradebook/
│   ├── update_gradebook_scores.py    (add points from forms)
│   └── README.md
│
└── examples/
    ├── sample_data/                   (example CSV files)
    └── README.md                      (how to use examples)
```

### 5. Make Repository Searchable

**Add Topics/Tags:**
1. On your repository page, click the gear icon next to "About"
2. Add topics (these make it searchable):
   - education
   - teacher-tools
   - automation
   - python
   - grading
   - classroom-management
   - google-forms
   - ace-rubric

**Edit Description:**
- Add a clear description like: "Python automation scripts for teachers: roster management, ACE rubric grading, Google Forms integration, and gradebook automation"

### 6. Share with Colleagues

**Share the Repository Link:**
- Your repo URL will be: `https://github.com/YOUR_USERNAME/edu-automation-scripts`
- Share this with fellow teachers
- They can:
  - Browse scripts online
  - Download individual files
  - Clone entire repository
  - Search by keywords

**Create a Quick Start Guide:**
- Pin the README.md to the repository
- Add a "Quick Start" section for non-technical teachers
- Include screenshots if helpful

### 7. Maintaining the Repository

**Adding New Scripts:**
1. Create the script with thorough comments
2. Test it with sample data
3. Add documentation to README.md
4. Upload to GitHub (repeat Step 3)

**Updating Existing Scripts:**
1. Make changes to the script
2. Test thoroughly
3. Update documentation if needed
4. Upload with descriptive commit message like "Fix: Handle compound names in grading script"

**Accepting Contributions:**
- Other teachers can suggest improvements via "Issues"
- Or submit "Pull Requests" with their modifications
- You review and accept/reject changes

### 8. Making It Teacher-Friendly

**For Non-Programmers:**
Create a simplified README with:
1. "I want to..." statements linking to scripts
2. Video tutorials (optional, record yourself using a script)
3. Troubleshooting section for common errors
4. Contact info for questions

**Example "I want to..." Section:**
```markdown
## I want to...

- **Combine my class rosters** → [create_master_roster.py](roster_management/create_master_roster.py)
  - Difficulty: Beginner
  - Time: 2 minutes
  - You'll need: Your roster CSV files

- **Grade writing assignments** → [grade_ace_writing.py](grading/grade_ace_writing.py)
  - Difficulty: Intermediate
  - Time: 5 minutes
  - You'll need: Google Form exports, master roster
```

### 9. Advanced: GitHub Pages (Optional)

Create a website for your scripts:
1. In repository Settings → Pages
2. Source: main branch
3. GitHub will create a website at: `https://USERNAME.github.io/REPO_NAME`
4. Now you have a searchable website instead of just files!

### 10. Backup and Safety

**Important:**
- Don't upload files with real student data
- Use example/dummy data for demonstrations
- Add `.gitignore` to exclude sensitive files:
```
# .gitignore
*.csv
*.xlsx
*_Roster.csv
student_data/
```

## Quick Command Reference

```bash
# Check status
git status

# Add specific file
git add filename.py

# Add all changes
git add .

# Commit with message
git commit -m "Your message here"

# Push to GitHub
git push

# Pull latest changes
git pull

# See history
git log
```

## Getting Help

**If you get stuck:**
1. GitHub has extensive documentation: https://docs.github.com
2. YouTube has many "GitHub for beginners" tutorials
3. Ask a tech-savvy colleague to help with first-time setup
4. Once set up, uploading new files is very simple!

## Benefits for Your Teaching Community

✅ **Searchable** - Teachers can find scripts by keyword
✅ **Shareable** - One link gives access to everything
✅ **Updatable** - Improvements benefit everyone
✅ **Collaborative** - Others can contribute improvements
✅ **Documented** - All scripts have clear instructions
✅ **Free** - GitHub is free for public repositories
✅ **Professional** - Shows initiative and technical skills

---

**Remember:** The hardest part is the initial setup. After that, adding scripts is as simple as uploading a file!
