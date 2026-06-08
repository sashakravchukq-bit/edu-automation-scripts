# Educational Automation Scripts for Teachers

A collection of Python scripts designed to automate common educational tasks, saving teachers time on administrative work so they can focus on teaching.

## 🎯 Quick Search by Task

**Need to...?**
- Combine multiple class rosters → [`create_master_roster.py`](#create-master-class-roster)
- Grade writing assignments with ACE rubric → [`grade_ace_writing.py`](#grade-ace-rubric-writing)
- Match Google Form submissions to gradebook → [`match_form_submissions.py`](#match-google-form-submissions)
- Update gradebook scores from forms → [`update_gradebook_scores.py`](#update-gradebook-scores)
- Create completion reports → See [`match_form_submissions.py`](#match-google-form-submissions)

## 📚 Scripts by Category

### Roster Management
- [Create Master Class Roster](#create-master-class-roster) - Combine multiple period rosters into one master list

### Grading & Assessment
- [Grade ACE Rubric Writing](#grade-ace-rubric-writing) - Automatically grade short answer responses using ACE strategy
- [Match Google Form Submissions](#match-google-form-submissions) - Create completion reports from Google Forms

### Gradebook Automation
- [Update Gradebook Scores](#update-gradebook-scores) - Match form submissions and update assignment scores

---

## 📖 Script Documentation

### Create Master Class Roster

**File**: `roster_management/create_master_roster.py`

**Purpose**: Combines individual class period rosters into a single master roster with all students.

**Use Case**: Beginning of semester when you need a complete list of all students across all your class periods.

**Inputs Required**:
- CSV files for each class period (exported from your gradebook system)
- Files should contain: Student Name, Grade, Student ID, and any assignment columns

**Output**:
- Excel file with columns: Student Name, Grade, Student ID, Period, Course

**Features**:
- Handles compound last names (e.g., "Garcia Rodriguez")
- Automatically extracts period number from file headers
- Sorts by period, then alphabetically by name
- Color-coded headers for easy reading

**Keywords**: roster, class list, master list, student management, organization, periods

---

### Grade ACE Rubric Writing

**File**: `grading/grade_ace_writing.py`

**Purpose**: Automatically grades short answer writing responses using the ACE (Answer, Cite, Explain) rubric strategy.

**Use Case**: Weekly writing assignments (like "Write it Wednesday"), TELPAS requirements, any ACE-based writing assessment.

**Inputs Required**:
- Google Form CSV exports with student responses (one per class period)
- Master roster Excel file (to include non-submitters)
- Prompts file with the text passages and questions for each period

**Output**:
- Excel file with individual scores for Answer (0-2), Cite (0-2), Explain (0-2), Total (0-6)
- Includes students who didn't submit (scored as 0)
- Color-coded: Green (5-6), Yellow (3-4), Red (0-2)
- Response preview for each student

**Grading Criteria**:
- **Answer (0-2)**: Checks if response addresses all parts of the question
- **Cite (0-2)**: Looks for quotation marks and citation language ("author states", "according to", etc.)
- **Explain (0-2)**: Identifies explanation words (because, this shows, therefore) and logical structure

**Features**:
- Matches student emails to names from master roster
- Customizable rubric criteria per period/topic
- Includes class and overall statistics
- Preview of each response for manual review

**Keywords**: grading, writing, ACE rubric, TELPAS, assessment, automatic grading, short answer

---

### Match Google Form Submissions

**File**: `grading/match_form_submissions.py`

**Purpose**: Creates a completion report showing which students submitted a Google Form assignment.

**Use Case**: Quickly see who turned in an assignment, create a completion checklist, identify missing work.

**Inputs Required**:
- Google Form CSV exports (one per class period)
- Master roster Excel file with all students

**Output**:
- Excel file with completion status (Yes/No) for each student
- Color-coded: Green (completed), Red (not completed)
- Includes all students from roster (not just those who submitted)

**Features**:
- Fuzzy name matching (handles variations in name formatting)
- Handles compound names and special characters
- Shows completion percentage by period and overall

**Keywords**: completion, Google Forms, assignment tracking, missing work, attendance

---

### Update Gradebook Scores

**File**: `gradebook/update_gradebook_scores.py`

**Purpose**: Matches Google Form submissions to your gradebook roster and adds points to specified assignment columns.

**Use Case**: Adding assignment credit to your gradebook after collecting responses via Google Forms.

**Inputs Required**:
- Google Form CSV export with student responses
- Gradebook roster CSV (one per class period)
- Assignment name and point value

**Output**:
- Updated roster CSV files with points added to assignment columns
- Report showing which students were matched and credited

**Features**:
- Handles different column positions across periods
- Matches compound last names correctly
- Preserves all existing gradebook data
- Reports unmatched students for manual review

**Keywords**: gradebook, points, scores, update, Google Forms, assignment credit

---

## 🚀 Getting Started

### Prerequisites
- Python 3.7 or higher
- Required packages: `pandas`, `openpyxl`, `csv`

Install required packages:
```bash
pip install pandas openpyxl
```

### Basic Usage Pattern

1. **Organize your input files** in a folder
2. **Choose the appropriate script** from the categories above
3. **Update the file paths** in the script to match your files
4. **Run the script**: `python script_name.py`
5. **Check the output** folder for your results

### Example Workflow

**Scenario**: Grade weekly writing assignment

```bash
# 1. Collect your files
#    - Period 1, 3, 4, 5 Google Form CSV exports
#    - Master roster Excel file
#    - Prompts text file

# 2. Run the grading script
python grading/grade_ace_writing.py

# 3. Review the output
#    - Write_it_Wednesday_Week_1_Grades.xlsx
#    - Shows all students with scores or "NO SUBMISSION"
```

---

## 📂 File Naming Conventions

For best results, name your input files consistently:

**Roster files**: `COURSE_Spring_2026_Period_#_Roster.csv`
- Example: `APCSP_Spring_2026_Period_1_Roster.csv`

**Google Form exports**: `Week_#_Period_#_Assignment_Name.csv`
- Example: `Week_1_Period_1_Writing_Project_Volcanoes.csv`

---

## 🤝 Contributing

These scripts were developed through real classroom use. If you:
- Find bugs or issues
- Have suggestions for improvements
- Want to add new scripts
- Need help adapting scripts to your workflow

Please feel free to open an issue or submit a pull request!

---

## 📋 Script Customization Tips

Most scripts can be easily customized by changing variables at the top:

```python
# Example: Change these paths to match your files
input_folder = '/path/to/your/files'
output_folder = '/path/to/save/results'

# Example: Adjust grading criteria
PASSING_SCORE = 4  # Out of 6 for ACE rubric
```

---

## 🏫 Real Classroom Context

These scripts were developed for:
- AP Computer Science Principles
- Advanced Computer Science I
- District requirements (TELPAS writing)
- Classes of 20-28 students per period
- Google Forms for assignment submission
- CSV exports from school gradebook systems

They can be adapted for other subjects and workflows!

---

## 📄 License

These scripts are provided as-is for educational use. Feel free to modify and share with fellow educators.

---

## 💡 Tips for Success

1. **Always test with sample data first** before using on real grades
2. **Keep backups** of original roster files before running update scripts
3. **Review output files** - automated grading is a time-saver but should be spot-checked
4. **Consistent naming** helps scripts find and match files correctly
5. **Read the comments** in each script for specific customization options

---

## 🔍 Search Keywords Reference

**By Task Type**:
- Organization: roster, master list, student management
- Grading: ACE rubric, writing, assessment, scores
- Tracking: completion, missing work, attendance
- Updates: gradebook, points, assignment credit

**By Input Type**:
- Google Forms, CSV, Excel, roster files

**By Output Type**:
- Reports, updated rosters, grade sheets, completion lists

---

**Created by teachers, for teachers. Happy automating! 🎓**
