# Quick Reference Guide for Teachers

## 🚀 Most Common Tasks

### 1. Create Master Roster (Start of Semester)
**Script:** `roster_management/create_master_roster.py`

**Steps:**
1. Export rosters from gradebook as CSV (one per period)
2. Open script, update file paths at top
3. Run: `python create_master_roster.py`
4. Get: One Excel file with all students

**Time:** 2 minutes

---

### 2. Grade Writing Assignment (Weekly)
**Script:** `grading/grade_ace_writing.py`

**Steps:**
1. Export Google Form responses as CSV (one per period)
2. Open script, update:
   - Form file paths
   - Prompts dictionary (your assignment text/questions)
3. Run: `python grade_ace_writing.py`
4. Get: Excel file with scores for each student

**Time:** 5 minutes (first time), 2 minutes (after setup)

**Note:** Spot-check borderline scores (3-4 points)

---

### 3. Update Gradebook (After Assignment)
**Script:** `gradebook/update_gradebook_scores.py`

**Steps:**
1. Have Google Form responses CSV
2. Have gradebook roster CSV
3. Update script with file paths and points to add
4. Run script
5. Get: Updated roster with points added

**Time:** 3 minutes

---

## 📱 Emergency Contacts

**If scripts aren't working:**
1. Check file paths are correct
2. Verify CSV files are formatted correctly
3. Look for error messages (red text)
4. Check the script's comments for hints

**Common Errors:**
- "File not found" → Fix file path
- "Column not found" → Check CSV format
- "Name not matched" → Check spelling/format

---

## 💡 Pro Tips

1. **Test with one period first** before running all periods
2. **Keep backups** of roster files before updating
3. **Standardize file naming** for easier script setup
4. **Review outputs** - automation speeds things up but always check results
5. **Share improvements** - if you modify a script, share it!

---

## 🔧 Installation (One Time Only)

```bash
# Install Python (if needed)
# Download from python.org

# Install required packages
pip install pandas openpyxl

# You're ready to go!
```

---

## 📞 Need Help?

- Check README.md for detailed documentation
- Look at script comments (lines starting with #)
- Review GITHUB_SETUP.md for repository help
- Ask a tech-savvy colleague for first-time setup

---

## ⏱️ Time Saved Per Week

Typical teacher using these scripts:
- Manual roster organization: 30 min → 2 min ✅ **28 min saved**
- Grading 100 writing responses: 3 hours → 15 min ✅ **2.75 hours saved**
- Updating gradebook: 20 min → 3 min ✅ **17 min saved**

**Total: ~3.5 hours saved per week!** 🎉

---

## 🎯 Script Decision Tree

```
Need to organize students?
└─ YES → create_master_roster.py

Need to grade writing?
└─ YES → Is it ACE rubric?
   ├─ YES → grade_ace_writing.py
   └─ NO → (add your custom script here)

Need to track completion?
└─ YES → match_form_submissions.py

Need to update grades?
└─ YES → update_gradebook_scores.py
```

---

**Last Updated:** Week of January 27, 2026
**Version:** 1.0
**Questions?** Check the full README.md
