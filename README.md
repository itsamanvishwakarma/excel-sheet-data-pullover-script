# Excel Data Update Script

This script helps you automatically update the "Physical Count" data in your Secondary Excel file using information from a Primary Excel file. Here's a simple guide to how it works:

---

## 🎯 **Purpose**

- Copies "Physical Count" numbers from the **Primary Excel File** to the **Secondary Excel File** when matching items are found.
- If no match is found, it writes "NA" in the Secondary file to indicate missing data.

---

## ⚙️ **Prerequisites**

1. **Install Python**: If not already installed, download from [python.org](https://www.python.org).
2. **Install Required Tools** (Run these commands in Command Prompt/Terminal):
   ```
   pip install pandas openpyxl
   ```

---

## 🔧 **Configuration**

Update these settings in the script **BEFORE running**:

1. **Primary File Path**: Location of your source Excel file (e.g., `C:\Users\...\primary.xlsx`).
2. **Primary Sheet Name**: Name of the sheet in the Primary file (default: `Sheet1`).
3. **Secondary File Path**: Location of your target Excel file (e.g., `C:\Users\...\secondary.xlsx`).
4. **Secondary Sheet Name**: Name of the sheet in the Secondary file (default: `Sheet1`).

---

## 🔄 **How It Works**

1. **Load Data**:
   - Reads the Primary and Secondary Excel files.
   - Focuses on columns **Item Name (C)**, **Color (D)**, **Opening Qty (E)**, and **Physical Count (Q)**.
2. **Match Rows**:
   - Looks for rows with the same **Item Name + Color + Opening Qty** in both files.
3. **Update or Flag**:
   - If a match is found: Copies the "Physical Count" from Primary to Secondary.
   - If no match is found: Writes "NA" in the Secondary file.

---

## 🖥️ **Usage**

1. Place your Primary and Secondary Excel files in the correct folders.
2. Update the configuration settings in the script (file paths and sheet names).
3. Run the script:
   ```
   python script_name.py
   ```
4. Check the Secondary Excel file for updates.

---

## 📊 **Output Explanation**

After running the script, you’ll see:

- A list of updated rows with details.
- A summary like:
  ```
  Total updates made: 25
  Total rows with no matches (set to NA): 5
  ```

---

## Thing to to after running the script

- Go to the place where data is updated in the secondary file and check if the data is updated correctly.
- This script is paste the data one column to the left from where it supposed to paste (`Pease consider that in mind`).
- And also this Script will paste the data one Cell above from where it supposed to paste (`Pease consider that in mind`).
- Simple add another cell to the top.

## 📝 **Notes**

- **Backup Your Files**: Always make a copy of your Excel files before running the script.
- **Close Excel Files**: Ensure both Primary and Secondary files are closed during the script run.
- **Formatting**: The script preserves Excel formatting (colors, formulas, etc.).

---

## 🚨 **Troubleshooting**

- **"File Not Found" Error**: Double-check the file paths in the configuration.
- **Data Not Updating**: Ensure the Item Name, Color, and Opening Qty columns match exactly in both files.
- **"NA" in All Rows**: This means no matches were found. Verify the data in both files.

---
