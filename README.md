# 📝 Excel VBA: Extract Comments Function

### 📊 Preview in Use
**Context:** A custom VBA-based User Defined Function (UDF) designed to bridge a gap in native Excel functionality: the inability to extract cell comments (threaded or classic) directly into a cell using a formula.

<p align="center">
  <img src="https://raw.githubusercontent.com/mauroemartinez/VBA-extract-comments-function/main/Images/Extractor%20de%20Comentarios.PNG" alt="Function Preview" width=85%>
</p>

> **Note:** As seen in the image above, the function is versatile and supports both **threaded conversations** (Office 365) and **classic notes**, ensuring no data is left behind regardless of the Excel version used. What's moer, when there are no comments, it will leave the cell empty, no errors.

## 🚀 The Solution
This script creates a new Excel function: `=ExtraerComentarios()`. 
It works just like any native formula (SUM, VLOOKUP), but its purpose is to pull text data from metadata (comments) into the grid, making it searchable and exportable.

## 🛠️ Installation Guide (Step-by-Step)
You don't need to be a developer to use this. Just follow these steps:

1. **Get the Code:** Open the `.bas` file in this repository and copy the script (from the second line onwards).
2. **Open VBA Editor:** In your Excel file, right-click the Sheet tab name and select **"View Code"** (or press `ALT + F11`).
3. **Insert Module:** Go to the top menu: `Insert` > `Module`.
4. **Paste:** Paste the copied script into the new module window.
5. **Save & Return:** Close the VBA window to return to your Excel sheet.
6. **Apply:** Use the function in any cell: 
   `=ExtraerComentarios(A1)` (where A1 is the cell containing the comments).

## 💡 Key Features
* **Full Support:** Handles both modern Threaded Comments and legacy Notes.
* **UDF Integration:** Behaves like a native Excel function once installed.
* **Bulk Processing:** Extract data from hundreds of cells simultaneously by dragging the formula.

## 📸 Script Architecture
Visual representation of the implementation within the VBA Editor:

<p align="center">
  <img src="https://raw.githubusercontent.com/mauroemartinez/VBA-extract-comments-function/main/Images/Script%20extractor%20de%20comentarios.PNG" alt="VBA Script Preview" width=85%>
</p>

---
**Note:** Remember to save your Excel file as **Excel Macro-Enabled Workbook (.xlsm)** to keep the function working.
