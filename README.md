# Terbilang (Number-to-Words) for Excel & Word 🚀

This repository provides a VBA-based function to automatically convert numbers into written words (Terbilang). Perfect for finance administration, generating invoices, or any document that requires formal currency/number naming.

---

## 📑 Table of Contents
1. [How to Install & Use (Excel Beginners)](#how-to-install--use-for-beginners)
2. [Level Up: Use in ALL Excel Files (Add-In)](#level-up-use-the-function-in-all-your-excel-files-excel-add-in)
3. [How to Use in Microsoft Word](#how-to-use-in-microsoft-word)
4. [Source Code](#source-code)

---

## 📄 Source Code
The source code is located in terbilang-vba.txt. Simply copy the text from that file and follow the steps below.

---

## 🟢 How to Install & Use (For Beginners)
If you are new to VBA, don't worry! Follow these simple steps to get the function working:

### Step 1: Enable the Developer Tab
1. Right-click anywhere on the Ribbon (the top menu bars).
2. Select Customize the Ribbon.
3. In the right-hand list under Main Tabs, check the box for Developer.
4. Click OK.

### Step 2: Add the Script to VBA
1. Press Alt + F11 to open the VBA Editor.
   *(Note: If the shortcut doesn't work, go to the Developer tab and click the Visual Basic icon).*
2. In the menu bar, click Insert > Module.
3. A blank white window will appear. Copy and paste the code from terbilang-vba.txt into that window.
4. Close the VBA Editor window.

### Step 3: Using the Function
You can now use it like a regular Excel formula:
- Type: =Terbilang(A1) (assuming the number is in cell A1).
- Press Enter, and the number will convert automatically!

> [!IMPORTANT]
> Important Note for Excel: When saving your file, make sure to save it as an Excel Macro-Enabled Workbook (.xlsm). If you save it as a standard .xlsx file, the script will be removed.

---

## ⚡ Level Up: Use in ALL Your Excel Files (Excel Add-In)
Make the function available permanently in every Excel file you open:

### Step 1: Save as an Add-In
1. Open a new, blank Excel file.
2. Follow the steps above to paste the code into a Module.
3. Click File > Save As.
4. Change the file type to Excel Add-in (.xlam).
5. Name it Terbilang_AddIn and save it (Excel will suggest the default AddIns folder).

### Step 2: Activate the Add-In
1. Go to the Developer tab.
2. Click on Excel Add-ins.
3. Find Terbilang_AddIn, check the box, and click OK.

> [!TIP]
> Quick Note: If you share an .xlsm file with others, the recipient must click "Enable Content" at the top for the function to work on their computer.

---

## 🔵 How to Use in Microsoft Word
In Word, the macro replaces a selected number with words.

### Step 1: The Setup
1. Press Alt + F11 > Insert > Module.
2. Paste the code and close the editor.

### Step 2: Running the Macro
1. Highlight/Select the number you want to convert (e.g., 1.500.000).
2. Press Alt + F8 to open the Macro dialog.
3. Select the function (e.g., TerbilangWord) and click Run.

> [!TIP]
> Pro-Tip: To use this in *every* Word document, paste the code under the "Normal" project in the VBA Editor sidebar instead of "Document1".

---
Built with ☕ and a passion for efficiency.
