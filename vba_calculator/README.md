# 🧮 VBA Calculator in Microsoft Excel

A desktop calculator application developed using **Microsoft Excel VBA (Visual Basic for Applications)** and Excel UserForms. The project demonstrates how VBA can be used to build interactive desktop tools within Excel.

## 📌 Project Overview

This project implements a basic calculator with an interactive UserForm interface. Users can perform common arithmetic operations through buttons and receive validation messages for invalid calculations.

## ✨ Features

- Interactive calculator interface using Excel UserForms
- Number and arithmetic-operation buttons
- Addition, subtraction, multiplication, and division
- Clear and backspace controls, as provided by the workbook's form
- Division-by-zero validation
- Input validation for incorrect values

## 🛠️ Technologies Used

| Technology | Purpose |
|---|---|
| Microsoft Excel | Desktop application |
| VBA | Application logic |
| Excel UserForms | Calculator interface |
| CommandButton Controls | User interactions |

## 📂 Project Structure

```text
VBA-Calculator-Excel/
│
├── calculator.xlsm
├── README.md
│
├── source_code/
│   ├── Module1.bas
│   └── CalculatorFormCode.txt
│
└── screenshots/
    └── calculator-ui-preview.png
```

## 🚀 How to Run the Project

1. Download or clone this repository.
2. Open `calculator.xlsm` using Microsoft Excel desktop on Windows.
3. Enable macros only if you trust the workbook and its source.
4. Press `Alt + F11` to open the VBA Editor.
5. Inspect the embedded UserForm named `Calculator` and standard module named `Module1`.
6. Run the calculator using the launch procedure available in the workbook.

**Important:** Excel for the web does not support running VBA macros. Use the desktop version of Excel.

## 💻 Reference Source Code

The `source_code` folder contains a clean reference implementation for a basic two-input calculator.

To use the reference implementation:

- Create a UserForm named `frmCalculator`.
- Add the required TextBoxes, Label, and CommandButtons.
- Copy the code from `CalculatorFormCode.txt` into the UserForm code window.
- Import or add the `Module1.bas` module to launch the form.

The reference implementation is separate from the original workbook's embedded VBA project and is provided for readability.

## 🧪 Testing Checklist

| Test Case | Expected Result |
|---|---:|
| 20 + 10 | 30 |
| 20 - 10 | 10 |
| 20 × 10 | 200 |
| 20 ÷ 10 | 2 |
| Division by zero | Warning message |
| Invalid input | Validation message |

## 📸 Project Preview

The `screenshots` folder contains an illustrative preview of the calculator interface. It is not an actual screenshot captured from Excel.

## 🎯 Learning Outcomes

Through this project, I explored:

- VBA programming fundamentals
- Creating and managing Excel UserForms
- Handling button-click events
- Performing arithmetic calculations
- Implementing basic input validation
- Working with macro-enabled Excel workbooks

## ⚠️ Project Notes

The original `.xlsm` workbook and its embedded VBA project were inspected, but the macros were not executed in Microsoft Excel in this environment. The workbook should be tested in Excel desktop before being presented as a verified working demonstration.

## 👨‍💻 Author

**Rushikesh**

GitHub: [https://github.com/Rushikesh-siddamshetty]

---

*This project was created as a beginner-level practical demonstration of Excel VBA and UserForm development.*