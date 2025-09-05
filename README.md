### 📊 HOP Excel Processor

Desktop application developed in **Python** with a graphical interface (**CustomTkinter**)  
to process Excel files with fish data from the farm in **L'Ametlla de Mar (MRAG)**.

This tool helps normalize and filter **HOP (Harvest Observation Program)** data,  
making observers' work easier by standardizing formats and calculating specific fields.

---

## 🚀 Features

- Select multiple Excel files (`.xlsx`, `.xls`).
- Choose an output folder for processed results.
- Multilingual interface: **ES, EN, FR, PT, HR**.
- Two main processing options:
  1. **Format without estimated weights**
     - Marks doubtful weights (two individuals sharing weight) as `NA`.
     - Adjusts the `Tipus` column according to the MRAG dictionary.
  2. **Calculate `PesIndiv`**
     - Calculates individual fish weight based on length/width ratio and total `Pes M`.
     - Normalizes the `Tipus` column.
     - Does not affect the total HOP weight.
- Displays selected and processed files on screen.
- Automatically generates processed Excel files:
  - `Filtered_HOPxxx.xlsx`
  - `Individual_HOPxxx.xlsx`

---

## 🛠️ Requirements

- **Operating System**: Windows (uses `pythoncom` and `win32com`).
- **Python**: 3.9 or higher
- **Dependencies**:
  ```bash
  pip install pandas numpy pillow customtkinter pywin32

<img width="890" height="674" alt="image" src="https://github.com/user-attachments/assets/cc268d79-063c-480e-b7ba-000d69612e83" />
