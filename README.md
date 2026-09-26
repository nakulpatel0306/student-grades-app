# Student Grades App - Excel VBA Marking System

An Excel VBA tool that connects to a Microsoft Access student database, calculates course statistics, and generates reports in Word.

## At a Glance

- **Stack:** Excel VBA, Microsoft Access, SQL, Word automation
- **Context:** 200-level data automation course, Wilfrid Laurier University
- **State:** Complete
- **Requires:** Windows with Microsoft Excel, Access and Word

## Features

- Connects to any compatible Access database through a file picker
- Lists courses and calculates course averages and standard deviations
- Looks up students by Student ID
- Generates course and student reports as Word documents
- Validates inputs and catches missing or wrong file types

## Project Structure

```
Student Grades App/
├── pate1079_a05.xlsm               # Macro-enabled workbook (the app)
├── Registrar.mdb                   # Sample Access database
├── student-grades-app-overview.pdf # Project write-up
└── macro code files/               # Exported VBA forms and modules
    ├── main selection form.frm
    ├── enrollment form.frm
    ├── generate report form.frm
    └── module one/two/three.bas
```

## Running Locally

1. Clone the repo, or download it as a ZIP:
   ```bash
   git clone https://github.com/nakulpatel0306/student-grades-app.git
   ```
2. Open `Student Grades App/pate1079_a05.xlsm` in Excel and enable macros.
3. Click **Browse**, pick `Registrar.mdb`, then click **Run**. A message confirms the connection.
4. Use the main form to view courses, calculate stats, search students or generate reports.

If Run does nothing, check that a file was selected and that it is an Access database.
