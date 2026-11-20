# Student Grades App - Excel VBA Grade Reports From Access

An Excel VBA app that reads a Microsoft Access student database with SQL, builds grade histograms and a statistics table for a course, and writes the results into a Word report. It can also pull a class list for any course.

## At a Glance

- **Stack:** Excel VBA, ADO, SQL, Microsoft Access, Word automation
- **Context:** CP212 Windows Application Programming, Wilfrid Laurier University (fall 2023)
- **State:** Complete
- **Requires:** Windows with Microsoft Excel, Word and the Access database engine (ACE OLEDB 12.0)

## Features

- **Generate Report.** Pick a course and an assessment (A1-A4, Midterm or Exam) and the app:
  - charts that assessment's marks as a histogram in 5-point bins
  - computes every student's final grade (assignments 5% each, midterm 30%, exam 50%) and charts those too
  - fills a table of average, standard deviation, min and max for each assessment on the `Data` sheet
  - opens Word and writes a report with both histograms and the table
- **Course Enrollment.** Pick a course and get a new worksheet named after it, listing each enrolled student's first name, last name and student ID
- Courses are read straight from the database, so the forms work with any Access file that has the same `courses`, `grades` and `students` tables

## Project Structure

```
Student Grades App/
├── pate1079_a05.xlsm               # Macro-enabled workbook (the app)
├── Registrar.mdb                   # Sample Access database
├── student-grades-app-overview.pdf # Project write-up and usage guide
└── macro code files/               # VBA exported from the workbook, for reading on GitHub
    ├── main selection form.frm     # Choose Generate Report or Course Enrollment
    ├── generate report form.frm    # Course and assessment picker, runs the report
    ├── enrollment form.frm         # Course picker for the class list
    ├── module one.bas              # Opens the main form from the sheet or ribbon button
    ├── module two.bas              # Class list query
    └── module three.bas            # Histograms, statistics, final grades, Word report
```

## Running Locally

1. Clone the repo, or download it as a ZIP:
   ```bash
   git clone https://github.com/nakulpatel0306/student-grades-app.git
   ```
2. Open `Student Grades App/pate1079_a05.xlsm` in Excel and enable macros.
3. Start the app with the button on the **Main Page** sheet, or **Open Task Form** on the **Assignment 5** ribbon tab.
4. Choose **Generate Report** or **Course Enrollment**, click **Continue**, and select `Registrar.mdb` in the file dialog.
5. Pick a course (and an assessment, for reports) and click **Generate Report** or **Continue**.

Reports land on the `Histogram` and `Data` sheets and in a new Word document. Class lists land on a new sheet named after the course.
