# Exam Access Incident Log Generator

This app makes Word incident logs from the exam access arrangements spreadsheet.

## How to use

1. Open the app.
2. Choose the access arrangements spreadsheet.
3. For **Exam 1**, choose the year and subject.
4. If there are more exams in the same hall, use **Exam 2** and **Exam 3**.
5. Need more? Click **+ Additional exam**. You can add up to 8 exams in total.
6. Leave unused extra rows as `No additional subject`.
7. Click **Generate room logs**.

The app saves the Word file here by default:

```text
Downloads/Generated Incident Logs
```

Each file name includes the date, for example:

```text
Incident Log - Dance_studio - 28-10-26.docx
```

Example:

- Exam 1: `Year 11` and `Biology`
- Exam 2: `Year 12` and `Physics`
- Exam 3: `No additional subject`

The app creates a separate Word document for each room in the `Room` column.
Each subject is shown as a separator row in the candidate table.
Students in `Alternative room` are left out because their logs are done manually.
Capital letters do not matter: `ALTERNATIVE ROOM` is also left out.
If the Room column is missing or blank, those students go in a `Room not specified` file.
Keep the column heading `Room` when using rooms.

The app does not include any student spreadsheet inside it.
The user must choose the spreadsheet each time.

## Updates

Click **Check for updates** at the bottom of the app.
If a newer version is available, choose to download and restart.
Generate any needed logs first; you will need to choose your spreadsheet again.
Update checks do not upload your spreadsheet or student information.

Get the first version with this button from the
[Windows download page](https://github.com/chaoticsponge/studentsort/releases/latest).
Teachers do not need Python or a GitHub account.

## Do not change these spreadsheet names

The app needs these sheet names to stay the same:

- `Year 10`
- `Year 11`
- `Year 12`
- `Year 13`

The app needs these column headings to stay the same:

- `First name`
- `Surname`
- `Access Arrangement`
- `Access Arrangements`
- `Subject(s) Applicable`

Use either `Access Arrangement` or `Access Arrangements`. Both work.

## For Year 10 and Year 11

Do not change the subject grid layout.

Keep the subject names above the student rows.
Keep using `Y` and `N` under each subject.

The app only includes students who have an access arrangement written in the
`Access Arrangements` column.

## Build the Windows EXE

On Windows, double-click:

```text
build_windows.bat
```

The finished app will be called:

```text
ExamAccessLogGenerator.exe
```
