## Excel VBA: Quick Reference Guide

This reference outlines the core syntax and methodologies required for interacting with Excel tables, User Forms, and external Access databases using VBA. Understanding these approaches ensures efficient data handling, clean variable management, and stable database connections.

---

### 1. Core Concepts

**Subs vs. Functions**
Every script operates within a routine. Categorize routines based on their intended output.

- **Subroutine (Sub):** Performs an action or sequence of operations. Does not return a value.
- **Function:** Performs a calculation or evaluates data. Returns a specific value to the caller.

| Routine Type | Purpose            | Example Use Case                                                |
| :----------- | :----------------- | :-------------------------------------------------------------- |
| **Sub**      | Execute a process. | Formatting a table, clearing a form, executing a SQL query.     |
| **Function** | Return a value.    | Calculating a custom metric, returning a formatted date string. |

**Variable Scopes**
Variables must be declared at the correct level to ensure data is available when needed and cleared when it is not.

- **Procedure-Level (Sub/Function):** Declared using `Dim` inside a Sub or Function. Exists only while the routine is running.
- **Module-Level:** Declared using `Private` or `Dim` at the very top of a Module. Available to all Subs and Functions within that specific Module.
- **Global-Level:** Declared using `Public` at the top of a Module. Available universally across the entire VBA project.

---

### 2. Control Flow (Loops)

Loops dictate how data structures are traversed. Choose the loop type based on the object being evaluated.

- **For Next:** Iterate a specific, predefined number of times.
    ```vba
    Dim i As Integer
    For i = 1 To 10
        ' Do something 10 times
    Next i
    ```
- **For Each:** Iterate through every item within a collection (ideal for Table rows or columns).
    ```vba
    Dim tblRow As ListRow
    For Each tblRow In ActiveSheet.ListObjects("MyTable").ListRows
        ' Do something to each row
    Next tblRow
    ```
- **Do While:** Continue looping as long as a specific condition remains true.
    ```vba
    Dim count As Integer
    count = 1
    Do While count <= 5
        ' Do something until count exceeds 5
        count = count + 1
    Loop
    ```

---

### 3. Excel Tables (ListObjects) CRUD Operations

Working directly with the `ListObject` API provides a structured way to manipulate data without relying on hardcoded cell ranges.

**Setup Reference:**
To interact with a table, map it to a variable first.
`vba
    Dim tbl As ListObject
    Set tbl = ActiveSheet.ListObjects("DataTable")
    `

**Read (Extracting Data):**
To pull a value from a specific row and column into a variable.
`vba
    Dim targetValue As String
    ' Read the value from the first row, "Status" column
    targetValue = tbl.ListRows(1).Range(1, tbl.ListColumns("Status").Index).Value
    `

**Update (Modifying Data):**
To change a value within a specific table record.
`vba
    ' Change the value in the first row, "Status" column
    tbl.ListRows(1).Range(1, tbl.ListColumns("Status").Index).Value = "Complete"
    `

**Create (Adding a Row):**
To append a new blank row to the bottom of the table and populate it.
```vba
Dim newRow As ListRow
Set newRow = tbl.ListRows.Add

    ' Populate the new row
    newRow.Range(1, tbl.ListColumns("ID").Index).Value = 101
    newRow.Range(1, tbl.ListColumns("Status").Index).Value = "Pending"
    ```

**Delete (Removing a Row):**
