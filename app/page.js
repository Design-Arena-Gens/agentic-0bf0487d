'use client';

import { useState } from 'react';

export default function Home() {
  const [activeChapter, setActiveChapter] = useState('intro');

  const chapters = {
    intro: {
      title: 'Introduction to VBA',
      content: `
# What is VBA?

Visual Basic for Applications (VBA) is an event-driven programming language from Microsoft that is primarily used with Microsoft Office applications such as Excel, Word, Access, and Outlook.

## Key Features
- **Automation**: Automate repetitive tasks
- **Integration**: Work seamlessly with Office applications
- **Customization**: Create custom functions and procedures
- **User Forms**: Build interactive dialogs and interfaces

## Getting Started
1. Open Excel (or any Office application)
2. Press \`Alt + F11\` to open the VBA Editor
3. Insert a new module: Insert → Module
4. Start writing your code!

## Your First VBA Program

\`\`\`vba
Sub HelloWorld()
    MsgBox "Hello, World!"
End Sub
\`\`\`
      `
    },
    basics: {
      title: 'VBA Basics',
      content: `
# VBA Basics

## Variables and Data Types

### Declaring Variables
\`\`\`vba
Dim myName As String
Dim myAge As Integer
Dim myHeight As Double
Dim isStudent As Boolean
\`\`\`

### Common Data Types
- **String**: Text data
- **Integer**: Whole numbers (-32,768 to 32,767)
- **Long**: Large whole numbers
- **Double**: Decimal numbers
- **Boolean**: True or False
- **Date**: Date and time values
- **Variant**: Any type of data

## Constants
\`\`\`vba
Const PI As Double = 3.14159
Const COMPANY_NAME As String = "Acme Corp"
\`\`\`

## Operators

### Arithmetic
- \`+\` Addition
- \`-\` Subtraction
- \`*\` Multiplication
- \`/\` Division
- \`Mod\` Modulus
- \`^\` Exponentiation

### Comparison
- \`=\` Equal to
- \`<>\` Not equal to
- \`<\` Less than
- \`>\` Greater than
- \`<=\` Less than or equal
- \`>=\` Greater than or equal

### Logical
- \`And\` Logical AND
- \`Or\` Logical OR
- \`Not\` Logical NOT

## Example
\`\`\`vba
Sub VariablesExample()
    Dim firstName As String
    Dim lastName As String
    Dim age As Integer

    firstName = "John"
    lastName = "Doe"
    age = 30

    MsgBox "Name: " & firstName & " " & lastName & vbNewLine & "Age: " & age
End Sub
\`\`\`
      `
    },
    control: {
      title: 'Control Structures',
      content: `
# Control Structures

## If...Then...Else
\`\`\`vba
Sub CheckScore()
    Dim score As Integer
    score = 85

    If score >= 90 Then
        MsgBox "Grade: A"
    ElseIf score >= 80 Then
        MsgBox "Grade: B"
    ElseIf score >= 70 Then
        MsgBox "Grade: C"
    ElseIf score >= 60 Then
        MsgBox "Grade: D"
    Else
        MsgBox "Grade: F"
    End If
End Sub
\`\`\`

## Select Case
\`\`\`vba
Sub CheckDayOfWeek()
    Dim dayNum As Integer
    dayNum = Weekday(Date)

    Select Case dayNum
        Case 1, 7
            MsgBox "It's the weekend!"
        Case 2 To 6
            MsgBox "It's a weekday."
        Case Else
            MsgBox "Invalid day"
    End Select
End Sub
\`\`\`

## For Loop
\`\`\`vba
Sub ForLoopExample()
    Dim i As Integer

    For i = 1 To 10
        Cells(i, 1).Value = i * 2
    Next i
End Sub
\`\`\`

## For Each Loop
\`\`\`vba
Sub ForEachExample()
    Dim cell As Range

    For Each cell In Range("A1:A10")
        cell.Value = cell.Value * 2
    Next cell
End Sub
\`\`\`

## Do While Loop
\`\`\`vba
Sub DoWhileExample()
    Dim i As Integer
    i = 1

    Do While i <= 10
        Cells(i, 1).Value = i
        i = i + 1
    Loop
End Sub
\`\`\`

## Do Until Loop
\`\`\`vba
Sub DoUntilExample()
    Dim i As Integer
    i = 1

    Do Until i > 10
        Cells(i, 1).Value = i
        i = i + 1
    Loop
End Sub
\`\`\`
      `
    },
    procedures: {
      title: 'Procedures & Functions',
      content: `
# Procedures and Functions

## Sub Procedures
Sub procedures perform actions but don't return values.

\`\`\`vba
Sub GreetUser(userName As String)
    MsgBox "Hello, " & userName & "!"
End Sub

' Calling the procedure
Sub Main()
    GreetUser "Alice"
End Sub
\`\`\`

## Functions
Functions perform calculations and return values.

\`\`\`vba
Function CalculateArea(length As Double, width As Double) As Double
    CalculateArea = length * width
End Function

' Using the function
Sub Main()
    Dim area As Double
    area = CalculateArea(5, 10)
    MsgBox "Area: " & area
End Sub
\`\`\`

## Optional Parameters
\`\`\`vba
Function Greet(userName As String, Optional title As String = "Mr.") As String
    Greet = "Hello, " & title & " " & userName
End Function

Sub Main()
    MsgBox Greet("Smith")           ' Hello, Mr. Smith
    MsgBox Greet("Jones", "Dr.")    ' Hello, Dr. Jones
End Sub
\`\`\`

## ByVal vs ByRef
\`\`\`vba
' ByVal: Passes a copy
Sub ChangeValueByVal(ByVal x As Integer)
    x = x + 10
End Sub

' ByRef: Passes reference (default)
Sub ChangeValueByRef(ByRef x As Integer)
    x = x + 10
End Sub

Sub TestParameters()
    Dim num As Integer
    num = 5

    ChangeValueByVal num    ' num is still 5
    ChangeValueByRef num    ' num is now 15
End Sub
\`\`\`

## Scope
\`\`\`vba
' Public: Available to all modules
Public globalVar As Integer

' Private: Available only in current module
Private moduleVar As String

' Procedure-level
Sub Example()
    Dim localVar As Integer  ' Only in this procedure
End Sub
\`\`\`
      `
    },
    excel: {
      title: 'Working with Excel',
      content: `
# Working with Excel Objects

## Workbook Object
\`\`\`vba
' Open a workbook
Workbooks.Open "C:\\Data\\MyFile.xlsx"

' Create new workbook
Workbooks.Add

' Save workbook
ActiveWorkbook.Save
ActiveWorkbook.SaveAs "C:\\Data\\NewFile.xlsx"

' Close workbook
ActiveWorkbook.Close SaveChanges:=True

' Reference specific workbook
Dim wb As Workbook
Set wb = Workbooks("MyFile.xlsx")
\`\`\`

## Worksheet Object
\`\`\`vba
' Reference worksheets
Worksheets("Sheet1").Select
Worksheets(1).Select

' Add new worksheet
Worksheets.Add Before:=Worksheets(1)

' Delete worksheet
Application.DisplayAlerts = False
Worksheets("Sheet2").Delete
Application.DisplayAlerts = True

' Rename worksheet
Worksheets("Sheet1").Name = "Data"

' Copy worksheet
Worksheets("Data").Copy After:=Worksheets("Data")
\`\`\`

## Range Object
\`\`\`vba
' Reference cells
Range("A1").Value = "Hello"
Cells(1, 1).Value = "Hello"  ' Row 1, Column 1

' Range of cells
Range("A1:C10").Value = 0
Range("A1", "C10").Value = 0
Range(Cells(1, 1), Cells(10, 3)).Value = 0

' Common Range properties
Range("A1").Value = 100
Range("A1").Formula = "=SUM(B1:B10)"
Range("A1").Font.Bold = True
Range("A1").Interior.Color = RGB(255, 255, 0)

' Find last row
Dim lastRow As Long
lastRow = Cells(Rows.Count, 1).End(xlUp).Row

' Find last column
Dim lastCol As Long
lastCol = Cells(1, Columns.Count).End(xlToLeft).Column
\`\`\`

## Working with Data
\`\`\`vba
Sub ProcessData()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long

    Set ws = ThisWorkbook.Worksheets("Data")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' Loop through rows
    For i = 2 To lastRow
        ' Read data
        Dim name As String
        Dim score As Integer
        name = ws.Cells(i, 1).Value
        score = ws.Cells(i, 2).Value

        ' Process and write results
        If score >= 70 Then
            ws.Cells(i, 3).Value = "Pass"
        Else
            ws.Cells(i, 3).Value = "Fail"
        End If
    Next i
End Sub
\`\`\`

## Copy and Paste
\`\`\`vba
' Copy range
Range("A1:C10").Copy Destination:=Range("E1")

' Copy values only
Range("A1:C10").Copy
Range("E1").PasteSpecial xlPasteValues

' Copy formatting
Range("E1").PasteSpecial xlPasteFormats
\`\`\`
      `
    },
    arrays: {
      title: 'Arrays & Collections',
      content: `
# Arrays and Collections

## Static Arrays
\`\`\`vba
Sub StaticArrayExample()
    Dim numbers(1 To 5) As Integer
    Dim i As Integer

    ' Fill array
    For i = 1 To 5
        numbers(i) = i * 10
    Next i

    ' Read array
    For i = 1 To 5
        Debug.Print numbers(i)
    Next i
End Sub
\`\`\`

## Dynamic Arrays
\`\`\`vba
Sub DynamicArrayExample()
    Dim numbers() As Integer
    Dim size As Integer

    size = 10
    ReDim numbers(1 To size)

    ' Resize and preserve data
    ReDim Preserve numbers(1 To size + 5)
End Sub
\`\`\`

## Multi-Dimensional Arrays
\`\`\`vba
Sub TwoDimensionalArray()
    Dim matrix(1 To 3, 1 To 3) As Integer
    Dim i As Integer, j As Integer

    ' Fill matrix
    For i = 1 To 3
        For j = 1 To 3
            matrix(i, j) = i * j
        Next j
    Next i
End Sub
\`\`\`

## Array Functions
\`\`\`vba
Sub ArrayFunctions()
    Dim arr As Variant

    ' Create array
    arr = Array(1, 2, 3, 4, 5)

    ' Array bounds
    Debug.Print LBound(arr)  ' Lower bound
    Debug.Print UBound(arr)  ' Upper bound

    ' Split string into array
    Dim words() As String
    words = Split("Hello World VBA", " ")

    ' Join array into string
    Dim sentence As String
    sentence = Join(words, "-")  ' Hello-World-VBA
End Sub
\`\`\`

## Collections
\`\`\`vba
Sub CollectionExample()
    Dim myCollection As New Collection

    ' Add items
    myCollection.Add "Apple"
    myCollection.Add "Banana"
    myCollection.Add "Cherry"

    ' Add with key
    myCollection.Add "Orange", "fruit1"

    ' Count items
    Debug.Print myCollection.Count

    ' Access items
    Debug.Print myCollection(1)          ' First item
    Debug.Print myCollection("fruit1")   ' By key

    ' Loop through collection
    Dim item As Variant
    For Each item In myCollection
        Debug.Print item
    Next item

    ' Remove item
    myCollection.Remove 1
End Sub
\`\`\`

## Dictionary Object
\`\`\`vba
Sub DictionaryExample()
    ' Requires: Tools → References → Microsoft Scripting Runtime
    Dim dict As New Dictionary

    ' Add items
    dict.Add "Name", "John"
    dict.Add "Age", 30
    dict.Add "City", "New York"

    ' Check if key exists
    If dict.Exists("Name") Then
        Debug.Print dict("Name")
    End If

    ' Update value
    dict("Age") = 31

    ' Loop through keys
    Dim key As Variant
    For Each key In dict.Keys
        Debug.Print key & ": " & dict(key)
    Next key

    ' Remove item
    dict.Remove "City"
End Sub
\`\`\`
      `
    },
    errors: {
      title: 'Error Handling',
      content: `
# Error Handling

## On Error Statement
\`\`\`vba
Sub ErrorHandlingBasic()
    On Error GoTo ErrorHandler

    ' Code that might cause error
    Dim x As Integer
    x = 100 / 0  ' Division by zero

    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description
End Sub
\`\`\`

## Error Object
\`\`\`vba
Sub ErrorObjectExample()
    On Error GoTo ErrorHandler

    ' Trigger error
    Err.Raise 1000, "MyFunction", "Custom error message"

    Exit Sub

ErrorHandler:
    Debug.Print "Error Number: " & Err.Number
    Debug.Print "Error Source: " & Err.Source
    Debug.Print "Error Description: " & Err.Description
    Err.Clear  ' Clear the error
End Sub
\`\`\`

## On Error Resume Next
\`\`\`vba
Sub OnErrorResumeNextExample()
    On Error Resume Next

    ' Try to open file
    Workbooks.Open "C:\\NonExistent.xlsx"

    If Err.Number <> 0 Then
        MsgBox "Could not open file: " & Err.Description
        Err.Clear
    End If

    On Error GoTo 0  ' Turn off error handling
End Sub
\`\`\`

## Proper Error Handling Pattern
\`\`\`vba
Sub ProperErrorHandling()
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim filePath As String

    filePath = "C:\\Data\\Report.xlsx"

    ' Your code here
    Set ws = Workbooks.Open(filePath).Worksheets(1)

    ' Process data
    ws.Range("A1").Value = "Report"

CleanUp:
    ' Clean up code (always runs)
    If Not ws Is Nothing Then
        ws.Parent.Close SaveChanges:=True
    End If
    Exit Sub

ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description
    Resume CleanUp
End Sub
\`\`\`

## Custom Error Messages
\`\`\`vba
Sub CustomErrors()
    On Error GoTo ErrorHandler

    Dim value As Integer
    value = -5

    If value < 0 Then
        Err.Raise Number:=vbObjectError + 1000, _
                  Description:="Value cannot be negative"
    End If

    Exit Sub

ErrorHandler:
    Select Case Err.Number
        Case vbObjectError + 1000
            MsgBox "Validation Error: " & Err.Description
        Case Else
            MsgBox "Unexpected Error: " & Err.Description
    End Select
End Sub
\`\`\`

## Logging Errors
\`\`\`vba
Sub LogError(errorNum As Long, errorDesc As String, procedureName As String)
    Dim ws As Worksheet
    Dim lastRow As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("ErrorLog")

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = "ErrorLog"
        ws.Range("A1:D1").Value = Array("Date", "Time", "Procedure", "Error", "Description")
    End If

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1

    ws.Cells(lastRow, 1).Value = Date
    ws.Cells(lastRow, 2).Value = Time
    ws.Cells(lastRow, 3).Value = procedureName
    ws.Cells(lastRow, 4).Value = errorNum
    ws.Cells(lastRow, 5).Value = errorDesc
End Sub
\`\`\`
      `
    },
    files: {
      title: 'File Operations',
      content: `
# File Operations

## File System Object
\`\`\`vba
Sub FileSystemObjectExample()
    ' Requires: Tools → References → Microsoft Scripting Runtime
    Dim fso As New FileSystemObject
    Dim file As File
    Dim folder As Folder

    ' Check if file exists
    If fso.FileExists("C:\\Data\\Report.xlsx") Then
        MsgBox "File exists"
    End If

    ' Check if folder exists
    If fso.FolderExists("C:\\Data") Then
        MsgBox "Folder exists"
    End If

    ' Create folder
    If Not fso.FolderExists("C:\\Output") Then
        fso.CreateFolder "C:\\Output"
    End If

    ' Get file object
    Set file = fso.GetFile("C:\\Data\\Report.xlsx")
    Debug.Print "Size: " & file.Size
    Debug.Print "Created: " & file.DateCreated
    Debug.Print "Modified: " & file.DateLastModified

    ' Copy file
    fso.CopyFile "C:\\Data\\Report.xlsx", "C:\\Backup\\Report.xlsx"

    ' Move file
    fso.MoveFile "C:\\Data\\Old.xlsx", "C:\\Archive\\Old.xlsx"

    ' Delete file
    fso.DeleteFile "C:\\Temp\\Temp.xlsx"
End Sub
\`\`\`

## Reading Text Files
\`\`\`vba
Sub ReadTextFile()
    Dim fso As New FileSystemObject
    Dim ts As TextStream
    Dim line As String

    Set ts = fso.OpenTextFile("C:\\Data\\input.txt", ForReading)

    ' Read line by line
    Do While Not ts.AtEndOfStream
        line = ts.ReadLine
        Debug.Print line
    Loop

    ts.Close
End Sub

Sub ReadEntireFile()
    Dim fso As New FileSystemObject
    Dim ts As TextStream
    Dim content As String

    Set ts = fso.OpenTextFile("C:\\Data\\input.txt", ForReading)
    content = ts.ReadAll
    ts.Close

    MsgBox content
End Sub
\`\`\`

## Writing Text Files
\`\`\`vba
Sub WriteTextFile()
    Dim fso As New FileSystemObject
    Dim ts As TextStream

    ' Create new file (overwrites existing)
    Set ts = fso.CreateTextFile("C:\\Output\\output.txt", True)

    ts.WriteLine "First line"
    ts.WriteLine "Second line"
    ts.Write "Text without newline"

    ts.Close
End Sub

Sub AppendToFile()
    Dim fso As New FileSystemObject
    Dim ts As TextStream

    ' Open for appending
    Set ts = fso.OpenTextFile("C:\\Output\\output.txt", ForAppending)

    ts.WriteLine "Appended line"

    ts.Close
End Sub
\`\`\`

## Directory Listing
\`\`\`vba
Sub ListFilesInFolder()
    Dim fso As New FileSystemObject
    Dim folder As Folder
    Dim file As File
    Dim ws As Worksheet
    Dim row As Long

    Set ws = ThisWorkbook.Worksheets("FileList")
    Set folder = fso.GetFolder("C:\\Data")

    row = 1
    For Each file In folder.Files
        ws.Cells(row, 1).Value = file.Name
        ws.Cells(row, 2).Value = file.Size
        ws.Cells(row, 3).Value = file.DateLastModified
        row = row + 1
    Next file
End Sub

Sub ListSubFolders()
    Dim fso As New FileSystemObject
    Dim folder As Folder
    Dim subfolder As Folder

    Set folder = fso.GetFolder("C:\\Data")

    For Each subfolder In folder.SubFolders
        Debug.Print subfolder.Name
    Next subfolder
End Sub
\`\`\`

## File Dialog
\`\`\`vba
Sub OpenFileDialog()
    Dim fd As FileDialog
    Dim selectedFile As String

    Set fd = Application.FileDialog(msoFileDialogFilePicker)

    With fd
        .Title = "Select a file"
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xlsx; *.xls"
        .Filters.Add "All Files", "*.*"
        .AllowMultiSelect = False

        If .Show = -1 Then
            selectedFile = .SelectedItems(1)
            MsgBox "You selected: " & selectedFile
        End If
    End With
End Sub

Sub SaveFileDialog()
    Dim fd As FileDialog
    Dim saveAsPath As String

    Set fd = Application.FileDialog(msoFileDialogSaveAs)

    With fd
        .Title = "Save file as"
        If .Show = -1 Then
            saveAsPath = .SelectedItems(1)
            ThisWorkbook.SaveAs saveAsPath
        End If
    End With
End Sub
\`\`\`
      `
    },
    advanced: {
      title: 'Advanced Topics',
      content: `
# Advanced Topics

## Events
\`\`\`vba
' In ThisWorkbook module
Private Sub Workbook_Open()
    MsgBox "Workbook opened!"
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Dim response As VbMsgBoxResult
    response = MsgBox("Save changes?", vbYesNoCancel)

    If response = vbCancel Then
        Cancel = True
    End If
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    ' Code before save
End Sub

' In Worksheet module
Private Sub Worksheet_Change(ByVal Target As Range)
    If Not Intersect(Target, Range("A:A")) Is Nothing Then
        MsgBox "Column A was changed"
    End If
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    ' Highlight selected row
    Cells.Interior.ColorIndex = xlNone
    Target.EntireRow.Interior.Color = RGB(200, 200, 255)
End Sub
\`\`\`

## Classes
\`\`\`vba
' Class Module: Person
Private pName As String
Private pAge As Integer

' Property Let (setter)
Public Property Let Name(value As String)
    pName = value
End Property

' Property Get (getter)
Public Property Get Name() As String
    Name = pName
End Property

Public Property Let Age(value As Integer)
    If value >= 0 Then
        pAge = value
    End If
End Property

Public Property Get Age() As Integer
    Age = pAge
End Property

' Method
Public Function GetInfo() As String
    GetInfo = pName & " is " & pAge & " years old"
End Function

' Using the class
Sub UsePersonClass()
    Dim person As New Person
    person.Name = "John"
    person.Age = 30
    MsgBox person.GetInfo()
End Sub
\`\`\`

## API Calls
\`\`\`vba
' At top of module
#If VBA7 Then
    Private Declare PtrSafe Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" _
        (ByVal lpBuffer As String, nSize As Long) As Long
#Else
    Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" _
        (ByVal lpBuffer As String, nSize As Long) As Long
#End If

Function GetCurrentUser() As String
    Dim buffer As String
    Dim size As Long

    buffer = Space(255)
    size = Len(buffer)

    If GetUserName(buffer, size) Then
        GetCurrentUser = Left(buffer, size - 1)
    End If
End Function
\`\`\`

## Regular Expressions
\`\`\`vba
Sub RegexExample()
    ' Requires: Tools → References → Microsoft VBScript Regular Expressions
    Dim regex As New RegExp
    Dim matches As MatchCollection
    Dim match As match

    ' Email validation
    regex.Pattern = "^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\\.[a-zA-Z]{2,}$"
    regex.IgnoreCase = True

    If regex.Test("user@example.com") Then
        MsgBox "Valid email"
    End If

    ' Find all matches
    regex.Pattern = "\\d+"
    regex.Global = True

    Set matches = regex.Execute("Order 123 costs $456")

    For Each match In matches
        Debug.Print match.value
    Next match

    ' Replace
    regex.Pattern = "\\d+"
    Debug.Print regex.Replace("Price: 100", "200")  ' Price: 200
End Sub
\`\`\`

## Late Binding
\`\`\`vba
Sub LateBingingExample()
    ' No reference needed
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    dict.Add "Key1", "Value1"
    MsgBox dict("Key1")
End Sub
\`\`\`

## Application Performance
\`\`\`vba
Sub OptimizePerformance()
    Dim startTime As Double
    startTime = Timer

    ' Turn off screen updating
    Application.ScreenUpdating = False

    ' Turn off automatic calculation
    Application.Calculation = xlCalculationManual

    ' Turn off events
    Application.EnableEvents = False

    ' Your code here
    Dim i As Long
    For i = 1 To 10000
        Cells(i, 1).Value = i
    Next i

    ' Restore settings
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True

    Debug.Print "Time elapsed: " & Timer - startTime & " seconds"
End Sub
\`\`\`

## SQL with ADO
\`\`\`vba
Sub QueryDatabase()
    ' Requires: Tools → References → Microsoft ActiveX Data Objects
    Dim conn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    Dim sql As String

    ' Connection string
    conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                           "Data Source=C:\\Data\\Database.accdb;"
    conn.Open

    ' Execute query
    sql = "SELECT * FROM Customers WHERE City = 'London'"
    rs.Open sql, conn

    ' Process results
    Do While Not rs.EOF
        Debug.Print rs.Fields("CustomerName").value
        rs.MoveNext
    Loop

    ' Clean up
    rs.Close
    conn.Close
    Set rs = Nothing
    Set conn = Nothing
End Sub
\`\`\`
      `
    },
    tips: {
      title: 'Best Practices & Tips',
      content: `
# Best Practices & Tips

## Code Organization
\`\`\`vba
' Use meaningful names
Dim customerName As String      ' Good
Dim cn As String               ' Avoid

' Use constants for magic numbers
Const TAX_RATE As Double = 0.08  ' Good
total = price * 1.08             ' Avoid

' Group related code
' ==================
' Data Processing Functions
' ==================

Function ProcessData()
    ' Code here
End Function

Function ValidateData()
    ' Code here
End Function
\`\`\`

## Error Handling
\`\`\`vba
' Always include error handling for critical code
Sub CriticalOperation()
    On Error GoTo ErrorHandler

    ' Your code

    Exit Sub
ErrorHandler:
    MsgBox "Error: " & Err.Description
    LogError Err.Number, Err.Description, "CriticalOperation"
End Sub
\`\`\`

## Performance Optimization

### 1. Use With Statement
\`\`\`vba
' Bad
Range("A1").Value = "Name"
Range("A1").Font.Bold = True
Range("A1").Interior.Color = RGB(255, 255, 0)

' Good
With Range("A1")
    .Value = "Name"
    .Font.Bold = True
    .Interior.Color = RGB(255, 255, 0)
End With
\`\`\`

### 2. Avoid Select/Activate
\`\`\`vba
' Bad
Range("A1").Select
Selection.Value = "Hello"

' Good
Range("A1").Value = "Hello"
\`\`\`

### 3. Use Arrays for Large Data
\`\`\`vba
' Bad - Slow for large datasets
For i = 1 To 10000
    Cells(i, 1).Value = i
Next i

' Good - Much faster
Dim arr(1 To 10000, 1 To 1) As Long
For i = 1 To 10000
    arr(i, 1) = i
Next i
Range("A1:A10000").Value = arr
\`\`\`

### 4. Turn Off Features During Bulk Operations
\`\`\`vba
Sub BulkOperation()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    ' Your bulk operations

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub
\`\`\`

## Code Documentation
\`\`\`vba
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function: CalculateTax
' Purpose: Calculate tax amount based on price and tax rate
' Parameters:
'   price - The base price (Double)
'   taxRate - Tax rate as decimal (e.g., 0.08 for 8%)
' Returns: Tax amount (Double)
' Author: John Doe
' Date: 2024-01-01
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function CalculateTax(price As Double, taxRate As Double) As Double
    CalculateTax = price * taxRate
End Function
\`\`\`

## Debugging Techniques

### 1. Debug.Print
\`\`\`vba
Sub DebugExample()
    Dim x As Integer
    x = 10
    Debug.Print "Value of x: " & x
End Sub
\`\`\`

### 2. Immediate Window
- Press Ctrl+G to open
- Type ? variable_name to see value
- Execute code directly

### 3. Breakpoints
- Click in left margin to set breakpoint
- F8 to step through code
- F5 to continue

### 4. Watch Window
- Add variables to watch
- See values change in real-time

## Common Pitfalls

### 1. Not Using Option Explicit
\`\`\`vba
' At top of every module
Option Explicit

' Forces variable declaration
' Catches typos
\`\`\`

### 2. Object Not Set
\`\`\`vba
' Always check before using objects
Dim ws As Worksheet
Set ws = Worksheets("Data")

If Not ws Is Nothing Then
    ws.Range("A1").Value = "Hello"
End If
\`\`\`

### 3. Infinite Loops
\`\`\`vba
' Always ensure loop will exit
Dim i As Integer
i = 1

Do While i <= 10
    Debug.Print i
    i = i + 1  ' Don't forget to increment!
Loop
\`\`\`

## Security

### 1. Protect Code
- Tools → VBAProject Properties → Protection
- Set password to protect code

### 2. Validate Input
\`\`\`vba
Function ValidateInput(userInput As String) As Boolean
    If Len(userInput) = 0 Then
        MsgBox "Input cannot be empty"
        ValidateInput = False
        Exit Function
    End If

    ValidateInput = True
End Function
\`\`\`

### 3. Avoid Hardcoded Passwords
\`\`\`vba
' Bad
conn.ConnectionString = "...Password=secret123..."

' Better - Use Windows Authentication or secure storage
conn.ConnectionString = "...Integrated Security=SSPI..."
\`\`\`

## Useful Shortcuts
- Alt+F11: Open VBA Editor
- F5: Run code
- F8: Step through code
- Ctrl+G: Immediate window
- Ctrl+Space: Auto-complete
- Ctrl+Shift+F9: Clear all breakpoints
- F2: Object Browser
      `
    }
  };

  return (
    <div style={styles.container}>
      <aside style={styles.sidebar}>
        <div style={styles.header}>
          <h1 style={styles.title}>VBA eBook</h1>
          <p style={styles.subtitle}>Complete Programming Guide</p>
        </div>

        <nav style={styles.nav}>
          {Object.entries(chapters).map(([key, chapter]) => (
            <button
              key={key}
              onClick={() => setActiveChapter(key)}
              style={{
                ...styles.navButton,
                ...(activeChapter === key ? styles.navButtonActive : {})
              }}
            >
              {chapter.title}
            </button>
          ))}
        </nav>
      </aside>

      <main style={styles.main}>
        <div style={styles.content}>
          <Content text={chapters[activeChapter].content} />
        </div>
      </main>
    </div>
  );
}

function Content({ text }) {
  const renderContent = () => {
    const lines = text.trim().split('\n');
    const elements = [];
    let codeBlock = [];
    let inCodeBlock = false;
    let codeLanguage = '';

    lines.forEach((line, idx) => {
      if (line.startsWith('```')) {
        if (!inCodeBlock) {
          inCodeBlock = true;
          codeLanguage = line.substring(3);
        } else {
          elements.push(
            <pre key={`code-${idx}`} style={styles.codeBlock}>
              <code style={styles.code}>{codeBlock.join('\n')}</code>
            </pre>
          );
          codeBlock = [];
          inCodeBlock = false;
        }
      } else if (inCodeBlock) {
        codeBlock.push(line);
      } else if (line.startsWith('# ')) {
        elements.push(<h1 key={idx} style={styles.h1}>{line.substring(2)}</h1>);
      } else if (line.startsWith('## ')) {
        elements.push(<h2 key={idx} style={styles.h2}>{line.substring(3)}</h2>);
      } else if (line.startsWith('### ')) {
        elements.push(<h3 key={idx} style={styles.h3}>{line.substring(4)}</h3>);
      } else if (line.startsWith('- ')) {
        elements.push(<li key={idx} style={styles.li}>{line.substring(2)}</li>);
      } else if (line.includes('`') && !line.startsWith('```')) {
        const parts = line.split('`');
        elements.push(
          <p key={idx} style={styles.p}>
            {parts.map((part, i) =>
              i % 2 === 0 ? part : <code key={i} style={styles.inlineCode}>{part}</code>
            )}
          </p>
        );
      } else if (line.trim()) {
        elements.push(<p key={idx} style={styles.p}>{line}</p>);
      }
    });

    return elements;
  };

  return <div>{renderContent()}</div>;
}

const styles = {
  container: {
    display: 'flex',
    minHeight: '100vh',
    fontFamily: '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif',
    margin: 0,
    padding: 0,
  },
  sidebar: {
    width: '280px',
    backgroundColor: '#1e293b',
    color: '#fff',
    padding: '20px',
    overflowY: 'auto',
    position: 'fixed',
    height: '100vh',
  },
  header: {
    marginBottom: '30px',
    borderBottom: '2px solid #3b82f6',
    paddingBottom: '15px',
  },
  title: {
    fontSize: '28px',
    margin: '0 0 5px 0',
    color: '#3b82f6',
  },
  subtitle: {
    fontSize: '14px',
    margin: 0,
    color: '#94a3b8',
  },
  nav: {
    display: 'flex',
    flexDirection: 'column',
    gap: '5px',
  },
  navButton: {
    backgroundColor: 'transparent',
    border: 'none',
    color: '#cbd5e1',
    padding: '12px 15px',
    textAlign: 'left',
    cursor: 'pointer',
    fontSize: '14px',
    borderRadius: '6px',
    transition: 'all 0.2s',
  },
  navButtonActive: {
    backgroundColor: '#3b82f6',
    color: '#fff',
  },
  main: {
    marginLeft: '280px',
    flex: 1,
    backgroundColor: '#f8fafc',
  },
  content: {
    maxWidth: '900px',
    margin: '0 auto',
    padding: '40px 30px',
    backgroundColor: '#fff',
    minHeight: '100vh',
    boxShadow: '0 0 20px rgba(0,0,0,0.05)',
  },
  h1: {
    fontSize: '32px',
    color: '#0f172a',
    marginTop: '30px',
    marginBottom: '15px',
    borderBottom: '3px solid #3b82f6',
    paddingBottom: '10px',
  },
  h2: {
    fontSize: '24px',
    color: '#1e293b',
    marginTop: '25px',
    marginBottom: '12px',
  },
  h3: {
    fontSize: '18px',
    color: '#334155',
    marginTop: '20px',
    marginBottom: '10px',
  },
  p: {
    fontSize: '16px',
    lineHeight: '1.7',
    color: '#475569',
    marginBottom: '15px',
  },
  li: {
    fontSize: '16px',
    lineHeight: '1.7',
    color: '#475569',
    marginBottom: '8px',
    marginLeft: '20px',
  },
  codeBlock: {
    backgroundColor: '#1e293b',
    padding: '20px',
    borderRadius: '8px',
    overflowX: 'auto',
    marginBottom: '20px',
    border: '1px solid #334155',
  },
  code: {
    color: '#e2e8f0',
    fontSize: '14px',
    fontFamily: '"Fira Code", "Courier New", monospace',
    lineHeight: '1.6',
  },
  inlineCode: {
    backgroundColor: '#e0f2fe',
    color: '#0369a1',
    padding: '2px 6px',
    borderRadius: '4px',
    fontSize: '14px',
    fontFamily: '"Fira Code", "Courier New", monospace',
  },
};
