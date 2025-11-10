Option Explicit

'Module-level constants
Private Const SCOPING_SHEET As String = "Scoping"
Private Const INPUT_SHEET As String = "Input"
Private Const TEMP_SHEET As String = "tempsheet"
Private Const EXCLUSION_COLOR As Long = 14805212  ' EXCLUSION_COLOR
Private Const MAX_DATA_ROW As Long = 9000

Public Sub Scoping_Exclusions()

Dim rowCount As Integer
Dim lastRow As Integer
Dim lastCol As Integer
Dim colCount As Integer
Dim BSRow As Variant
Dim rowTotal As Double
Dim componentAddress As String
Dim businessUnitAddress As String
Dim fsliAddress As String
Dim doubleQuotes As String
Dim copyRange As String
Dim totalRange As Range
Dim currentRange As Range
Dim rowRange As Range
Dim cellColor As Long
Dim currentCell As Range



Application.ScreenUpdating = False

Ask_Row:
'Ask what row the IS starts
Dim Message, Title, Default
Message = "Enter the row number where the IS FSLIs starts"   ' Set prompt.
'Display message, title, and default value.
BSRow = InputBox(Message, Title, Default)
If BSRow = "" Then
    Exit Sub ' This ends the macro if 'cancel' is selected
End If

'Replay the message box if value entered is not a number greater than10
If BSRow < 10 Then
    MsgBox ("Please enter a numerical value greater than 10")
    GoTo Ask_Row
Else:
    If IsNumeric(BSRow) = False Then
       MsgBox ("Please enter a numerical value greater than 10")
       GoTo Ask_Row
   End If
End If

BSRow = BSRow - 1

'Find the Total Range of the pivot table
rowCount = Sheets(SCOPING_SHEET).Range("C7").End(xlDown).Row
colCount = Sheets(SCOPING_SHEET).Range("D6").End(xlToRight).Column
Set totalRange = Sheets(SCOPING_SHEET).Range(Worksheets(SCOPING_SHEET).Cells(7, 4), Worksheets(SCOPING_SHEET).Cells(rowCount, colCount))
lastRow = rowCount
lastCol = colCount + 1

'Remove highlights in scoping tab
totalRange.Interior.ColorIndex = xlNone

'Create new sheet called "tempsheet"
Sheets.Add.Name = TEMP_SHEET

rowCount = 5
colCount = 4
doubleQuotes = Chr(34)

Application.Calculation = xlCalculationManual

'Goes through the sheet again looking for the FSLIs Excluded where BU is out of scope
Do Until IsEmpty(Sheets(SCOPING_SHEET).Cells(rowCount, colCount).End(xlToRight))
    componentAddress = Sheets(SCOPING_SHEET).Cells(rowCount, colCount).Address
    businessUnitAddress = Sheets(SCOPING_SHEET).Cells(rowCount + 1, colCount).Address(True, False)
    
    fsliAddress = "$C7"
    Set currentRange = Worksheets(SCOPING_SHEET).Range(Worksheets(SCOPING_SHEET).Cells(7, colCount), Worksheets(SCOPING_SHEET).Cells(lastRow, (Worksheets(SCOPING_SHEET).Cells(5, colCount).End(xlToRight).Column) - 1))
    
    copyRange = currentRange.Address
    
    'This exclusion formula allows for all FSLIs across both BS and IS to be excluded for a given BU - compare to exclusion formulas used later(excluded if countif returns value greater than 0)
    Sheets(TEMP_SHEET).Cells(rowCount + 2, colCount) = "=COUNTIFS(Input!$E$1:$E$9000,Scoping!" & fsliAddress & ",Input!$D$1:$D$9000, Scoping!" & businessUnitAddress & ",Input!$C$1:$C$9000,Scoping!" & componentAddress & ") + COUNTIFS(Input!$E$1:$E$9000," & doubleQuotes & "All FSLIs" & doubleQuotes & ",Input!$D$1:$D$9000,Scoping!" & businessUnitAddress & ",Input!$C$1:$C$9000,Scoping!" & componentAddress & ") + COUNTIFS(Input!$E$1:$E$9000," & doubleQuotes & "All BS" & doubleQuotes & ",Input!$D$1:$D$9000,Scoping!" & businessUnitAddress & ",Input!$C$1:$C$9000,Scoping!" & componentAddress & ") + COUNTIFS(Input!$D$1:$D$9000," & doubleQuotes & "All" & doubleQuotes & ",Input!$C$1:$C$9000, Scoping!" & componentAddress & ")"
    Sheets(TEMP_SHEET).Cells(rowCount + 2, colCount).Copy Destination:=Sheets(TEMP_SHEET).Range(copyRange)
    
    Application.CutCopyMode = False
    
    rowCount = Sheets(SCOPING_SHEET).Cells(rowCount, colCount).End(xlToRight).Row
    colCount = Sheets(SCOPING_SHEET).Cells(rowCount, colCount).End(xlToRight).Column
Loop

Application.Calculation = xlCalculationAutomatic

copyRange = totalRange.Address

cellColor = EXCLUSION_COLOR

'Make cells a bluish gray if it is part of the Total of FSLIs Excluded where BU is out of scope
For Each currentCell In Sheets(TEMP_SHEET).Range(copyRange)
    If currentCell.Value > 0 Then
        Sheets(SCOPING_SHEET).Cells(currentCell.Row, currentCell.Column).Interior.Color = cellColor
    End If
Next

'Adds up the total of grayed out cells in each row and inserts the total excluded past the end of the pivot table for the all excluded BUs
For Each rowRange In totalRange.Rows
    rowTotal = 0
    copyRange = rowRange.Address
    For Each currentCell In Sheets(SCOPING_SHEET).Range(copyRange)
        If currentCell.Interior.ColorIndex <> xlNone Then
            rowTotal = rowTotal + currentCell.Value
        End If
    Next currentCell
    copyRange = rowRange.Address
    Sheets(SCOPING_SHEET).Cells(rowRange.Row, lastCol + 3).Value = rowTotal
    Sheets(SCOPING_SHEET).Cells(rowRange.Row, lastCol + 3).NumberFormat = "#,##0; (#,##0)"
    rowCount = rowRange.Row
Next

'Find the next Scoping Component section, starting with the first in the pivot table
rowCount = 5
colCount = 4
doubleQuotes = Chr(34)

Application.Calculation = xlCalculationManual

Do Until IsEmpty(Sheets(SCOPING_SHEET).Cells(rowCount, colCount).End(xlToRight))
    componentAddress = Sheets(SCOPING_SHEET).Cells(rowCount, colCount).Address
    businessUnitAddress = Sheets(SCOPING_SHEET).Cells(rowCount + 1, colCount).Address(True, False)
    
    'BS Section
    fsliAddress = "$C7"
    Set currentRange = Worksheets(SCOPING_SHEET).Range(Worksheets(SCOPING_SHEET).Cells(7, colCount), Worksheets(SCOPING_SHEET).Cells(BSRow, (Worksheets(SCOPING_SHEET).Cells(5, colCount).End(xlToRight).Column) - 1))
    
    copyRange = currentRange.Address
    
    'This exclusion formula allows for all BS exclusions to be handled for a given BU (excluded if countif returns value greater than 0) in the BS range
    Sheets(TEMP_SHEET).Cells(rowCount + 2, colCount) = "=COUNTIFS(Input!$E$1:$E$9000,Scoping!" & fsliAddress & ",Input!$D$1:$D$9000, Scoping!" & businessUnitAddress & ",Input!$C$1:$C$9000,Scoping!" & componentAddress & ") + COUNTIFS(Input!$E$1:$E$9000," & doubleQuotes & "All FSLIs" & doubleQuotes & ",Input!$D$1:$D$9000,Scoping!" & businessUnitAddress & ",Input!$C$1:$C$9000,Scoping!" & componentAddress & ") + COUNTIFS(Input!$E$1:$E$9000," & doubleQuotes & "All BS" & doubleQuotes & ",Input!$D$1:$D$9000,Scoping!" & businessUnitAddress & ",Input!$C$1:$C$9000,Scoping!" & componentAddress & ") + COUNTIFS(Input!$D$1:$D$9000," & doubleQuotes & "All" & doubleQuotes & ",Input!$C$1:$C$9000, Scoping!" & componentAddress & ")"
    Sheets(TEMP_SHEET).Cells(rowCount + 2, colCount).Copy Destination:=Sheets(TEMP_SHEET).Range(copyRange)
    
    'IS Section
    fsliAddress = Worksheets(SCOPING_SHEET).Cells(BSRow + 1, 3).Address(False, True)
    Set currentRange = Worksheets(SCOPING_SHEET).Range(Worksheets(SCOPING_SHEET).Cells(BSRow + 1, colCount), Worksheets(SCOPING_SHEET).Cells(lastRow, (Worksheets(SCOPING_SHEET).Cells(5, colCount).End(xlToRight).Column) - 1))
    
    copyRange = currentRange.Address
    
    'This exclusion formula allows for all IS exclusions to be handled for a given BU (excluded if countif returns value greater than 0) in the BS range
    Sheets(TEMP_SHEET).Cells(BSRow + 1, colCount) = "=COUNTIFS(Input!$E$1:$E$9000,Scoping!" & fsliAddress & ",Input!$D$1:$D$9000,Scoping!" & businessUnitAddress & ",Input!$C$1:$C$9000,Scoping!" & componentAddress & ") + COUNTIFS(Input!$E$1:$E$9000," & doubleQuotes & "All FSLIs" & doubleQuotes & ",Input!$D$1:$D$9000,Scoping!" & businessUnitAddress & ",Input!$C$1:$C$9000,Scoping!" & componentAddress & ") + COUNTIFS(Input!$E$1:$E$9000," & doubleQuotes & "All IS" & doubleQuotes & ",Input!$D$1:$D$9000,Scoping!" & businessUnitAddress & ",Input!$C$1:$C$9000,Scoping!" & componentAddress & ")+ COUNTIFS(Input!$D$1:$D$9000," & doubleQuotes & "All" & doubleQuotes & ",Input!$C$1:$C$9000,Scoping!" & componentAddress & ")"
    Sheets(TEMP_SHEET).Cells(BSRow + 1, colCount).Copy Destination:=Sheets(TEMP_SHEET).Range(copyRange)
    
    Application.CutCopyMode = False
    
    rowCount = Sheets(SCOPING_SHEET).Cells(rowCount, colCount).End(xlToRight).Row
    colCount = Sheets(SCOPING_SHEET).Cells(rowCount, colCount).End(xlToRight).Column
Loop

Application.Calculation = xlCalculationAutomatic

copyRange = totalRange.Address

'Make all cells gray where greater than 0
For Each currentCell In Sheets(TEMP_SHEET).Range(copyRange)
    If currentCell.Value > 0 Then
        currentCell.Interior.Color = EXCLUSION_COLOR
    End If
Next

'Paste the grey highlighting format onto the pivot table
Sheets(TEMP_SHEET).Range(copyRange).Copy
totalRange.PasteSpecial xlPasteFormats

Application.CutCopyMode = False

'Delete the temporary tab
Application.DisplayAlerts = False
Sheets(TEMP_SHEET).Delete
Application.DisplayAlerts = True

'On the scoping tab, all highlights are removed from all cells equal to 0
For Each currentCell In totalRange
        If currentCell.Value = 0 Then
            currentCell.Interior.ColorIndex = xlNone
        End If
Next

'Adds up the total of grayed out cells in each row and inserts the total excluded past the end of the pivot table
For Each rowRange In totalRange.Rows
    rowTotal = 0
    copyRange = rowRange.Address
    For Each currentCell In Worksheets(SCOPING_SHEET).Range(copyRange)
        If currentCell.Interior.ColorIndex <> xlNone Then
                rowTotal = rowTotal + currentCell.Value
            End If
    Next currentCell
    copyRange = rowRange.Address
    Worksheets(SCOPING_SHEET).Cells(rowRange.Row, lastCol + 2).Value = rowTotal
    Worksheets(SCOPING_SHEET).Cells(rowRange.Row, lastCol + 2).NumberFormat = "#,##0; (#,##0)"
    Worksheets(SCOPING_SHEET).Cells(rowRange.Row, lastCol + 3) = "=" & Worksheets(SCOPING_SHEET).Cells(rowRange.Row, lastCol + 2).Address & "/" & Worksheets(SCOPING_SHEET).Cells(rowRange.Row, lastCol).Address
    Worksheets(SCOPING_SHEET).Cells(rowRange.Row, lastCol + 3).NumberFormat = "0.0%"

    rowCount = rowRange.Row
Next

'Adds the two column headers at the end showing the total excluded and percentage of total excluded
Worksheets(SCOPING_SHEET).Cells(Worksheets(SCOPING_SHEET).Cells(rowCount, lastCol + 2).End(xlUp).Row - 2, lastCol + 2).Value = "Total Excluded by FSLI"
Worksheets(SCOPING_SHEET).Cells(Worksheets(SCOPING_SHEET).Cells(rowCount, lastCol + 2).End(xlUp).Row - 2, lastCol + 2).Font.Bold = True

Worksheets(SCOPING_SHEET).Cells(Worksheets(SCOPING_SHEET).Cells(rowCount, lastCol + 2).End(xlUp).Row - 2, lastCol + 3).Value = "Percentage of total Excluded"
Worksheets(SCOPING_SHEET).Cells(Worksheets(SCOPING_SHEET).Cells(rowCount, lastCol + 2).End(xlUp).Row - 2, lastCol + 3).Font.Bold = True

totalRange.NumberFormat = "#,##0; (#,##0)"

Application.ScreenUpdating = True

Sheets(SCOPING_SHEET).Range("A1").Activate

MsgBox ("The Macro has finished running")

End Sub

Sub Clean_Scoping()

Dim rowCount As Integer
Dim lastCol As Integer
Dim colCount As Integer
Dim deleteRange As Range
Dim totalRange As Range

Application.ScreenUpdating = False

'Find the Total Range of the Pivot Table
rowCount = Sheets(SCOPING_SHEET).Range("C7").End(xlDown).Row
colCount = Sheets(SCOPING_SHEET).Range("D6").End(xlToRight).Column
Set totalRange = Sheets(SCOPING_SHEET).Range(Worksheets(SCOPING_SHEET).Cells(7, 4), Worksheets(SCOPING_SHEET).Cells(rowCount, colCount))

lastCol = colCount + 1

'Remove highlights in scoping tab
totalRange.Interior.ColorIndex = xlNone

Set deleteRange = Sheets(SCOPING_SHEET).Range(Worksheets(SCOPING_SHEET).Cells(5, colCount + 3), Worksheets(SCOPING_SHEET).Cells(rowCount, colCount + 6))
deleteRange.Clear

Application.ScreenUpdating = True

End Sub

