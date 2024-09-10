## Project VBA Code
You can view all the VBA code used in this project here:


' This procedure imports multiple Excel files and creates a new sheet for each file in the current workbook.
Public Sub ImportExcelFiles()
    Dim Wb As Workbook
    Dim sheetName As String
    Dim openwbs() As Variant
    Dim i As Integer
 
    ' Get a list of files to import
    openwbs = Getfiles()
    Application.ScreenUpdating = False

    ' Loop through each file and copy its content into a new worksheet
    For i = 1 To Application.CountA(openwbs)
        Set Wb = Workbooks.Open(openwbs(i))
        Wb.Sheets(1).UsedRange.Copy
        Workbooks(1).Activate
        Workbooks(1).Worksheets.Add
        ActiveSheet.Range("A1").PasteSpecial
        
        ' Use the file name as the new sheet's name
        sheetName = Left(Wb.Name, InStr(Wb.Name, ".") - 1)
        ActiveSheet.Name = sheetName
        
        ' Clean up
        Application.CutCopyMode = False
        Wb.Close
    Next i

    Application.ScreenUpdating = True
End Sub

' This function opens a file picker to select multiple workbooks.
Public Function Getfiles() As Variant
    Getfiles = Application.GetOpenFilename(Title:="Select workbook(s) to import", MultiSelect:=True)
End Function


' This procedure cleans up headers and titles in each worksheet.
Public Sub CleanHeadersAndTitles()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim deleteRows As Range
    Dim headerFound As Boolean
    Dim rowIsEmpty As Boolean
    Dim headerRow As Long

    ' Turn off screen updating and automatic calculation for performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Loop through each worksheet
    For Each ws In ThisWorkbook.Sheets
        headerFound = False
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        
        ' PART 1: Identify and clean headers
        For i = 1 To lastRow
            rowIsEmpty = Application.WorksheetFunction.CountA(ws.Rows(i)) = 0

            If rowIsEmpty Then
                If deleteRows Is Nothing Then
                    Set deleteRows = ws.Rows(i)
                Else
                    Set deleteRows = Union(deleteRows, ws.Rows(i))
                End If
            ElseIf Not headerFound Then
                ' Locate headers like "Dept Nbr", "Store Nbr", or "Store Description"
                If InStr(1, ws.Cells(i, 1).Value, "Dept Nbr", vbTextCompare) > 0 Or _
                   InStr(1, ws.Cells(i, 2).Value, "Store Nbr", vbTextCompare) > 0 Or _
                   InStr(1, ws.Cells(i, 3).Value, "Store Description", vbTextCompare) > 0 Then
                    headerFound = True
                    headerRow = i
                End If
            ElseIf InStr(1, ws.Cells(i, 1).Value, "Dept Nbr", vbTextCompare) > 0 Then
                If deleteRows Is Nothing Then
                    Set deleteRows = ws.Rows(i)
                Else
                    Set deleteRows = Union(deleteRows, ws.Rows(i))
                End If
            End If
        Next i
        
        ' PART 2: Clean titles and extra rows
        For i = lastRow To 1 Step -1
            rowIsEmpty = Application.WorksheetFunction.CountA(ws.Rows(i)) = 0
            If rowIsEmpty Or i = headerRow Then
                If rowIsEmpty Then
                    If deleteRows Is Nothing Then
                        Set deleteRows = ws.Rows(i)
                    Else
                        Set deleteRows = Union(deleteRows, ws.Rows(i))
                    End If
                End If
            ElseIf InStr(1, ws.Cells(i, 1).Value, "Report ID", vbTextCompare) > 0 Then
                If deleteRows Is Nothing Then
                    Set deleteRows = ws.Rows(i)
                Else
                    Set deleteRows = Union(deleteRows, ws.Rows(i))
                End If
            End If
        Next i

        ' Delete all flagged rows at once
        If Not deleteRows Is Nothing Then
            deleteRows.Delete
            Set deleteRows = Nothing
        End If
    Next ws

    ' Restore screen updating and calculation
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub

' This procedure inserts new columns "Store Name" and "Period" next to specific headers.
Public Sub InsertColumnsAndHeaders()
    Dim ws As Worksheet
    Dim storeDescCol As Range
    Dim closingInvCol As Range

    ' Loop through each worksheet
    For Each ws In ThisWorkbook.Sheets
        ' Find the "Store Description" column
        Set storeDescCol = ws.Rows(1).Find("Store Description", LookIn:=xlValues, LookAt:=xlWhole)
        If Not storeDescCol Is Nothing Then
            storeDescCol.Offset(0, 1).EntireColumn.Insert
            storeDescCol.Offset(0, 1).Value = "Store Name"
        End If

        ' Find the "Closing Inventory" column
        Set closingInvCol = ws.Rows(1).Find("Closing Inventory", LookIn:=xlValues, LookAt:=xlWhole)
        If Not closingInvCol Is Nothing Then
            closingInvCol.Offset(0, 1).EntireColumn.Insert
            closingInvCol.Offset(0, 1).Value = "Period"
        End If
    Next ws
End Sub


' This procedure unmerges cells and splits text into adjacent cells based on a space.
Public Sub UnmergeAndSplitCells()
    Dim ws As Worksheet
    Dim cell As Range
    Dim newCell As Range
    Dim cellText As String
    Dim firstPart As String
    Dim secondPart As String
    Dim spacePos As Long

    Application.ScreenUpdating = False

    ' Loop through each worksheet
    For Each ws In ThisWorkbook.Sheets
        ' Loop through each merged cell in the worksheet
        For Each cell In ws.UsedRange
            If cell.MergeCells Then
                cellText = cell.Value
                cell.UnMerge

                spacePos = InStr(1, cellText, " ")
                If spacePos > 0 Then
                    firstPart = Left(cellText, spacePos - 1)
                    secondPart = Mid(cellText, spacePos + 1)
                    cell.Value = firstPart
                    Set newCell = cell.Offset(0, 1)
                    newCell.Value = secondPart
                Else
                    cell.Value = cellText
                End If
            End If
        Next cell
    Next ws

    Application.ScreenUpdating = True
End Sub


' This procedure moves "Total" values to adjacent cells for better formatting.
Public Sub MoveTotalToAdjacentCell()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long

    Application.ScreenUpdating = False

    ' Loop through each worksheet
    For Each ws In ThisWorkbook.Sheets
        lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row

        ' Check for "Total" in column B and move its value to column C
        For i = 1 To lastRow
            If LCase(Trim(ws.Cells(i, 2).Value)) = "total" And Trim(ws.Cells(i, 3).Value) <> "" Then
                ws.Cells(i, 3).Value = ws.Cells(i, 2).Value & " " & ws.Cells(i, 3).Value
                ws.Cells(i, 2).ClearContents
            End If
        Next i
    Next ws

    Application.ScreenUpdating = True
End Sub


' This procedure processes the "Store Description" and ensures adjacent cells are blank when required.
Public Sub ProcessAndEnsureAdjacentCellIsBlank()
    Dim ws As Worksheet
    Dim colDesc As Long, colName As Long
    Dim lastRow As Long
    Dim i As Long, j As Long
    Dim storeDesc As String, storeName As String

    ' Loop through each worksheet
    For Each ws In ThisWorkbook.Sheets
        ' Find the "Store Description" column
        For colDesc = 1 To ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
            If Trim(LCase(ws.Cells(1, colDesc).Value)) = "store description" Then
                colName = colDesc + 1 ' Assuming "Store Name" is the next column

                ' Process each row in the "Store Description" column
                lastRow = ws.Cells(ws.Rows.Count, colDesc).End(xlUp).Row
                For i = 2 To lastRow ' Assuming data starts from row 2
                    storeDesc = ws.Cells(i, colDesc).Value

                    ' Check if the store description starts with "Total"
                    If LCase(Left(storeDesc, Len("Total"))) = "total" Then
                        storeName = Trim(Mid(storeDesc, Len("Total ") + 1))
                        ws.Cells(i, colName).Value = storeName

                        ' Populate the store name upwards
                        For j = i - 1 To 2 Step -1
                            If Trim(ws.Cells(j, colName).Value) = "" Then
                                ws.Cells(j, colName).Value = storeName
                            Else
                                Exit For
                            End If
                        Next j

                        ' Clear the store name in the current cell (leave it blank)
                        ws.Cells(i, colName).ClearContents
                    End If
                Next i

                ' After processing, clear any cells adjacent to "Total"
                For i = 2 To lastRow
                    storeDesc = ws.Cells(i, colDesc).Value
                    If LCase(Left(storeDesc, Len("Total"))) = "total" Then
                        ws.Cells(i, colName).ClearContents
                    End If
                Next i

                Exit For
            End If
        Next colDesc
    Next ws
End Sub


' This procedure updates the "Period" column based on the worksheet name.
Sub UpdatePeriod()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim period As String

    ' Loop through each worksheet
    For Each ws In ThisWorkbook.Sheets
        ' Extract the period from the worksheet name
        period = ExtractPeriodFromSheetName(ws.Name)

        ' Find the last row with data
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

        ' Populate the Period column with the extracted period
        For i = 2 To lastRow
            ws.Cells(i, ws.Columns.Count).End(xlToLeft).Offset(0, 1).Value = period
        Next i
    Next ws
End Sub

' This function extracts the period from the worksheet name.
Function ExtractPeriodFromSheetName(sheetName As String) As String
    Dim period As String
    Dim pos As Integer

    ' Find the position of "p" in the sheet name
    pos = InStr(1, sheetName, "p", vbTextCompare)

    ' Extract the period number
    If pos > 0 Then
        period = Mid(sheetName, pos, 2)
    Else
        period = "Unknown" ' Default value if no period is found
    End If

    ExtractPeriodFromSheetName = period
End Function


' This procedure clears formatting in all sheets and adjusts columns/rows.
Public Sub ClearFormatAndAdjustAllSheets()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long

    ' Loop through each worksheet
    For Each ws In ThisWorkbook.Worksheets
        ' Find the last used row and column
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

        ' Clear all formatting
        ws.Cells.ClearFormats

        ' Auto-fit all rows and columns to adjust them
        ws.Rows("1:" & lastRow).AutoFit
        ws.Columns("A:" & Split(ws.Cells(1, lastCol).Address, "$")(1)).AutoFit
    Next ws

    MsgBox "All formats cleared and rows/columns adjusted in all sheets."
End Sub


' This procedure inserts or replaces tables with minimal formatting in all sheets.
Public Sub InsertOrReplaceClearTableInAllSheets()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim lastRow As Long
    Dim lastCol As Long
    Dim tableRange As Range

    ' Loop through each worksheet
    For Each ws In ThisWorkbook.Worksheets
        ' Find the last used row and column
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

        ' Check if there is any data
        If lastRow > 1 Or lastCol > 1 Then
            ' Define the table range
            Set tableRange = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))

            ' Check and remove existing tables to avoid overlapping
            If ws.ListObjects.Count > 0 Then
                ws.ListObjects(1).Unlist
            End If

            ' Add a new table
            Set tbl = ws.ListObjects.Add(xlSrcRange, tableRange, , xlYes)
            tbl.TableStyle = "TableStyleLight1" ' Apply a minimal style

            ' Optionally, name the table based on the sheet name
            tbl.Name = ws.Name & "_Table"
        End If
    Next ws

    MsgBox "Tables with minimal formatting have been inserted or replaced in all sheets!"
End Sub


' This procedure consolidates data from all sheets into a "Master" sheet and renames it to "Consolidated Data".
Sub ConsolidateDataAndRename()
    Dim ws As Worksheet
    Dim masterWs As Worksheet
    Dim lastRow As Long, masterLastRow As Long
    Dim lastCol As Long
    Dim isFirstSheet As Boolean

    ' Reference or create the Master sheet
    On Error Resume Next
    Set masterWs = ThisWorkbook.Worksheets("Master")
    On Error GoTo 0

    If masterWs Is Nothing Then
        Set masterWs = ThisWorkbook.Worksheets.Add
        masterWs.Name = "Master"
    Else
        ' Clear the Master sheet if it already exists
        masterWs.Cells.Clear
    End If

    isFirstSheet = True

    ' Loop through each worksheet
    For Each ws In ThisWorkbook.Worksheets
        ' Skip the Master sheet
        If ws.Name <> "Master" Then
            lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

            ' Find the last row in the Master sheet
            masterLastRow = masterWs.Cells(masterWs.Rows.Count, 1).End(xlUp).Row
            If masterLastRow = 1 And IsEmpty(masterWs.Cells(1, 1)) Then
                masterLastRow = 0
            End If

            ' If it's the first sheet, copy the header row; otherwise, skip it
            If isFirstSheet Then
                ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).Copy
                masterWs.Cells(masterLastRow + 1, 1).PasteSpecial Paste:=xlPasteValues
                isFirstSheet = False
            Else
                ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, lastCol)).Copy
                masterWs.Cells(masterLastRow + 1, 1).PasteSpecial Paste:=xlPasteValues
            End If

            Application.CutCopyMode = False
        End If
    Next ws

    ' Auto-fit columns in the Master sheet
    masterWs.Columns.AutoFit

    ' Rename the Master sheet to "Consolidated Data"
    masterWs.Name = "Consolidated Data"

    MsgBox "Data consolidation completed and worksheet renamed to 'Consolidated Data'!"
End Sub
