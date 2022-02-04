Attribute VB_Name = "modWorksheet"
Option Explicit

'Meta Data=============================================================
'======================================================================

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'  Copyright © 2022 Peter D Roach. All Rights Reserved.
'
'  Permission is hereby granted, free of charge, to any person
'  obtaining a copy of this software and associated documentation
'  files (the "Software"), to deal in the Software without restriction,
'  including without limitation the rights to use, copy, modify, merge,
'  publish, distribute, sublicense, and/or sell copies of the Software,
'  and to permit persons to whom the Software is furnished to do so,
'  subject to the following conditions:
'
'  The above copyright notice and this permission notice shall be
'  included in all copies or substantial portions of the Software.
'
'  THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
'  EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES
'  OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
'  NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT
'  HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,
'  WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
'  OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
'  DEALINGS IN THE SOFTWARE.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'  Module Type: Standard
'  Module Name: modWorksheet
'  Module Description: Contains common Excel Worksheet functions.
'  Module Version: 1.0
'  Module License: MIT
'  Module Author: Peter Roach; PeterRoach.Code@gmail.com
'  Module Contents:
'  ----------------------------------------
'   Public Procedures:
'      * FindFirstRow
'      * FindFirstColumn
'      * FindLastRow
'      * FindLastColumn
'     ** GetFirstRow
'        GetFirstRowInColumn
'     ** GetFirstColumn
'        GetFirstColumnInRow
'     ** GetLastRow
'        GetLastRowInColumn
'     ** GetLastColumn
'        GetLastColumnInRow
'        GetWholeRange
'        GetUsedRange
'        GetDataRange
'    *** GetCharacterDictionary
'        CreateWorkbookFromWorksheet
'        JoinRange
'        NameWorksheet
'        JoinRangeText
'   Test Procedures:
'        TestmodWorksheet
'        TestFindFirstRow
'        TestFindFirstColumn
'        TestFindLastRow
'        TestFindLastColumn
'        TestGetFirstRow
'        TestGetFirstRowInColumn
'        TestGetFirstColumn
'        TestGetFirstColumnInRow
'        TestGetLastRow
'        TestGetLastRowInColumn
'        TestGetLastColumn
'        TestGetLastColumnInRow
'        TestGetUsedRange
'        TestGetDataRange
'    *** TestGetCharacterDictionary
'        TestCreateWorkbookFromWorksheet
'        TestJoinRange
'        TestNameWorksheet
'        TestJoinRangeText
'
'   * Using the Find method can change settings in the Find (Ctrl + f) dialog
'   * Using the Find method cuts through hidden rows and columns but not filters
'  ** Many blank rows/columns in UsedRange can cause extreme inefficiencies
' *** Requires Microsoft Scripting Runtime Library
'  ----------------------------------------

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'  Example Usage:

Private Sub Example()

    Excel.Application.ScreenUpdating = False
    
    Dim WB As Workbook
    Set WB = Excel.Application.Workbooks.Add
    
    Dim WS As Worksheet
    Set WS = WB.Worksheets(1)
    
    WS.Range("B2:D4").Formula = "=ADDRESS(ROW(),COLUMN())"
    
    Debug.Print FindFirstRow(WS)
    Debug.Print FindFirstColumn(WS)
    Debug.Print FindLastRow(WS)
    Debug.Print FindLastColumn(WS)
    Debug.Print GetFirstRow(WS)
    Debug.Print GetFirstRowInColumn(WS, 2)
    Debug.Print GetFirstColumn(WS)
    Debug.Print GetFirstColumnInRow(WS, 2)
    Debug.Print GetLastRow(WS)
    Debug.Print GetLastRowInColumn(WS, 2)
    Debug.Print GetLastColumn(WS)
    Debug.Print GetLastColumnInRow(WS, 2)
    
    Debug.Print GetWholeRange(WS).Address
    Debug.Print GetDataRange(WS).Address
    
    'Requires Scripting Runtime Library
    Debug.Print Join(GetCharacterDictionary(WS).Keys(), "")
    
    Dim NB As Workbook
    Set NB = CreateWorkbookFromWorksheet(WS)
    Dim Rng As Range
    Dim i As Long
    For i = 1 To 10
        Set Rng = JoinRange(Rng, WS.Cells(i, 1))
    Next i
    Debug.Print Rng.Address
    Excel.Application.DisplayAlerts = False
    NB.Close
    Excel.Application.DisplayAlerts = True
    
    Dim NS As Worksheet
    Set NS = WB.Worksheets.Add
    Debug.Print NameWorksheet(NS, WS.Name)
    
    Debug.Print JoinRangeText(WS.Range("B2:D4"), ",")
    
    Excel.Application.DisplayAlerts = False
    WB.Close
    Excel.Application.DisplayAlerts = True

    Excel.Application.ScreenUpdating = True
    
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


'Public Procedures=====================================================
'======================================================================

Public Function GetNonBlankCells(WS As Worksheet) As Range
    Dim CRange As Range
    Dim FRange As Range
    On Error Resume Next
    Set CRange = WS.Cells.SpecialCells(xlCellTypeConstants)
    Set FRange = WS.Cells.SpecialCells(xlCellTypeFormulas)
    On Error GoTo 0
    Set GetNonBlankCells = JoinRange(CRange, FRange)
End Function

Public Function FindFirstRow&(WS As Worksheet)
    Dim Rng As Range: Set Rng = WS.Cells.Find( _
    What:="*", _
    After:=WS.Cells(WS.Rows.Count, WS.Cells.Columns.Count), _
    LookIn:=xlFormulas, _
    LookAt:=xlPart, _
    SearchOrder:=xlByRows, _
    SearchDirection:=xlNext, _
    MatchCase:=False, _
    MatchByte:=False)
    If Not Rng Is Nothing Then FindFirstRow = Rng.Row
End Function

Public Function FindFirstColumn&(WS As Worksheet)
    Dim Rng As Range: Set Rng = WS.Cells.Find( _
    What:="*", _
    After:=WS.Cells(WS.Rows.Count, WS.Cells.Columns.Count), _
    LookIn:=xlFormulas, _
    LookAt:=xlPart, _
    SearchOrder:=xlByColumns, _
    SearchDirection:=xlNext, _
    MatchCase:=False, _
    MatchByte:=False)
    If Not Rng Is Nothing Then FindFirstColumn = Rng.Column
End Function

Public Function FindLastRow&(WS As Worksheet)
    If Not IsEmpty(WS.Cells(WS.Rows.Count, WS.Columns.Count)) Then
        FindLastRow = WS.Rows.Count
        Exit Function
    End If
    Dim Rng As Range: Set Rng = WS.Cells.Find( _
    What:="*", _
    After:=WS.Cells(WS.Rows.Count, WS.Columns.Count), _
    LookIn:=xlFormulas, _
    LookAt:=xlPart, _
    SearchOrder:=xlByRows, _
    SearchDirection:=xlPrevious, _
    MatchCase:=False, _
    MatchByte:=False)
    If Not Rng Is Nothing Then FindLastRow = Rng.Row
End Function

Public Function FindLastColumn&(WS As Worksheet)
    If Not IsEmpty(WS.Cells(WS.Rows.Count, WS.Columns.Count)) Then
        FindLastColumn = WS.Columns.Count
        Exit Function
    End If
    Dim Rng As Range: Set Rng = WS.Cells.Find( _
    What:="*", _
    After:=WS.Cells(WS.Rows.Count, WS.Columns.Count), _
    LookIn:=xlFormulas, _
    LookAt:=xlPart, _
    SearchOrder:=xlByColumns, _
    SearchDirection:=xlPrevious, _
    MatchCase:=False, _
    MatchByte:=False)
    If Not Rng Is Nothing Then FindLastColumn = Rng.Column
End Function

Public Function GetFirstRow&(WS As Worksheet, _
Optional IncludeHiddenRows As Boolean = True, _
Optional IncludeHiddenColumns As Boolean = True)
    Dim UR As Range: Set UR = WS.UsedRange
    If UR Is Nothing Then Exit Function
    Dim FR&: FR = UR.Row
    Dim FC&: FC = UR.Column
    Dim LR&: LR = FR + UR.Rows.Count - 1
    Dim LC&: LC = FC + UR.Columns.Count - 1
    Dim i&, j&
    For i = FR To LR
        For j = FC To LC
            If Not IsEmpty(WS.Cells(i, j)) Then
                If Not IncludeHiddenRows Then
                    If WS.Rows(i).EntireRow.Hidden Then
                        GoTo NextRow
                    End If
                End If
                If Not IncludeHiddenColumns Then
                    If WS.Columns(j).EntireColumn.Hidden Then
                        GoTo NextCol
                    End If
                End If
                GetFirstRow = i
                Exit Function
            End If
NextCol: Next j
NextRow: Next i
End Function

Public Function GetFirstRowInColumn&(WS As Worksheet, ColumnIndex&, _
Optional IncludeHiddenRows As Boolean = True, _
Optional IncludeHiddenColumns As Boolean = True)
    Dim UR As Range: Set UR = WS.UsedRange
    If UR Is Nothing Then Exit Function
    If ColumnIndex < UR.Column Then
        Exit Function
    End If
    If ColumnIndex > UR.Column + UR.Columns.Count - 1 Then
        Exit Function
    End If
    If Not IncludeHiddenColumns Then
        If WS.Columns(ColumnIndex).EntireColumn.Hidden Then
            Exit Function
        End If
    End If
    Dim i&
    For i = UR.Row To UR.Row + UR.Rows.Count - 1
        If Not IsEmpty(WS.Cells(i, ColumnIndex)) Then
            If Not IncludeHiddenRows Then
                If WS.Rows(i).EntireRow.Hidden Then
                    GoTo NextRow
                End If
            End If
            GetFirstRowInColumn = i
            Exit Function
        End If
NextRow:
    Next i
End Function

Public Function GetFirstColumn&(WS As Worksheet, _
Optional IncludeHiddenRows As Boolean = True, _
Optional IncludeHiddenColumns As Boolean = True)
    Dim UR As Range: Set UR = WS.UsedRange
    If UR Is Nothing Then Exit Function
    Dim FR&: FR = UR.Row
    Dim FC&: FC = UR.Column
    Dim LR&: LR = FR + UR.Rows.Count - 1
    Dim LC&: LC = FC + UR.Columns.Count - 1
    Dim i&, j&
    For i = FC To LC
        For j = FR To LR
            If Not IsEmpty(WS.Cells(j, i)) Then
                If Not IncludeHiddenRows Then
                    If WS.Rows(j).EntireRow.Hidden Then
                        GoTo NextRow
                    End If
                End If
                If Not IncludeHiddenColumns Then
                    If WS.Columns(i).EntireColumn.Hidden Then
                        GoTo NextCol
                    End If
                End If
                GetFirstColumn = i
                Exit Function
            End If
NextRow: Next j
NextCol: Next i
End Function

Public Function GetFirstColumnInRow&(WS As Worksheet, RowIndex&, _
Optional IncludeHiddenRows As Boolean = True, _
Optional IncludeHiddenColumns As Boolean = True)
    Dim UR As Range: Set UR = WS.UsedRange
    If UR Is Nothing Then Exit Function
    If RowIndex < UR.Row Then
        Exit Function
    End If
    If RowIndex > UR.Row + UR.Rows.Count - 1 Then
        Exit Function
    End If
    If Not IncludeHiddenRows Then
        If WS.Rows(RowIndex).EntireRow.Hidden Then
            Exit Function
        End If
    End If
    Dim i&
    For i = UR.Column To UR.Column + UR.Columns.Count - 1
        If Not IsEmpty(WS.Cells(RowIndex, i)) Then
            If Not IncludeHiddenColumns Then
                If WS.Columns(i).EntireColumn.Hidden Then
                    GoTo NextCol
                End If
            End If
            GetFirstColumnInRow = i
            Exit Function
        End If
NextCol:
    Next i
End Function

Public Function GetLastRow&(WS As Worksheet, _
Optional IncludeHiddenRows As Boolean = True, _
Optional IncludeHiddenColumns As Boolean = True)
    Dim UR As Range: Set UR = WS.UsedRange
    If UR Is Nothing Then Exit Function
    Dim FR&: FR = UR.Row
    Dim FC&: FC = UR.Column
    Dim LR&: LR = FR + UR.Rows.Count - 1
    Dim LC&: LC = FC + UR.Columns.Count - 1
    Dim i&, j&
    For i = LR To FR Step -1
        For j = LC To FC Step -1
            If Not IsEmpty(WS.Cells(i, j)) Then
                If Not IncludeHiddenRows Then
                    If WS.Rows(i).EntireRow.Hidden Then
                        GoTo NextRow
                    End If
                End If
                If Not IncludeHiddenColumns Then
                    If WS.Columns(j).EntireColumn.Hidden Then
                        GoTo NextCol
                    End If
                End If
                GetLastRow = i
                Exit Function
            End If
NextCol: Next j
NextRow: Next i
End Function

Public Function GetLastRowInColumn&(WS As Worksheet, ColumnIndex&, _
Optional IncludeHiddenRows As Boolean = True, _
Optional IncludeHiddenColumns As Boolean = True)
    Dim UR As Range: Set UR = WS.UsedRange
    If UR Is Nothing Then Exit Function
    If ColumnIndex < UR.Column Then
        Exit Function
    End If
    If ColumnIndex > UR.Column + UR.Columns.Count - 1 Then
        Exit Function
    End If
    If Not IncludeHiddenColumns Then
        If WS.Columns(ColumnIndex).EntireColumn.Hidden Then
            Exit Function
        End If
    End If
    Dim i&
    For i = UR.Row + UR.Rows.Count To 1 Step -1
        If Not IsEmpty(WS.Cells(i, ColumnIndex)) Then
            If Not IncludeHiddenRows Then
                If WS.Rows(i).EntireRow.Hidden Then
                    GoTo NextRow
                End If
            End If
            GetLastRowInColumn = i
            Exit Function
        End If
NextRow:
    Next i
End Function

Public Function GetLastColumn&(WS As Worksheet, _
Optional IncludeHiddenRows As Boolean = True, _
Optional IncludeHiddenColumns As Boolean = True)
    Dim UR As Range: Set UR = WS.UsedRange
    If UR Is Nothing Then Exit Function
    Dim FR&: FR = UR.Row
    Dim FC&: FC = UR.Column
    Dim LR&: LR = FR + UR.Rows.Count - 1
    Dim LC&: LC = FC + UR.Columns.Count - 1
    Dim i&, j&
    For i = LC To FC Step -1
        For j = LR To FR Step -1
            If Not IsEmpty(WS.Cells(j, i)) Then
                If Not IncludeHiddenRows Then
                    If WS.Rows(j).EntireRow.Hidden Then
                        GoTo NextRow
                    End If
                End If
                If Not IncludeHiddenColumns Then
                    If WS.Columns(i).EntireColumn.Hidden Then
                        GoTo NextCol
                    End If
                End If
                GetLastColumn = i
                Exit Function
            End If
NextRow: Next j
NextCol: Next i
End Function

Public Function GetLastColumnInRow&(WS As Worksheet, RowIndex&, _
Optional IncludeHiddenRows As Boolean = True, _
Optional IncludeHiddenColumns As Boolean = True)
    Dim UR As Range: Set UR = WS.UsedRange
    If UR Is Nothing Then Exit Function
    If RowIndex < UR.Row Then
        Exit Function
    End If
    If RowIndex > UR.Row + UR.Rows.Count - 1 Then
        Exit Function
    End If
    If Not IncludeHiddenRows Then
        If WS.Rows(RowIndex).EntireRow.Hidden Then
            Exit Function
        End If
    End If
    Dim i&
    For i = UR.Column + UR.Columns.Count - 1 To 1 Step -1
        If Not IsEmpty(WS.Cells(RowIndex, i)) Then
            If Not IncludeHiddenColumns Then
                If WS.Columns(i).EntireColumn.Hidden Then
                    GoTo NextCol
                End If
            End If
            GetLastColumnInRow = i
            Exit Function
        End If
NextCol:
    Next i
End Function

Public Function GetWholeRange(WS As Worksheet) As Range
    Dim LR&
    Dim LC&
    LR = GetLastRow(WS)
    LC = GetLastColumn(WS)
    If LR > 0 And LC > 0 Then
        Set GetWholeRange = _
        Range(WS.Cells(1, 1), WS.Cells(LR, LC))
    End If
End Function

Public Function GetDataRange(WS As Worksheet) As Range
    Dim FR&
    Dim FC&
    Dim LR&
    Dim LC&
    FR = GetFirstRow(WS)
    FC = GetFirstColumn(WS)
    LR = GetLastRow(WS)
    LC = GetLastColumn(WS)
    If FR > 0 And FC > 0 And LR > 0 And LC > 0 Then
        Set GetDataRange = _
        Range(WS.Cells(FR, FC), WS.Cells(LR, LC))
    End If
End Function

Public Function GetCharacterDictionary(WS As Worksheet) As Object
    Dim Dict As Object
    Set Dict = CreateObject("Scripting.Dictionary")
    Dim UR As Range
    Set UR = WS.UsedRange
    If UR Is Nothing Then
        Set GetCharacterDictionary = Dict
        Exit Function
    End If
    Dim Rng As Range
    For Each Rng In UR
        If Not IsEmpty(Rng) Then
            Dim RngValue$
            RngValue = Rng.Value
            Dim i&
            For i = 1 To Len(RngValue)
                Dim Char As String * 1
                Char = Mid$(RngValue, i, 1)
                If Dict.Exists(Char) Then
                    Dict(Char) = _
                    Dict(Char) + 1
                Else
                    Dict(Char) = 1
                End If
            Next i
        End If
    Next Rng
    Set GetCharacterDictionary = Dict
End Function

Public Function CreateWorkbookFromWorksheet(WS As Worksheet) As Workbook
    WS.Copy
    Set CreateWorkbookFromWorksheet = Application.ActiveWorkbook
End Function

Public Function JoinRange(Rng As Range, JoinRng As Range) As Range
    If Rng Is Nothing Then
        Set JoinRange = JoinRng
    ElseIf JoinRng Is Nothing Then
        Set JoinRange = Rng
    Else
        Set JoinRange = Union(Rng, JoinRng)
    End If
End Function

Public Function NameWorksheet$(WS As Worksheet, WSName$)
    NameWorksheet = WSName
    If NameWorksheet = WS.Name Then
        Exit Function
    End If
    Dim WSColl As Object
    Set WSColl = WS.Parent.Worksheets
    Dim WSCount&
    WSCount = WSColl.Count
    Dim Arr$()
    ReDim Arr(1 To WSCount)
    Dim i&
    For i = 1 To WSCount
        Arr(i) = WSColl(i).Name
    Next i
    Dim Named As Boolean
    Do While Not Named
        For i = 1 To WSCount
            If NameWorksheet = Arr(i) Then
                If NameWorksheet = WS.Name Then
                    Exit Function
                End If
                Dim c&
                c = c + 1
                NameWorksheet = WSName & " (" & c & ")"
                GoTo NextIteration
            End If
        Next i
        WS.Name = NameWorksheet
        Named = True
NextIteration:
    Loop
End Function

Public Function JoinRangeText$(Rng As Range, Delimiter$, _
Optional IgnoreBlanks As Boolean = False, _
Optional Func$ = "", _
Optional Direction$ = "Row")
    If Rng.Cells.CountLarge = 1 Then
        If IgnoreBlanks Then
            If IsEmpty(Rng) Then
                Exit Function
            End If
        End If
        If Func <> "" Then
            JoinRangeText = _
            Excel.Application.Run(Func, Rng.Value)
        Else
            JoinRangeText = Rng.Value
        End If
    End If
    Dim Arr(): Arr = Rng.Value
    Dim RC&: RC = Rng.Rows.Count
    Dim CC&: CC = Rng.Columns.Count
    Dim SArr$(): ReDim SArr(1 To Rng.Cells.CountLarge)
    Dim i&
    Dim j&
    Dim k&
    Dim c&
    Select Case Direction
        Case "Row"
            If Func <> "" Then
                If IgnoreBlanks Then
                    For i = 1 To RC
                        For j = 1 To CC
                            If Arr(i, j) <> Empty Then
                                k = k + 1
                                SArr(k) = _
                                Excel.Application.Run(Func, Arr(i, j))
                            Else
                                c = c + 1
                            End If
                        Next j
                    Next i
                Else
                    For i = 1 To RC
                        For j = 1 To CC
                            k = k + 1
                            SArr(k) = _
                            Excel.Application.Run(Func, Arr(i, j))
                        Next j
                    Next i
                End If
            Else
                If IgnoreBlanks Then
                    For i = 1 To RC
                        For j = 1 To CC
                            If Arr(i, j) <> Empty Then
                                k = k + 1
                                SArr(k) = Arr(i, j)
                            Else
                                c = c + 1
                            End If
                        Next j
                    Next i
                Else
                    For i = 1 To RC
                        For j = 1 To CC
                            k = k + 1
                            SArr(k) = Arr(i, j)
                        Next j
                    Next i
                End If
            End If
        Case "Column"
            If Func <> "" Then
                If IgnoreBlanks Then
                    For i = 1 To CC
                        For j = 1 To RC
                            If Arr(j, i) <> Empty Then
                                k = k + 1
                                SArr(k) = _
                                Excel.Application.Run(Func, Arr(j, i))
                            Else
                                c = c + 1
                            End If
                        Next j
                    Next i
                Else
                    For i = 1 To CC
                        For j = 1 To RC
                            k = k + 1
                            SArr(k) = _
                            Excel.Application.Run(Func, Arr(j, i))
                        Next j
                    Next i
                End If
            Else
                If IgnoreBlanks Then
                    For i = 1 To CC
                        For j = 1 To RC
                            If Arr(j, i) <> Empty Then
                                k = k + 1
                                SArr(k) = Arr(j, i)
                            Else
                                c = c + 1
                            End If
                        Next j
                    Next i
                Else
                    For i = 1 To CC
                        For j = 1 To RC
                            k = k + 1
                            SArr(k) = Arr(j, i)
                        Next j
                    Next i
                End If
            End If
        Case Else
            Err.Raise 5
    End Select
    If IgnoreBlanks Then
        ReDim Preserve SArr(1 To UBound(SArr) - c)
    End If
    JoinRangeText = Join(SArr, Delimiter)
End Function


'Unit Tests============================================================
'======================================================================

Private Function TestmodWorksheet() As Boolean

    TestmodWorksheet = _
        TestGetNonBlankCells And _
        TestFindFirstRow And _
        TestFindFirstColumn And _
        TestFindLastRow And _
        TestFindLastColumn And _
        TestGetFirstRow And _
        TestGetFirstRowInColumn And _
        TestGetFirstColumn And _
        TestGetFirstColumnInRow And _
        TestGetLastRow And _
        TestGetLastRowInColumn And _
        TestGetLastColumn And _
        TestGetLastColumnInRow And _
        TestGetWholeRange And _
        TestGetDataRange And _
        TestGetCharacterDictionary And _
        TestCreateWorkbookFromWorksheet And _
        TestJoinRange And _
        TestNameWorksheet And _
        TestJoinRangeText

    Debug.Print "TestmodWorksheet: " & TestmodWorksheet

End Function

Private Function TestGetNonBlankCells() As Boolean
    
    TestGetNonBlankCells = True
    
    Excel.Application.ScreenUpdating = False
    
    Dim WB As Workbook
    Set WB = Excel.Application.Workbooks.Add
    
    Dim WS As Worksheet
    Set WS = WB.Worksheets(1)
    
    'Empty
    If Not GetNonBlankCells(WS) Is Nothing Then
        TestGetNonBlankCells = False
        Debug.Print "Empty"
    End If
    
    'Single cell
        'Value
        WS.Cells.Clear
        WS.Cells(1, 1).Value = "Test"
        If GetNonBlankCells(WS).Address <> "$A$1" Then
            TestGetNonBlankCells = False
            Debug.Print "Single cell Value"
        End If
        'Formula
        WS.Cells.Clear
        WS.Cells(1, 1).Formula = "=ADDRESS(ROW(),COLUMN())"
        If GetNonBlankCells(WS).Address <> "$A$1" Then
            TestGetNonBlankCells = False
            Debug.Print "Single cell Formula"
        End If
    
    'Multiple contiguous cells
        'Value
        WS.Cells.Clear
        WS.Cells(1, 1).Value = "Test"
        WS.Cells(1, 2).Value = "Test"
        If GetNonBlankCells(WS).Address <> "$A$1:$B$1" Then
            TestGetNonBlankCells = False
            Debug.Print "Multiple contiguous cells Value"
        End If
        'Formula
        WS.Cells.Clear
        WS.Cells(1, 1).Formula = "=ADDRESS(ROW(),COLUMN())"
        WS.Cells(1, 2).Formula = "=ADDRESS(ROW(),COLUMN())"
        If GetNonBlankCells(WS).Address <> "$A$1:$B$1" Then
            TestGetNonBlankCells = False
            Debug.Print "Multiple contiguous cells Formula"
        End If
        'Both
        WS.Cells.Clear
        WS.Cells(1, 1).Value = "Test"
        WS.Cells(1, 2).Formula = "=ADDRESS(ROW(),COLUMN())"
        If GetNonBlankCells(WS).Address <> "$A$1:$B$1" Then
            TestGetNonBlankCells = False
            Debug.Print "Multiple contiguous cells Both"
        End If
    
    'Multiple noncontiguous cells
        'Value
        WS.Cells.Clear
        WS.Cells(1, 1).Value = "Test"
        WS.Cells(3, 3).Value = "Test"
        If GetNonBlankCells(WS).Address <> "$A$1,$C$3" Then
            TestGetNonBlankCells = False
            Debug.Print "Multiple contiguous cells Value"
        End If
        'Formula
        WS.Cells.Clear
        WS.Cells(1, 1).Formula = "=ADDRESS(ROW(),COLUMN())"
        WS.Cells(3, 3).Formula = "=ADDRESS(ROW(),COLUMN())"
        If GetNonBlankCells(WS).Address <> "$A$1,$C$3" Then
            TestGetNonBlankCells = False
            Debug.Print "Multiple contiguous cells Formula"
        End If
        'Both
        WS.Cells.Clear
        WS.Cells(1, 1).Value = "Test"
        WS.Cells(3, 3).Formula = "=ADDRESS(ROW(),COLUMN())"
        If GetNonBlankCells(WS).Address <> "$A$1,$C$3" Then
            TestGetNonBlankCells = False
            Debug.Print "Multiple contiguous cells Both"
        End If
    
    'Error
    WS.Cells.Clear
    WS.Cells(1, 1).Formula = "=1/0"
    If GetNonBlankCells(WS).Address <> "$A$1" Then
        TestGetNonBlankCells = False
        Debug.Print "Error"
    End If
    
    Excel.Application.DisplayAlerts = False
    WB.Close
    Excel.Application.DisplayAlerts = True
    
    Excel.Application.ScreenUpdating = True
    
    Debug.Print "TestGetNonBlankCells: " & TestGetNonBlankCells
    
End Function

Private Function TestFindFirstRow() As Boolean
    
    TestFindFirstRow = True
    
    Excel.Application.ScreenUpdating = False
    
    Dim WB As Workbook
    Set WB = Excel.Application.Workbooks.Add
    
    Dim WS As Worksheet
    Set WS = WB.Worksheets(1)
    
    'Blank
        'Unhidden
            If FindFirstRow(WS) <> 0 Then
                TestFindFirstRow = False
                Debug.Print "Blank Unhidden"
            End If
        'Hidden Row
        WS.Rows(1).Hidden = True
            If FindFirstRow(WS) <> 0 Then
                TestFindFirstRow = False
                Debug.Print "Blank Hidden Row"
            End If
        WS.Rows.Hidden = False
        'Hidden Column
        WS.Columns(1).Hidden = True
            If FindFirstRow(WS) <> 0 Then
                TestFindFirstRow = False
                Debug.Print "Blank Hidden Column"
            End If
        WS.Columns.Hidden = False
        'Hidden Row And Column
        WS.Columns(1).Hidden = True
        WS.Rows(1).Hidden = True
            If FindFirstRow(WS) <> 0 Then
                TestFindFirstRow = False
                Debug.Print "Blank Hidden Row And Column"
            End If
        WS.Rows.Hidden = False
        WS.Columns.Hidden = False
        'Filter
            'Cannot filter blank
            
    'Single cell
        'Start 1, 1
        WS.Cells(1, 1).Value = "Test"
            'Unhidden
                If FindFirstRow(WS) <> 1 Then
                    TestFindFirstRow = False
                    Debug.Print "Single cell Start 1, 1 Unhidden"
                End If
            'Hidden Row
            WS.Rows(1).Hidden = True
                If FindFirstRow(WS) <> 1 Then
                    TestFindFirstRow = False
                    Debug.Print "Single cell Start 1, 1 Hidden Row"
                End If
            WS.Rows.Hidden = False
            'Hidden Column
            WS.Columns(1).Hidden = True
                If FindFirstRow(WS) <> 1 Then
                    TestFindFirstRow = False
                    Debug.Print "Single cell Start 1, 1 Hidden Column"
                End If
            WS.Columns.Hidden = False
            'Hidden Row And Column
            WS.Columns(1).Hidden = True
            WS.Rows(1).Hidden = True
                If FindFirstRow(WS) <> 1 Then
                    TestFindFirstRow = False
                    Debug.Print "Single cell Start 1, 1 Hidden Row And Column"
                End If
            WS.Rows.Hidden = False
            WS.Columns.Hidden = False
            'Filter
                'Cannot filter single cell
        'Start 2, 2
        WS.Cells.Clear
        WS.Cells(2, 2).Value = "Test"
            'Unhidden
                If FindFirstRow(WS) <> 2 Then
                    TestFindFirstRow = False
                    Debug.Print "Single cell Start 2, 2 Unhidden"
                End If
            'Hidden Row
            WS.Rows(2).Hidden = True
                If FindFirstRow(WS) <> 2 Then
                    TestFindFirstRow = False
                    Debug.Print "Single cell Start 2, 2 Hidden Row"
                End If
            WS.Rows.Hidden = False
            'Hidden Column
            WS.Columns(2).Hidden = True
                If FindFirstRow(WS) <> 2 Then
                    TestFindFirstRow = False
                    Debug.Print "Single cell Start 2, 2 Hidden Column"
                End If
            WS.Columns.Hidden = False
            'Hidden Row And Column
            WS.Columns(2).Hidden = True
            WS.Rows(2).Hidden = True
                If FindFirstRow(WS) <> 2 Then
                    TestFindFirstRow = False
                    Debug.Print "Single cell Start 2, 2 Hidden Row And Column"
                End If
            WS.Rows.Hidden = False
            WS.Columns.Hidden = False
            'Filter
                'Cannot filter single cell
                
    'Muliple cell
        'Start 1, 1
        WS.Cells.Clear
        WS.Range("A1:C1").Value = "Header"
        WS.Range("A2:C4").Formula = "=ADDRESS(ROW(),COLUMN())"
            'Unhidden
                If FindFirstRow(WS) <> 1 Then
                    TestFindFirstRow = False
                    Debug.Print "Muliple cell Start 1, 1 Unhidden"
                End If
            'Hidden Row
            WS.Rows(1).Hidden = True
                If FindFirstRow(WS) <> 1 Then
                    TestFindFirstRow = False
                    Debug.Print "Muliple cell Start 1, 1 Hidden Row"
                End If
            WS.Rows.Hidden = False
            'Hidden Column
            WS.Columns(1).Hidden = True
                If FindFirstRow(WS) <> 1 Then
                    TestFindFirstRow = False
                    Debug.Print "Muliple cell Start 1, 1 Hidden Column"
                End If
            WS.Columns.Hidden = False
            'Hidden Row And Column
            WS.Columns(1).Hidden = True
            WS.Rows(1).Hidden = True
                If FindFirstRow(WS) <> 1 Then
                    TestFindFirstRow = False
                    Debug.Print "Muliple cell Start 1, 1 Hidden Row And Column"
                End If
            WS.Rows.Hidden = False
            WS.Columns.Hidden = False
            'Filter
            WS.Range("A1:C1").Clear
            WS.Range("$A$1:$C$4").AutoFilter Field:=1, Criteria1:="$A$3"
                If FindFirstRow(WS) <> 3 Then
                    TestFindFirstRow = False
                    Debug.Print "Muliple cell Start 1, 1 Filter"
                End If
            WS.AutoFilterMode = False
        'Start 2, 2
        WS.Cells.Clear
        WS.Range("B2:D2").Value = "Header"
        WS.Range("B3:D5").Formula = "=ADDRESS(ROW(),COLUMN())"
            'Unhidden
                If FindFirstRow(WS) <> 2 Then
                    TestFindFirstRow = False
                    Debug.Print "Muliple cell Start 2, 2 Unhidden"
                End If
            'Hidden Row
            WS.Rows(2).Hidden = True
                If FindFirstRow(WS) <> 2 Then
                    TestFindFirstRow = False
                    Debug.Print "Muliple cell Start 2, 2 Hidden Row"
                End If
            WS.Rows.Hidden = False
            'Hidden Column
            WS.Columns(2).Hidden = True
                If FindFirstRow(WS) <> 2 Then
                    TestFindFirstRow = False
                    Debug.Print "Muliple cell Start 2, 2 Hidden Column"
                End If
            WS.Columns.Hidden = False
            'Hidden Row And Column
            WS.Columns(2).Hidden = True
            WS.Rows(2).Hidden = True
                If FindFirstRow(WS) <> 2 Then
                    TestFindFirstRow = False
                    Debug.Print "Muliple cell Start 2, 2 Hidden Row And Column"
                End If
            WS.Rows.Hidden = False
            WS.Columns.Hidden = False
            'Filter
            WS.Range("$B$2:$D$2").Clear
            WS.Range("$B$2:$D$5").AutoFilter Field:=1, Criteria1:="$B$4"
                If FindFirstRow(WS) <> 4 Then
                    TestFindFirstRow = False
                    Debug.Print "Muliple cell Start 2, 2 Filter"
                End If
            WS.AutoFilterMode = False
            
    'Dynamic
    WS.Cells.Clear
    Dim i&
    For i = 5 To 1 Step -1
        WS.Cells(i, i).Value = i
        If FindFirstRow(WS) <> i Then
            TestFindFirstRow = False
            Debug.Print "Dynamic"
        End If
    Next i
    
    Excel.Application.DisplayAlerts = False
    WB.Close
    Excel.Application.DisplayAlerts = True
    
    Excel.Application.ScreenUpdating = True
    
    Debug.Print "TestFindFirstRow: " & TestFindFirstRow
    
End Function

Private Function TestFindFirstColumn() As Boolean

    TestFindFirstColumn = True
    
    Excel.Application.ScreenUpdating = False
    
    Dim WB As Workbook
    Set WB = Excel.Application.Workbooks.Add
    
    Dim WS As Worksheet
    Set WS = WB.Worksheets(1)
    
    'Blank
        'Unhidden
            If FindFirstColumn(WS) <> 0 Then
                TestFindFirstColumn = False
                Debug.Print "Blank Unhidden"
            End If
        'Hidden Row
        WS.Rows(1).Hidden = True
            If FindFirstColumn(WS) <> 0 Then
                TestFindFirstColumn = False
                Debug.Print "Blank Hidden Row"
            End If
        WS.Rows.Hidden = False
        'Hidden Column
        WS.Columns(1).Hidden = True
            If FindFirstColumn(WS) <> 0 Then
                TestFindFirstColumn = False
                Debug.Print "Blank Hidden Column"
            End If
        WS.Columns.Hidden = False
        'Hidden Row And Column
        WS.Columns(1).Hidden = True
        WS.Rows(1).Hidden = True
            If FindFirstColumn(WS) <> 0 Then
                TestFindFirstColumn = False
                Debug.Print "Blank Hidden Row And Column"
            End If
        WS.Rows.Hidden = False
        WS.Columns.Hidden = False
        'Filter
            'Cannot filter blank
            
    'Single cell
        'Start 1, 1
        WS.Cells(1, 1).Value = "Test"
            'Unhidden
                If FindFirstColumn(WS) <> 1 Then
                    TestFindFirstColumn = False
                    Debug.Print "Single cell Start 1, 1 Unhidden"
                End If
            'Hidden Row
            WS.Rows(1).Hidden = True
                If FindFirstColumn(WS) <> 1 Then
                    TestFindFirstColumn = False
                    Debug.Print "Single cell Start 1, 1 Hidden Row"
                End If
            WS.Rows.Hidden = False
            'Hidden Column
            WS.Columns(1).Hidden = True
                If FindFirstColumn(WS) <> 1 Then
                    TestFindFirstColumn = False
                    Debug.Print "Single cell Start 1, 1 Hidden Column"
                End If
            WS.Columns.Hidden = False
            'Hidden Row And Column
            WS.Columns(1).Hidden = True
            WS.Rows(1).Hidden = True
                If FindFirstColumn(WS) <> 1 Then
                    TestFindFirstColumn = False
                    Debug.Print "Single cell Start 1, 1 Hidden Row And Column"
                End If
            WS.Rows.Hidden = False
            WS.Columns.Hidden = False
            'Filter
                'Cannot filter single cell
        'Start 2, 2
        WS.Cells.Clear
        WS.Cells(2, 2).Value = "Test"
            'Unhidden
                If FindFirstColumn(WS) <> 2 Then
                    TestFindFirstColumn = False
                    Debug.Print "Single cell Start 2, 2 Unhidden"
                End If
            'Hidden Row
            WS.Rows(2).Hidden = True
                If FindFirstColumn(WS) <> 2 Then
                    TestFindFirstColumn = False
                    Debug.Print "Single cell Start 2, 2 Hidden Row"
                End If
            WS.Rows.Hidden = False
            'Hidden Column
            WS.Columns(2).Hidden = True
                If FindFirstColumn(WS) <> 2 Then
                    TestFindFirstColumn = False
                    Debug.Print "Single cell Start 2, 2 Hidden Column"
                End If
            WS.Columns.Hidden = False
            'Hidden Row And Column
            WS.Columns(2).Hidden = True
            WS.Rows(2).Hidden = True
                If FindFirstColumn(WS) <> 2 Then
                    TestFindFirstColumn = False
                    Debug.Print "Single cell Start 2, 2 Hidden Row And Column"
                End If
            WS.Rows.Hidden = False
            WS.Columns.Hidden = False
            'Filter
                'Cannot filter single cell
                
    'Muliple cell
        'Start 1, 1
        WS.Cells.Clear
        WS.Range("A1:C1").Value = "Header"
        WS.Range("A2:C4").Formula = "=ADDRESS(ROW(),COLUMN())"
            'Unhidden
                If FindFirstColumn(WS) <> 1 Then
                    TestFindFirstColumn = False
                    Debug.Print "Muliple cell Start 1, 1 Unhidden"
                End If
            'Hidden Row
            WS.Rows(1).Hidden = True
                If FindFirstColumn(WS) <> 1 Then
                    TestFindFirstColumn = False
                    Debug.Print "Muliple cell Start 1, 1 Hidden Row"
                End If
            WS.Rows.Hidden = False
            'Hidden Column
            WS.Columns(1).Hidden = True
                If FindFirstColumn(WS) <> 1 Then
                    TestFindFirstColumn = False
                    Debug.Print "Muliple cell Start 1, 1 Hidden Column"
                End If
            WS.Columns.Hidden = False
            'Hidden Row And Column
            WS.Columns(1).Hidden = True
            WS.Rows(1).Hidden = True
                If FindFirstColumn(WS) <> 1 Then
                    TestFindFirstColumn = False
                    Debug.Print "Muliple cell Start 1, 1 Hidden Row And Column"
                End If
            WS.Rows.Hidden = False
            WS.Columns.Hidden = False
            'Filter
            WS.Range("A1:C1").Clear
            WS.Range("$A$1:$C$4").AutoFilter Field:=1, Criteria1:="$A$3"
                If FindFirstColumn(WS) <> 1 Then
                    TestFindFirstColumn = False
                    Debug.Print "Muliple cell Start 1, 1 Filter"
                End If
            WS.AutoFilterMode = False
        'Start 2, 2
        WS.Cells.Clear
        WS.Range("B2:D2").Value = "Header"
        WS.Range("B3:D5").Formula = "=ADDRESS(ROW(),COLUMN())"
            'Unhidden
                If FindFirstColumn(WS) <> 2 Then
                    TestFindFirstColumn = False
                    Debug.Print "Muliple cell Start 2, 2 Unhidden"
                End If
            'Hidden Row
            WS.Rows(2).Hidden = True
                If FindFirstColumn(WS) <> 2 Then
                    TestFindFirstColumn = False
                    Debug.Print "Muliple cell Start 2, 2 Hidden Row"
                End If
            WS.Rows.Hidden = False
            'Hidden Column
            WS.Columns(2).Hidden = True
                If FindFirstColumn(WS) <> 2 Then
                    TestFindFirstColumn = False
                    Debug.Print "Muliple cell Start 2, 2 Hidden Column"
                End If
            WS.Columns.Hidden = False
            'Hidden Row And Column
            WS.Columns(2).Hidden = True
            WS.Rows(2).Hidden = True
                If FindFirstColumn(WS) <> 2 Then
                    TestFindFirstColumn = False
                    Debug.Print "Muliple cell Start 2, 2 Hidden Row And Column"
                End If
            WS.Rows.Hidden = False
            WS.Columns.Hidden = False
            'Filter
            WS.Range("$B$2:$D$2").Clear
            WS.Range("$B$2:$D$5").AutoFilter Field:=1, Criteria1:="$B$4"
                If FindFirstColumn(WS) <> 2 Then
                    TestFindFirstColumn = False
                    Debug.Print "Muliple cell Start 2, 2 Filter"
                End If
            WS.AutoFilterMode = False
            
    'Dynamic
    WS.Cells.Clear
    Dim i&
    For i = 5 To 1 Step -1
        WS.Cells(i, i).Value = i
        If FindFirstColumn(WS) <> i Then
            TestFindFirstColumn = False
            Debug.Print "Dynamic"
        End If
    Next i
    
    Excel.Application.DisplayAlerts = False
    WB.Close
    Excel.Application.DisplayAlerts = True
    
    Excel.Application.ScreenUpdating = True
    
    Debug.Print "TestFindFirstColumn: " & TestFindFirstColumn
    
End Function

Private Function TestFindLastRow() As Boolean
    
    TestFindLastRow = True
    
    Excel.Application.ScreenUpdating = False
    
    Dim WB As Workbook
    Set WB = Excel.Application.Workbooks.Add
    
    Dim WS As Worksheet
    Set WS = WB.Worksheets(1)
    
    'Blank
        'Unhidden
            If FindLastRow(WS) <> 0 Then
                TestFindLastRow = False
                Debug.Print "Blank Unhidden"
            End If
        'Hidden Row
        WS.Rows(1).Hidden = True
            If FindLastRow(WS) <> 0 Then
                TestFindLastRow = False
                Debug.Print "Blank Hidden Row"
            End If
        WS.Rows.Hidden = False
        'Hidden Column
        WS.Columns(1).Hidden = True
            If FindLastRow(WS) <> 0 Then
                TestFindLastRow = False
                Debug.Print "Blank Hidden Column"
            End If
        WS.Columns.Hidden = False
        'Hidden Row And Column
        WS.Columns(1).Hidden = True
        WS.Rows(1).Hidden = True
            If FindLastRow(WS) <> 0 Then
                TestFindLastRow = False
                Debug.Print "Blank Hidden Row And Column"
            End If
        WS.Rows.Hidden = False
        WS.Columns.Hidden = False
        'Filter
            'Cannot filter blank
            
    'Single cell
        'Start 1, 1
        WS.Cells(1, 1).Value = "Test"
            'Unhidden
                If FindLastRow(WS) <> 1 Then
                    TestFindLastRow = False
                    Debug.Print "Single cell Start 1, 1 Unhidden"
                End If
            'Hidden Row
            WS.Rows(1).Hidden = True
                If FindLastRow(WS) <> 1 Then
                    TestFindLastRow = False
                    Debug.Print "Single cell Start 1, 1 Hidden Row"
                End If
            WS.Rows.Hidden = False
            'Hidden Column
            WS.Columns(1).Hidden = True
                If FindLastRow(WS) <> 1 Then
                    TestFindLastRow = False
                    Debug.Print "Single cell Start 1, 1 Hidden Column"
                End If
            WS.Columns.Hidden = False
            'Hidden Row And Column
            WS.Columns(1).Hidden = True
            WS.Rows(1).Hidden = True
                If FindLastRow(WS) <> 1 Then
                    TestFindLastRow = False
                    Debug.Print "Single cell Start 1, 1 Hidden Row And Column"
                End If
            WS.Rows.Hidden = False
            WS.Columns.Hidden = False
            'Filter
                'Cannot filter single cell
        'Start 2, 2
        WS.Cells.Clear
        WS.Cells(2, 2).Value = "Test"
            'Unhidden
                If FindLastRow(WS) <> 2 Then
                    TestFindLastRow = False
                    Debug.Print "Single cell Start 2, 2 Unhidden"
                End If
            'Hidden Row
            WS.Rows(2).Hidden = True
                If FindLastRow(WS) <> 2 Then
                    TestFindLastRow = False
                    Debug.Print "Single cell Start 2, 2 Hidden Row"
                End If
            WS.Rows.Hidden = False
            'Hidden Column
            WS.Columns(2).Hidden = True
                If FindLastRow(WS) <> 2 Then
                    TestFindLastRow = False
                    Debug.Print "Single cell Start 2, 2 Hidden Column"
                End If
            WS.Columns.Hidden = False
            'Hidden Row And Column
            WS.Columns(2).Hidden = True
            WS.Rows(2).Hidden = True
                If FindLastRow(WS) <> 2 Then
                    TestFindLastRow = False
                    Debug.Print "Single cell Start 2, 2 Hidden Row And Column"
                End If
            WS.Rows.Hidden = False
            WS.Columns.Hidden = False
            'Filter
                'Cannot filter single cell
                
    'Muliple cell
        'Start 1, 1
        WS.Cells.Clear
        WS.Range("A1:C1").Value = "Header"
        WS.Range("A2:C4").Formula = "=ADDRESS(ROW(),COLUMN())"
            'Unhidden
                If FindLastRow(WS) <> 4 Then
                    TestFindLastRow = False
                    Debug.Print "Muliple cell Start 1, 1 Unhidden"
                End If
            'Hidden Row
            WS.Rows(1).Hidden = True
                If FindLastRow(WS) <> 4 Then
                    TestFindLastRow = False
                    Debug.Print "Muliple cell Start 1, 1 Hidden Row"
                End If
            WS.Rows.Hidden = False
            'Hidden Column
            WS.Columns(1).Hidden = True
                If FindLastRow(WS) <> 4 Then
                    TestFindLastRow = False
                    Debug.Print "Muliple cell Start 1, 1 Hidden Column"
                End If
            WS.Columns.Hidden = False
            'Hidden Row And Column
            WS.Columns(1).Hidden = True
            WS.Rows(1).Hidden = True
                If FindLastRow(WS) <> 4 Then
                    TestFindLastRow = False
                    Debug.Print "Muliple cell Start 1, 1 Hidden Row And Column"
                End If
            WS.Rows.Hidden = False
            WS.Columns.Hidden = False
            'Filter
            WS.Range("A1:C1").Clear
            WS.Range("$A$1:$C$4").AutoFilter Field:=1, Criteria1:="$A$3"
                If FindLastRow(WS) <> 3 Then
                    TestFindLastRow = False
                    Debug.Print "Muliple cell Start 1, 1 Filter"
                End If
            WS.AutoFilterMode = False
        'Start 2, 2
        WS.Cells.Clear
        WS.Range("B2:D2").Value = "Header"
        WS.Range("B3:D5").Formula = "=ADDRESS(ROW(),COLUMN())"
            'Unhidden
                If FindLastRow(WS) <> 5 Then
                    TestFindLastRow = False
                    Debug.Print "Muliple cell Start 2, 2 Unhidden"
                End If
            'Hidden Row
            WS.Rows(2).Hidden = True
                If FindLastRow(WS) <> 5 Then
                    TestFindLastRow = False
                    Debug.Print "Muliple cell Start 2, 2 Hidden Row"
                End If
            WS.Rows.Hidden = False
            'Hidden Column
            WS.Columns(2).Hidden = True
                If FindLastRow(WS) <> 5 Then
                    TestFindLastRow = False
                    Debug.Print "Muliple cell Start 2, 2 Hidden Column"
                End If
            WS.Columns.Hidden = False
            'Hidden Row And Column
            WS.Columns(2).Hidden = True
            WS.Rows(2).Hidden = True
                If FindLastRow(WS) <> 5 Then
                    TestFindLastRow = False
                    Debug.Print "Muliple cell Start 2, 2 Hidden Row And Column"
                End If
            WS.Rows.Hidden = False
            WS.Columns.Hidden = False
            'Filter
            WS.Range("$B$2:$D$2").Clear
            WS.Range("$B$2:$D$5").AutoFilter Field:=1, Criteria1:="$B$4"
                If FindLastRow(WS) <> 4 Then
                    TestFindLastRow = False
                    Debug.Print "Muliple cell Start 2, 2 Filter"
                End If
            WS.AutoFilterMode = False
            
    'Last Cell
    WS.Cells.Clear
    WS.Range("A1:C1").Value = "Header"
    WS.Range("A2:C4").Formula = "=ADDRESS(ROW(),COLUMN())"
    WS.Range("C1:C3").Clear
    WS.Range("A4:B4").Clear
    If FindLastRow(WS) <> 4 Then
        TestFindLastRow = False
        Debug.Print "Last Cell"
    End If
    
    'Dynamic
    WS.Cells.Clear
    Dim i&
    For i = 1 To 5
        WS.Cells(i, i).Value = i
        If FindLastRow(WS) <> i Then
            TestFindLastRow = False
            Debug.Print "Dynamic"
        End If
    Next i
    
    Excel.Application.DisplayAlerts = False
    WB.Close
    Excel.Application.DisplayAlerts = True
    
    Excel.Application.ScreenUpdating = True
    
    Debug.Print "TestFindLastRow: " & TestFindLastRow
    
End Function

Private Function TestFindLastColumn() As Boolean

    TestFindLastColumn = True
    
    Excel.Application.ScreenUpdating = False
    
    Dim WB As Workbook
    Set WB = Excel.Application.Workbooks.Add
    
    Dim WS As Worksheet
    Set WS = WB.Worksheets(1)
    
    'Blank
        'Unhidden
            If FindLastColumn(WS) <> 0 Then
                TestFindLastColumn = False
                Debug.Print "Blank Unhidden"
            End If
        'Hidden Row
        WS.Rows(1).Hidden = True
            If FindLastColumn(WS) <> 0 Then
                TestFindLastColumn = False
                Debug.Print "Blank Hidden Row"
            End If
        WS.Rows.Hidden = False
        'Hidden Column
        WS.Columns(1).Hidden = True
            If FindLastColumn(WS) <> 0 Then
                TestFindLastColumn = False
                Debug.Print "Blank Hidden Column"
            End If
        WS.Columns.Hidden = False
        'Hidden Row And Column
        WS.Columns(1).Hidden = True
        WS.Rows(1).Hidden = True
            If FindLastColumn(WS) <> 0 Then
                TestFindLastColumn = False
                Debug.Print "Blank Hidden Row And Column"
            End If
        WS.Rows.Hidden = False
        WS.Columns.Hidden = False
        'Filter
            'Cannot filter blank
            
    'Single cell
        'Start 1, 1
        WS.Cells(1, 1).Value = "Test"
            'Unhidden
                If FindLastColumn(WS) <> 1 Then
                    TestFindLastColumn = False
                    Debug.Print "Single cell Start 1, 1 Unhidden"
                End If
            'Hidden Row
            WS.Rows(1).Hidden = True
                If FindLastColumn(WS) <> 1 Then
                    TestFindLastColumn = False
                    Debug.Print "Single cell Start 1, 1 Hidden Row"
                End If
            WS.Rows.Hidden = False
            'Hidden Column
            WS.Columns(1).Hidden = True
                If FindLastColumn(WS) <> 1 Then
                    TestFindLastColumn = False
                    Debug.Print "Single cell Start 1, 1 Hidden Column"
                End If
            WS.Columns.Hidden = False
            'Hidden Row And Column
            WS.Columns(1).Hidden = True
            WS.Rows(1).Hidden = True
                If FindLastColumn(WS) <> 1 Then
                    TestFindLastColumn = False
                    Debug.Print "Single cell Start 1, 1 Hidden Row And Column"
                End If
            WS.Rows.Hidden = False
            WS.Columns.Hidden = False
            'Filter
                'Cannot filter single cell
        'Start 2, 2
        WS.Cells.Clear
        WS.Cells(2, 2).Value = "Test"
            'Unhidden
                If FindLastColumn(WS) <> 2 Then
                    TestFindLastColumn = False
                    Debug.Print "Single cell Start 2, 2 Unhidden"
                End If
            'Hidden Row
            WS.Rows(2).Hidden = True
                If FindLastColumn(WS) <> 2 Then
                    TestFindLastColumn = False
                    Debug.Print "Single cell Start 2, 2 Hidden Row"
                End If
            WS.Rows.Hidden = False
            'Hidden Column
            WS.Columns(2).Hidden = True
                If FindLastColumn(WS) <> 2 Then
                    TestFindLastColumn = False
                    Debug.Print "Single cell Start 2, 2 Hidden Column"
                End If
            WS.Columns.Hidden = False
            'Hidden Row And Column
            WS.Columns(2).Hidden = True
            WS.Rows(2).Hidden = True
                If FindLastColumn(WS) <> 2 Then
                    TestFindLastColumn = False
                    Debug.Print "Single cell Start 2, 2 Hidden Row And Column"
                End If
            WS.Rows.Hidden = False
            WS.Columns.Hidden = False
            'Filter
                'Cannot filter single cell
                
    'Muliple cell
        'Start 1, 1
        WS.Cells.Clear
        WS.Range("A1:C1").Value = "Header"
        WS.Range("A2:C4").Formula = "=ADDRESS(ROW(),COLUMN())"
            'Unhidden
                If FindLastColumn(WS) <> 3 Then
                    TestFindLastColumn = False
                    Debug.Print "Muliple cell Start 1, 1 Unhidden"
                End If
            'Hidden Row
            WS.Rows(1).Hidden = True
                If FindLastColumn(WS) <> 3 Then
                    TestFindLastColumn = False
                    Debug.Print "Muliple cell Start 1, 1 Hidden Row"
                End If
            WS.Rows.Hidden = False
            'Hidden Column
            WS.Columns(1).Hidden = True
                If FindLastColumn(WS) <> 3 Then
                    TestFindLastColumn = False
                    Debug.Print "Muliple cell Start 1, 1 Hidden Column"
                End If
            WS.Columns.Hidden = False
            'Hidden Row And Column
            WS.Columns(1).Hidden = True
            WS.Rows(1).Hidden = True
                If FindLastColumn(WS) <> 3 Then
                    TestFindLastColumn = False
                    Debug.Print "Muliple cell Start 1, 1 Hidden Row And Column"
                End If
            WS.Rows.Hidden = False
            WS.Columns.Hidden = False
            'Filter
            WS.Range("A1:C1").Clear
            WS.Range("$A$1:$C$4").AutoFilter Field:=1, Criteria1:="$A$3"
                If FindLastColumn(WS) <> 3 Then
                    TestFindLastColumn = False
                    Debug.Print "Muliple cell Start 1, 1 Filter"
                End If
            WS.AutoFilterMode = False
        'Start 2, 2
        WS.Cells.Clear
        WS.Range("B2:D2").Value = "Header"
        WS.Range("B3:D5").Formula = "=ADDRESS(ROW(),COLUMN())"
            'Unhidden
                If FindLastColumn(WS) <> 4 Then
                    TestFindLastColumn = False
                    Debug.Print "Muliple cell Start 2, 2 Unhidden"
                End If
            'Hidden Row
            WS.Rows(2).Hidden = True
                If FindLastColumn(WS) <> 4 Then
                    TestFindLastColumn = False
                    Debug.Print "Muliple cell Start 2, 2 Hidden Row"
                End If
            WS.Rows.Hidden = False
            'Hidden Column
            WS.Columns(2).Hidden = True
                If FindLastColumn(WS) <> 4 Then
                    TestFindLastColumn = False
                    Debug.Print "Muliple cell Start 2, 2 Hidden Column"
                End If
            WS.Columns.Hidden = False
            'Hidden Row And Column
            WS.Columns(2).Hidden = True
            WS.Rows(2).Hidden = True
                If FindLastColumn(WS) <> 4 Then
                    TestFindLastColumn = False
                    Debug.Print "Muliple cell Start 2, 2 Hidden Row And Column"
                End If
            WS.Rows.Hidden = False
            WS.Columns.Hidden = False
            'Filter
            WS.Range("$B$2:$D$2").Clear
            WS.Range("$B$2:$D$5").AutoFilter Field:=1, Criteria1:="$B$4"
                If FindLastColumn(WS) <> 4 Then
                    TestFindLastColumn = False
                    Debug.Print "Muliple cell Start 2, 2 Filter"
                End If
            WS.AutoFilterMode = False
            
    'Last Cell
    WS.Cells.Clear
    WS.Range("A1:C1").Value = "Header"
    WS.Range("A2:C4").Formula = "=ADDRESS(ROW(),COLUMN())"
    WS.Range("C1:C3").Clear
    WS.Range("A4:B4").Clear
    If FindLastColumn(WS) <> 3 Then
        TestFindLastColumn = False
        Debug.Print "Last Cell"
    End If
    
    'Dynamic
    WS.Cells.Clear
    Dim i&
    For i = 1 To 5
        WS.Cells(i, i).Value = i
        If FindLastColumn(WS) <> i Then
            TestFindLastColumn = False
            Debug.Print "Dynamic"
        End If
    Next i
    
    Excel.Application.DisplayAlerts = False
    WB.Close
    Excel.Application.DisplayAlerts = True
    
    Excel.Application.ScreenUpdating = True
    
    Debug.Print "TestFindLastColumn: " & TestFindLastColumn
    
End Function

Private Function TestGetFirstRow() As Boolean

    TestGetFirstRow = True

    Excel.Application.ScreenUpdating = False

    Dim WB As Workbook
    Set WB = Excel.Application.Workbooks.Add

    Dim WS As Worksheet
    Set WS = WB.Worksheets(1)

    'Blank
        'Unhidden
            If GetFirstRow(WS) <> 0 Then
                TestGetFirstRow = False
                Debug.Print "Blank Unhidden"
            End If
            If GetFirstRow(WS, True, True) <> 0 Then
                TestGetFirstRow = False
                Debug.Print "Blank Unhidden True True"
            End If
            If GetFirstRow(WS, True, False) <> 0 Then
                TestGetFirstRow = False
                Debug.Print "Blank Unhidden True False"
            End If
            If GetFirstRow(WS, False, True) <> 0 Then
                TestGetFirstRow = False
                Debug.Print "Blank Unhidden False True"
            End If
            If GetFirstRow(WS, False, False) <> 0 Then
                TestGetFirstRow = False
                Debug.Print "Blank Unhidden False False"
            End If
        'Hidden Row
        WS.Rows(1).Hidden = True
            If GetFirstRow(WS) <> 0 Then
                TestGetFirstRow = False
                Debug.Print "Blank Hidden Row"
            End If
            If GetFirstRow(WS, True, True) <> 0 Then
                TestGetFirstRow = False
                Debug.Print "Blank Hidden Row True True"
            End If
            If GetFirstRow(WS, True, False) <> 0 Then
                TestGetFirstRow = False
                Debug.Print "Blank Hidden Row True False"
            End If
            If GetFirstRow(WS, False, True) <> 0 Then
                TestGetFirstRow = False
                Debug.Print "Blank Hidden Row False True"
            End If
            If GetFirstRow(WS, False, False) <> 0 Then
                TestGetFirstRow = False
                Debug.Print "Blank Hidden Row False False"
            End If
        WS.Rows.Hidden = False
        'Hidden Column
        WS.Columns(1).Hidden = True
            If GetFirstRow(WS) <> 0 Then
                TestGetFirstRow = False
                Debug.Print "Blank Hidden Column"
            End If
            If GetFirstRow(WS, True, True) <> 0 Then
                TestGetFirstRow = False
                Debug.Print "Blank Hidden Column True True"
            End If
            If GetFirstRow(WS, True, False) <> 0 Then
                TestGetFirstRow = False
                Debug.Print "Blank Hidden Column True False"
            End If
            If GetFirstRow(WS, False, True) <> 0 Then
                TestGetFirstRow = False
                Debug.Print "Blank Hidden Column False True"
            End If
            If GetFirstRow(WS, False, False) <> 0 Then
                TestGetFirstRow = False
                Debug.Print "Blank Hidden Column False False"
            End If
        WS.Columns.Hidden = False
        'Hidden Row And Column
        WS.Columns(1).Hidden = True
        WS.Rows(1).Hidden = True
            If GetFirstRow(WS) <> 0 Then
                TestGetFirstRow = False
                Debug.Print "Blank Hidden Row And Column"
            End If
            If GetFirstRow(WS, True, True) <> 0 Then
                TestGetFirstRow = False
                Debug.Print "Blank Hidden Row And Column True True"
            End If
            If GetFirstRow(WS, True, False) <> 0 Then
                TestGetFirstRow = False
                Debug.Print "Blank Hidden Row And Column True False"
            End If
            If GetFirstRow(WS, False, True) <> 0 Then
                TestGetFirstRow = False
                Debug.Print "Blank Hidden Row And Column False True"
            End If
            If GetFirstRow(WS, False, False) <> 0 Then
                TestGetFirstRow = False
                Debug.Print "Blank Hidden Row And Column False False"
            End If
        WS.Rows.Hidden = False
        WS.Columns.Hidden = False
        'Filter
            'Cannot filter blank
            
    'Single cell
        'Start 1, 1
        WS.Cells(1, 1).Value = "Test"
            'Unhidden
                If GetFirstRow(WS) <> 1 Then
                    TestGetFirstRow = False
                    Debug.Print "Single cell Start 1, 1 Unhidden"
                End If
                If GetFirstRow(WS, True, True) <> 1 Then
                    TestGetFirstRow = False
                    Debug.Print "Single cell Start 1, 1 Unhidden True True"
                End If
                If GetFirstRow(WS, True, False) <> 1 Then
                    TestGetFirstRow = False
                    Debug.Print "Single cell Start 1, 1 Unhidden True False"
                End If
                If GetFirstRow(WS, False, True) <> 1 Then
                    TestGetFirstRow = False
                    Debug.Print "Single cell Start 1, 1 Unhidden False True"
                End If
                If GetFirstRow(WS, False, False) <> 1 Then
                    TestGetFirstRow = False
                    Debug.Print "Single cell Start 1, 1 Unhidden False False"
                End If
            'Hidden Row
            WS.Rows(1).Hidden = True
                If GetFirstRow(WS) <> 1 Then
                    TestGetFirstRow = False
                    Debug.Print "Single cell Start 1, 1 Hidden Row"
                End If
                If GetFirstRow(WS, True, True) <> 1 Then
                    TestGetFirstRow = False
                    Debug.Print "Single cell Start 1, 1 Hidden Row True True"
                End If
                If GetFirstRow(WS, True, False) <> 1 Then
                    TestGetFirstRow = False
                    Debug.Print "Single cell Start 1, 1 Hidden Row True False"
                End If
                If GetFirstRow(WS, False, True) <> 0 Then
                    TestGetFirstRow = False
                    Debug.Print "Single cell Start 1, 1 Hidden Row False True"
                End If
                If GetFirstRow(WS, False, False) <> 0 Then
                    TestGetFirstRow = False
                    Debug.Print "Single cell Start 1, 1 Hidden Row False False"
                End If
            WS.Rows.Hidden = False
            'Hidden Column
            WS.Columns(1).Hidden = True
                If GetFirstRow(WS) <> 1 Then
                    TestGetFirstRow = False
                    Debug.Print "Single cell Start 1, 1 Hidden Column"
                End If
                If GetFirstRow(WS, True, True) <> 1 Then
                    TestGetFirstRow = False
                    Debug.Print "Single cell Start 1, 1 Hidden Column True True"
                End If
                If GetFirstRow(WS, True, False) <> 0 Then
                    TestGetFirstRow = False
                    Debug.Print "Single cell Start 1, 1 Hidden Column True False"
                End If
                If GetFirstRow(WS, False, True) <> 1 Then
                    TestGetFirstRow = False
                    Debug.Print "Single cell Start 1, 1 Hidden Column False True"
                End If
                If GetFirstRow(WS, False, False) <> 0 Then
                    TestGetFirstRow = False
                    Debug.Print "Single cell Start 1, 1 Hidden Column False False"
                End If
            WS.Columns.Hidden = False
            'Hidden Row And Column
            WS.Columns(1).Hidden = True
            WS.Rows(1).Hidden = True
                If GetFirstRow(WS) <> 1 Then
                    TestGetFirstRow = False
                    Debug.Print "Single cell Start 1, 1 Hidden Row And Column"
                End If
                If GetFirstRow(WS, True, True) <> 1 Then
                    TestGetFirstRow = False
                    Debug.Print "Single cell Start 1, 1 Hidden Row And Column True True"
                End If
                If GetFirstRow(WS, True, False) <> 0 Then
                    TestGetFirstRow = False
                    Debug.Print "Single cell Start 1, 1 Hidden Row And Column True False"
                End If
                If GetFirstRow(WS, False, True) <> 0 Then
                    TestGetFirstRow = False
                    Debug.Print "Single cell Start 1, 1 Hidden Row And Column False True"
                End If
                If GetFirstRow(WS, False, False) <> 0 Then
                    TestGetFirstRow = False
                    Debug.Print "Single cell Start 1, 1 Hidden Row And Column False False"
                End If
            WS.Rows.Hidden = False
            WS.Columns.Hidden = False
            'Filter
                'Cannot filter single cell
        'Start 2, 2
        WS.Cells.Clear
        WS.Cells(2, 2).Value = "Test"
            'Unhidden
                If GetFirstRow(WS) <> 2 Then
                    TestGetFirstRow = False
                    Debug.Print "Single cell Start 2, 2 Unhidden"
                End If
                If GetFirstRow(WS, True, True) <> 2 Then
                    TestGetFirstRow = False
                    Debug.Print "Single cell Start 2, 2 Unhidden True True"
                End If
                If GetFirstRow(WS, True, False) <> 2 Then
                    TestGetFirstRow = False
                    Debug.Print "Single cell Start 2, 2 Unhidden True False"
                End If
                If GetFirstRow(WS, False, True) <> 2 Then
                    TestGetFirstRow = False
                    Debug.Print "Single cell Start 2, 2 Unhidden False True"
                End If
                If GetFirstRow(WS, False, False) <> 2 Then
                    TestGetFirstRow = False
                    Debug.Print "Single cell Start 2, 2 Unhidden False False"
                End If
            'Hidden Row
            WS.Rows(2).Hidden = True
                If GetFirstRow(WS) <> 2 Then
                    TestGetFirstRow = False
                    Debug.Print "Single cell Start 2, 2 Hidden Row"
                End If
                If GetFirstRow(WS, True, True) <> 2 Then
                    TestGetFirstRow = False
                    Debug.Print "Single cell Start 2, 2 Hidden Row True True"
                End If
                If GetFirstRow(WS, True, False) <> 2 Then
                    TestGetFirstRow = False
                    Debug.Print "Single cell Start 2, 2 Hidden Row True False"
                End If
                If GetFirstRow(WS, False, True) <> 0 Then
                    TestGetFirstRow = False
                    Debug.Print "Single cell Start 2, 2 Hidden Row False True"
                End If
                If GetFirstRow(WS, False, False) <> 0 Then
                    TestGetFirstRow = False
                    Debug.Print "Single cell Start 2, 2 Hidden Row False False"
                End If
            WS.Rows.Hidden = False
            'Hidden Column
            WS.Columns(2).Hidden = True
                If GetFirstRow(WS) <> 2 Then
                    TestGetFirstRow = False
                    Debug.Print "Single cell Start 2, 2 Hidden Column"
                End If
                If GetFirstRow(WS, True, True) <> 2 Then
                    TestGetFirstRow = False
                    Debug.Print "Single cell Start 2, 2 Hidden Column True True"
                End If
                If GetFirstRow(WS, True, False) <> 0 Then
                    TestGetFirstRow = False
                    Debug.Print "Single cell Start 2, 2 Hidden Column True False"
                End If
                If GetFirstRow(WS, False, True) <> 2 Then
                    TestGetFirstRow = False
                    Debug.Print "Single cell Start 2, 2 Hidden Column False True"
                End If
                If GetFirstRow(WS, False, False) <> 0 Then
                    TestGetFirstRow = False
                    Debug.Print "Single cell Start 2, 2 Hidden Column False False"
                End If
            WS.Columns.Hidden = False
            'Hidden Row And Column
            WS.Columns(2).Hidden = True
            WS.Rows(2).Hidden = True
                If GetFirstRow(WS) <> 2 Then
                    TestGetFirstRow = False
                    Debug.Print "Single cell Start 2, 2 Hidden Row And Column"
                End If
                If GetFirstRow(WS, True, True) <> 2 Then
                    TestGetFirstRow = False
                    Debug.Print "Single cell Start 2, 2 Hidden Row And Column True True"
                End If
                If GetFirstRow(WS, True, False) <> 0 Then
                    TestGetFirstRow = False
                    Debug.Print "Single cell Start 2, 2 Hidden Row And Column True False"
                End If
                If GetFirstRow(WS, False, True) <> 0 Then
                    TestGetFirstRow = False
                    Debug.Print "Single cell Start 2, 2 Hidden Row And Column False True"
                End If
                If GetFirstRow(WS, False, False) <> 0 Then
                    TestGetFirstRow = False
                    Debug.Print "Single cell Start 2, 2 Hidden Row And Column False False"
                End If
            WS.Rows.Hidden = False
            WS.Columns.Hidden = False
            'Filter
                'Cannot filter single cell
                
    'Muliple cell
        'Start 1, 1
        WS.Cells.Clear
        WS.Range("A1:C1").Value = "Header"
        WS.Range("A2:C4").Formula = "=ADDRESS(ROW(),COLUMN())"
            'Unhidden
                If GetFirstRow(WS) <> 1 Then
                    TestGetFirstRow = False
                    Debug.Print "Muliple cell Start 1, 1 Unhidden"
                End If
                If GetFirstRow(WS, True, True) <> 1 Then
                    TestGetFirstRow = False
                    Debug.Print "Muliple cell Start 1, 1 Unhidden True True"
                End If
                If GetFirstRow(WS, True, False) <> 1 Then
                    TestGetFirstRow = False
                    Debug.Print "Muliple cell Start 1, 1 Unhidden True False"
                End If
                If GetFirstRow(WS, False, True) <> 1 Then
                    TestGetFirstRow = False
                    Debug.Print "Muliple cell Start 1, 1 Unhidden False True"
                End If
                If GetFirstRow(WS, False, False) <> 1 Then
                    TestGetFirstRow = False
                    Debug.Print "Muliple cell Start 1, 1 Unhidden False False"
                End If
            'Hidden Row
            WS.Rows(1).Hidden = True
                If GetFirstRow(WS) <> 1 Then
                    TestGetFirstRow = False
                    Debug.Print "Muliple cell Start 1, 1 Hidden Row"
                End If
                If GetFirstRow(WS, True, True) <> 1 Then
                    TestGetFirstRow = False
                    Debug.Print "Muliple cell Start 1, 1 Hidden Row True True"
                End If
                If GetFirstRow(WS, True, False) <> 1 Then
                    TestGetFirstRow = False
                    Debug.Print "Muliple cell Start 1, 1 Hidden Row True False"
                End If
                If GetFirstRow(WS, False, True) <> 2 Then
                    TestGetFirstRow = False
                    Debug.Print "Muliple cell Start 1, 1 Hidden Row False True"
                End If
                If GetFirstRow(WS, False, False) <> 2 Then
                    TestGetFirstRow = False
                    Debug.Print "Muliple cell Start 1, 1 Hidden Row False False"
                End If
            WS.Rows.Hidden = False
            'Hidden Column
            WS.Columns(1).Hidden = True
                If GetFirstRow(WS) <> 1 Then
                    TestGetFirstRow = False
                    Debug.Print "Muliple cell Start 1, 1 Hidden Column"
                End If
                If GetFirstRow(WS, True, True) <> 1 Then
                    TestGetFirstRow = False
                    Debug.Print "Muliple cell Start 1, 1 Hidden Column True True"
                End If
                If GetFirstRow(WS, True, False) <> 1 Then
                    TestGetFirstRow = False
                    Debug.Print "Muliple cell Start 1, 1 Hidden Column True False"
                End If
                If GetFirstRow(WS, False, True) <> 1 Then
                    TestGetFirstRow = False
                    Debug.Print "Muliple cell Start 1, 1 Hidden Column False True"
                End If
                If GetFirstRow(WS, False, False) <> 1 Then
                    TestGetFirstRow = False
                    Debug.Print "Muliple cell Start 1, 1 Hidden Column False False"
                End If
            WS.Columns.Hidden = False
            'Hidden Row And Column
            WS.Columns(1).Hidden = True
            WS.Rows(1).Hidden = True
                If GetFirstRow(WS) <> 1 Then
                    TestGetFirstRow = False
                    Debug.Print "Muliple cell Start 1, 1 Hidden Row And Column"
                End If
                If GetFirstRow(WS, True, True) <> 1 Then
                    TestGetFirstRow = False
                    Debug.Print "Muliple cell Start 1, 1 Hidden Row And Column True True"
                End If
                If GetFirstRow(WS, True, False) <> 1 Then
                    TestGetFirstRow = False
                    Debug.Print "Muliple cell Start 1, 1 Hidden Row And Column True False"
                End If
                If GetFirstRow(WS, False, True) <> 2 Then
                    TestGetFirstRow = False
                    Debug.Print "Muliple cell Start 1, 1 Hidden Row And Column False True"
                End If
                If GetFirstRow(WS, False, False) <> 2 Then
                    TestGetFirstRow = False
                    Debug.Print "Muliple cell Start 1, 1 Hidden Row And Column False False"
                End If
            WS.Rows.Hidden = False
            WS.Columns.Hidden = False
            'Filter
            WS.Range("A1:C1").Clear
            WS.Range("$A$1:$C$4").AutoFilter Field:=1, Criteria1:="$A$3"
                If GetFirstRow(WS) <> 2 Then
                    TestGetFirstRow = False
                    Debug.Print "Muliple cell Start 1, 1 Filter"
                End If
                If GetFirstRow(WS, True, True) <> 2 Then
                    TestGetFirstRow = False
                    Debug.Print "Muliple cell Start 1, 1 Filter True True"
                End If
                If GetFirstRow(WS, True, False) <> 2 Then
                    TestGetFirstRow = False
                    Debug.Print "Muliple cell Start 1, 1 Filter True False"
                End If
                If GetFirstRow(WS, False, True) <> 3 Then
                    TestGetFirstRow = False
                    Debug.Print "Muliple cell Start 1, 1 Filter False True"
                End If
                If GetFirstRow(WS, False, False) <> 3 Then
                    TestGetFirstRow = False
                    Debug.Print "Muliple cell Start 1, 1 Filter False False"
                End If
            WS.AutoFilterMode = False
        'Start 2, 2
        WS.Cells.Clear
        WS.Range("B2:D2").Value = "Header"
        WS.Range("B3:D5").Formula = "=ADDRESS(ROW(),COLUMN())"
            'Unhidden
                If GetFirstRow(WS) <> 2 Then
                    TestGetFirstRow = False
                    Debug.Print "Muliple cell Start 2, 2 Unhidden"
                End If
                If GetFirstRow(WS, True, True) <> 2 Then
                    TestGetFirstRow = False
                    Debug.Print "Muliple cell Start 2, 2 Unhidden True True"
                End If
                If GetFirstRow(WS, True, False) <> 2 Then
                    TestGetFirstRow = False
                    Debug.Print "Muliple cell Start 2, 2 Unhidden True False"
                End If
                If GetFirstRow(WS, False, True) <> 2 Then
                    TestGetFirstRow = False
                    Debug.Print "Muliple cell Start 2, 2 Unhidden False True"
                End If
                If GetFirstRow(WS, False, False) <> 2 Then
                    TestGetFirstRow = False
                    Debug.Print "Muliple cell Start 2, 2 Unhidden False False"
                End If
            'Hidden Row
            WS.Rows(2).Hidden = True
                If GetFirstRow(WS) <> 2 Then
                    TestGetFirstRow = False
                    Debug.Print "Muliple cell Start 2, 2 Hidden Row"
                End If
                If GetFirstRow(WS, True, True) <> 2 Then
                    TestGetFirstRow = False
                    Debug.Print "Muliple cell Start 2, 2 Hidden Row True True"
                End If
                If GetFirstRow(WS, True, False) <> 2 Then
                    TestGetFirstRow = False
                    Debug.Print "Muliple cell Start 2, 2 Hidden Row True False"
                End If
                If GetFirstRow(WS, False, True) <> 3 Then
                    TestGetFirstRow = False
                    Debug.Print "Muliple cell Start 2, 2 Hidden Row False True"
                End If
                If GetFirstRow(WS, False, False) <> 3 Then
                    TestGetFirstRow = False
                    Debug.Print "Muliple cell Start 2, 2 Hidden Row False False"
                End If
            WS.Rows.Hidden = False
            'Hidden Column
            WS.Columns(2).Hidden = True
                If GetFirstRow(WS) <> 2 Then
                    TestGetFirstRow = False
                    Debug.Print "Muliple cell Start 2, 2 Hidden Column"
                End If
                If GetFirstRow(WS, True, True) <> 2 Then
                    TestGetFirstRow = False
                    Debug.Print "Muliple cell Start 2, 2 Hidden Column True True"
                End If
                If GetFirstRow(WS, True, False) <> 2 Then
                    TestGetFirstRow = False
                    Debug.Print "Muliple cell Start 2, 2 Hidden Column True False"
                End If
                If GetFirstRow(WS, False, True) <> 2 Then
                    TestGetFirstRow = False
                    Debug.Print "Muliple cell Start 2, 2 Hidden Column False True"
                End If
                If GetFirstRow(WS, False, False) <> 2 Then
                    TestGetFirstRow = False
                    Debug.Print "Muliple cell Start 2, 2 Hidden Column False False"
                End If
            WS.Columns.Hidden = False
            'Hidden Row And Column
            WS.Columns(2).Hidden = True
            WS.Rows(2).Hidden = True
                If GetFirstRow(WS) <> 2 Then
                    TestGetFirstRow = False
                    Debug.Print "Muliple cell Start 2, 2 Hidden Row And Column"
                End If
                If GetFirstRow(WS, True, True) <> 2 Then
                    TestGetFirstRow = False
                    Debug.Print "Muliple cell Start 2, 2 Hidden Row And Column True True"
                End If
                If GetFirstRow(WS, True, False) <> 2 Then
                    TestGetFirstRow = False
                    Debug.Print "Muliple cell Start 2, 2 Hidden Row And Column True False"
                End If
                If GetFirstRow(WS, False, True) <> 3 Then
                    TestGetFirstRow = False
                    Debug.Print "Muliple cell Start 2, 2 Hidden Row And Column False True"
                End If
                If GetFirstRow(WS, False, False) <> 3 Then
                    TestGetFirstRow = False
                    Debug.Print "Muliple cell Start 2, 2 Hidden Row And Column False False"
                End If
            WS.Rows.Hidden = False
            WS.Columns.Hidden = False
            'Filter
            WS.Range("$B$2:$D$2").Clear
            WS.Range("$B$2:$D$5").AutoFilter Field:=1, Criteria1:="$B$4"
                If GetFirstRow(WS) <> 3 Then
                    TestGetFirstRow = False
                    Debug.Print "Muliple cell Start 2, 2 Filter"
                End If
                If GetFirstRow(WS, True, True) <> 3 Then
                    TestGetFirstRow = False
                    Debug.Print "Muliple cell Start 2, 2 Filter True True"
                End If
                If GetFirstRow(WS, True, False) <> 3 Then
                    TestGetFirstRow = False
                    Debug.Print "Muliple cell Start 2, 2 Filter True False"
                End If
                If GetFirstRow(WS, False, True) <> 4 Then
                    TestGetFirstRow = False
                    Debug.Print "Muliple cell Start 2, 2 Filter False True"
                End If
                If GetFirstRow(WS, False, False) <> 4 Then
                    TestGetFirstRow = False
                    Debug.Print "Muliple cell Start 2, 2 Filter False False"
                End If
            WS.AutoFilterMode = False
            
    'Dynamic
    WS.Cells.Clear
    Dim i&
    For i = 5 To 1 Step -1
        WS.Cells(i, i).Value = i
        If GetFirstRow(WS) <> i Then
            TestGetFirstRow = False
            Debug.Print "Dynamic"
        End If
    Next i

    Excel.Application.DisplayAlerts = False
    WB.Close
    Excel.Application.DisplayAlerts = True
    
    Excel.Application.ScreenUpdating = True
    
    Debug.Print "TestGetFirstRow: " & TestGetFirstRow
    
End Function

Private Function TestGetFirstRowInColumn() As Boolean

    TestGetFirstRowInColumn = True

    Excel.Application.ScreenUpdating = False

    Dim WB As Workbook
    Set WB = Excel.Application.Workbooks.Add

    Dim WS As Worksheet
    Set WS = WB.Worksheets(1)

    'Blank
        'Unhidden
            If GetFirstRowInColumn(WS, 1) <> 0 Then
                TestGetFirstRowInColumn = False
                Debug.Print "Blank Unhidden"
            End If
            If GetFirstRowInColumn(WS, 1, True, True) <> 0 Then
                TestGetFirstRowInColumn = False
                Debug.Print "Blank Unhidden True True"
            End If
            If GetFirstRowInColumn(WS, 1, True, False) <> 0 Then
                TestGetFirstRowInColumn = False
                Debug.Print "Blank Unhidden True False"
            End If
            If GetFirstRowInColumn(WS, 1, False, True) <> 0 Then
                TestGetFirstRowInColumn = False
                Debug.Print "Blank Unhidden False True"
            End If
            If GetFirstRowInColumn(WS, 1, False, False) <> 0 Then
                TestGetFirstRowInColumn = False
                Debug.Print "Blank Unhidden False False"
            End If
        'Hidden Rows
        WS.Rows(1).Hidden = True
            If GetFirstRowInColumn(WS, 1) <> 0 Then
                TestGetFirstRowInColumn = False
                Debug.Print "Blank Hidden Rows"
            End If
            If GetFirstRowInColumn(WS, 1, True, True) <> 0 Then
                TestGetFirstRowInColumn = False
                Debug.Print "Blank Hidden Rows True True"
            End If
            If GetFirstRowInColumn(WS, 1, True, False) <> 0 Then
                TestGetFirstRowInColumn = False
                Debug.Print "Blank Hidden Rows True False"
            End If
            If GetFirstRowInColumn(WS, 1, False, True) <> 0 Then
                TestGetFirstRowInColumn = False
                Debug.Print "Blank Hidden Rows False True"
            End If
            If GetFirstRowInColumn(WS, 1, False, False) <> 0 Then
                TestGetFirstRowInColumn = False
                Debug.Print "Blank Hidden Rows False False"
            End If
        WS.Rows.Hidden = False
        'Hidden Columns
        WS.Columns(1).Hidden = True
            If GetFirstRowInColumn(WS, 1) <> 0 Then
                TestGetFirstRowInColumn = False
                Debug.Print "Blank Hidden Columns"
            End If
            If GetFirstRowInColumn(WS, 1, True, True) <> 0 Then
                TestGetFirstRowInColumn = False
                Debug.Print "Blank Hidden Columns True True"
            End If
            If GetFirstRowInColumn(WS, 1, True, False) <> 0 Then
                TestGetFirstRowInColumn = False
                Debug.Print "Blank Hidden Columns True False"
            End If
            If GetFirstRowInColumn(WS, 1, False, True) <> 0 Then
                TestGetFirstRowInColumn = False
                Debug.Print "Blank Hidden Columns False True"
            End If
            If GetFirstRowInColumn(WS, 1, False, False) <> 0 Then
                TestGetFirstRowInColumn = False
                Debug.Print "Blank Hidden Columns False False"
            End If
        WS.Columns.Hidden = False
        'Hidden Rows And Columns
        WS.Rows(1).Hidden = True
        WS.Columns(1).Hidden = True
            If GetFirstRowInColumn(WS, 1) <> 0 Then
                TestGetFirstRowInColumn = False
                Debug.Print "Blank Hidden Rows And Columns"
            End If
            If GetFirstRowInColumn(WS, 1, True, True) <> 0 Then
                TestGetFirstRowInColumn = False
                Debug.Print "Blank Hidden Rows And Columns True True"
            End If
            If GetFirstRowInColumn(WS, 1, True, False) <> 0 Then
                TestGetFirstRowInColumn = False
                Debug.Print "Blank Hidden Rows And Columns True False"
            End If
            If GetFirstRowInColumn(WS, 1, False, True) <> 0 Then
                TestGetFirstRowInColumn = False
                Debug.Print "Blank Hidden Rows And Columns False True"
            End If
            If GetFirstRowInColumn(WS, 1, False, False) <> 0 Then
                TestGetFirstRowInColumn = False
                Debug.Print "Blank Hidden Rows And Columns False False"
            End If
        WS.Rows.Hidden = False
        WS.Columns.Hidden = False
        'Filter
            'Cannot filter blank
            
    'Single cell
        'Start 1, 1
        WS.Cells(1, 1).Value = "Test"
            'Unhidden
                If GetFirstRowInColumn(WS, 1) <> 1 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Single cell Start 1, 1 Unhidden"
                End If
                If GetFirstRowInColumn(WS, 1, True, True) <> 1 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Single cell Start 1, 1 Unhidden True True"
                End If
                If GetFirstRowInColumn(WS, 1, True, False) <> 1 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Single cell Start 1, 1 Unhidden True False"
                End If
                If GetFirstRowInColumn(WS, 1, False, True) <> 1 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Single cell Start 1, 1 Unhidden False True"
                End If
                If GetFirstRowInColumn(WS, 1, False, False) <> 1 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Single cell Start 1, 1 Unhidden False False"
                End If
                'Outside
                If GetFirstRowInColumn(WS, 4) <> 0 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Single cell Start 1, 1 Unhidden Outside more"
                End If
            'Hidden Rows
            WS.Rows(1).Hidden = True
                If GetFirstRowInColumn(WS, 1) <> 1 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Single cell Start 1, 1 Hidden Rows"
                End If
                If GetFirstRowInColumn(WS, 1, True, True) <> 1 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Single cell Start 1, 1 Hidden Rows True True"
                End If
                If GetFirstRowInColumn(WS, 1, True, False) <> 1 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Single cell Start 1, 1 Hidden Rows True False"
                End If
                If GetFirstRowInColumn(WS, 1, False, True) <> 0 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Single cell Start 1, 1 Hidden Rows False True"
                End If
                If GetFirstRowInColumn(WS, 1, False, False) <> 0 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Single cell Start 1, 1 Hidden Rows False False"
                End If
                'Outside
                If GetFirstRowInColumn(WS, 4) <> 0 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Single cell Start 1, 1 Hidden Rows Outside more"
                End If
            WS.Rows.Hidden = False
            'Hidden Columns
            WS.Columns(1).Hidden = True
                If GetFirstRowInColumn(WS, 1) <> 1 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Single cell Start 1, 1 Hidden Columns"
                End If
                If GetFirstRowInColumn(WS, 1, True, True) <> 1 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Single cell Start 1, 1 Hidden Columns True True"
                End If
                If GetFirstRowInColumn(WS, 1, True, False) <> 0 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Single cell Start 1, 1 Hidden Columns True False"
                End If
                If GetFirstRowInColumn(WS, 1, False, True) <> 1 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Single cell Start 1, 1 Hidden Columns False True"
                End If
                If GetFirstRowInColumn(WS, 1, False, False) <> 0 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Single cell Start 1, 1 Hidden Columns False False"
                End If
                'Outside
                If GetFirstRowInColumn(WS, 4) <> 0 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Single cell Start 1, 1 Hidden Columns Outside more"
                End If
            WS.Columns.Hidden = False
            'Hidden Rows And Columns
            WS.Rows(1).Hidden = True
            WS.Columns(1).Hidden = True
                If GetFirstRowInColumn(WS, 1) <> 1 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Single cell Start 1, 1 Hidden Rows And Columns"
                End If
                If GetFirstRowInColumn(WS, 1, True, True) <> 1 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Single cell Start 1, 1 Hidden Rows And Columns True True"
                End If
                If GetFirstRowInColumn(WS, 1, True, False) <> 0 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Single cell Start 1, 1 Hidden Rows And Columns True False"
                End If
                If GetFirstRowInColumn(WS, 1, False, True) <> 0 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Single cell Start 1, 1 Hidden Rows And Columns False True"
                End If
                If GetFirstRowInColumn(WS, 1, False, False) <> 0 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Single cell Start 1, 1 Hidden Rows And Columns False False"
                End If
                'Outside
                If GetFirstRowInColumn(WS, 4) <> 0 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Single cell Start 1, 1 Hidden Rows And Columns Outside more"
                End If
            WS.Rows.Hidden = False
            WS.Columns.Hidden = False
            'Filter
                'Cannot filter single cell

        'Start 2, 2
        WS.Cells.Clear
        WS.Cells(2, 2).Value = "Test"
            'Unhidden
                If GetFirstRowInColumn(WS, 2) <> 2 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Single cell Start 2, 2 Unhidden"
                End If
                If GetFirstRowInColumn(WS, 2, True, True) <> 2 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Single cell Start 2, 2 Unhidden True True"
                End If
                If GetFirstRowInColumn(WS, 2, True, False) <> 2 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Single cell Start 2, 2 Unhidden True False"
                End If
                If GetFirstRowInColumn(WS, 2, False, True) <> 2 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Single cell Start 2, 2 Unhidden False True"
                End If
                If GetFirstRowInColumn(WS, 2, False, False) <> 2 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Single cell Start 2, 2 Unhidden False False"
                End If
                'Outside
                If GetFirstRowInColumn(WS, 1) <> 0 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Single cell Start 2, 2 Unhidden Outside less"
                End If
                If GetFirstRowInColumn(WS, 5) <> 0 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Single cell Start 2, 2 Unhidden Outside more"
                End If
            'Hidden Rows
            WS.Rows(2).Hidden = True
                If GetFirstRowInColumn(WS, 2) <> 2 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Single cell Start 2, 2 Hidden Rows"
                End If
                If GetFirstRowInColumn(WS, 2, True, True) <> 2 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Single cell Start 2, 2 Hidden Rows True True"
                End If
                If GetFirstRowInColumn(WS, 2, True, False) <> 2 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Single cell Start 2, 2 Hidden Rows True False"
                End If
                If GetFirstRowInColumn(WS, 2, False, True) <> 0 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Single cell Start 2, 2 Hidden Rows False True"
                End If
                If GetFirstRowInColumn(WS, 2, False, False) <> 0 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Single cell Start 2, 2 Hidden Rows False False"
                End If
                'Outside
                If GetFirstRowInColumn(WS, 1) <> 0 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Single cell Start 2, 2 Hidden Rows Outside less"
                End If
                If GetFirstRowInColumn(WS, 5) <> 0 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Single cell Start 2, 2 Hidden Rows Outside more"
                End If
            WS.Rows.Hidden = False
            'Hidden Columns
            WS.Columns(2).Hidden = True
                If GetFirstRowInColumn(WS, 2) <> 2 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Single cell Start 2, 2 Hidden Columns"
                End If
                If GetFirstRowInColumn(WS, 2, True, True) <> 2 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Single cell Start 2, 2 Hidden Columns True True"
                End If
                If GetFirstRowInColumn(WS, 2, True, False) <> 0 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Single cell Start 2, 2 Hidden Columns True False"
                End If
                If GetFirstRowInColumn(WS, 2, False, True) <> 2 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Single cell Start 2, 2 Hidden Columns False True"
                End If
                If GetFirstRowInColumn(WS, 2, False, False) <> 0 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Single cell Start 2, 2 Hidden Columns False False"
                End If
                'Outside
                If GetFirstRowInColumn(WS, 1) <> 0 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Single cell Start 2, 2 Hidden Columns Outside less"
                End If
                If GetFirstRowInColumn(WS, 5) <> 0 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Single cell Start 2, 2 Hidden Columns Outside more"
                End If
            WS.Columns.Hidden = False
            'Hidden Rows And Columns
            WS.Rows(2).Hidden = True
            WS.Columns(2).Hidden = True
                If GetFirstRowInColumn(WS, 2) <> 2 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Single cell Start 2, 2 Hidden Rows And Columns"
                End If
                If GetFirstRowInColumn(WS, 2, True, True) <> 2 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Single cell Start 2, 2 Hidden Rows And Columns True True"
                End If
                If GetFirstRowInColumn(WS, 2, True, False) <> 0 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Single cell Start 2, 2 Hidden Rows And Columns True False"
                End If
                If GetFirstRowInColumn(WS, 2, False, True) <> 0 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Single cell Start 2, 2 Hidden Rows And Columns False True"
                End If
                If GetFirstRowInColumn(WS, 2, False, False) <> 0 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Single cell Start 2, 2 Hidden Rows And Columns False False"
                End If
                'Outside
                If GetFirstRowInColumn(WS, 1) <> 0 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Single cell Start 2, 2 Hidden Rows And Columns Outside less"
                End If
                If GetFirstRowInColumn(WS, 5) <> 0 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Single cell Start 2, 2 Hidden Rows And Columns Outside more"
                End If
            WS.Rows.Hidden = False
            WS.Columns.Hidden = False
            'Filter
                'Cannot filter single cell
    
    'Multiple cell
        'Start 1, 1
        WS.Cells.Clear
        WS.Range("A1:C1").Value = "Header"
        WS.Range("A2:C4").Formula = "=ADDRESS(ROW(),COLUMN())"
            'Unhidden
                If GetFirstRowInColumn(WS, 1) <> 1 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Multiple cell Start 1, 1 Unhidden"
                End If
                If GetFirstRowInColumn(WS, 1, True, True) <> 1 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Multiple cell Start 1, 1 Unhidden True True"
                End If
                If GetFirstRowInColumn(WS, 1, True, False) <> 1 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Multiple cell Start 1, 1 Unhidden True False"
                End If
                If GetFirstRowInColumn(WS, 1, False, True) <> 1 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Multiple cell Start 1, 1 Unhidden False True"
                End If
                If GetFirstRowInColumn(WS, 1, False, False) <> 1 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Multiple cell Start 1, 1 Unhidden False False"
                End If
                'Outside
                If GetFirstRowInColumn(WS, 4) <> 0 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Multiple cell Start 1, 1 Unhidden Outside more"
                End If
            'Hidden Rows
            WS.Rows(1).Hidden = True
                If GetFirstRowInColumn(WS, 1) <> 1 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Multiple cell Start 1, 1 Hidden Rows"
                End If
                If GetFirstRowInColumn(WS, 1, True, True) <> 1 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Multiple cell Start 1, 1 Hidden Rows True True"
                End If
                If GetFirstRowInColumn(WS, 1, True, False) <> 1 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Multiple cell Start 1, 1 Hidden Rows True False"
                End If
                If GetFirstRowInColumn(WS, 1, False, True) <> 2 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Multiple cell Start 1, 1 Hidden Rows False True"
                End If
                If GetFirstRowInColumn(WS, 1, False, False) <> 2 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Multiple cell Start 1, 1 Hidden Rows False False"
                End If
                'Outside
                If GetFirstRowInColumn(WS, 4) <> 0 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Multiple cell Start 1, 1 Hidden Rows Outside more"
                End If
            WS.Rows.Hidden = False
            'Hidden Columns
            WS.Columns(1).Hidden = True
                If GetFirstRowInColumn(WS, 1) <> 1 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Multiple cell Start 1, 1 Hidden Columns"
                End If
                If GetFirstRowInColumn(WS, 1, True, True) <> 1 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Multiple cell Start 1, 1 Hidden Columns True True"
                End If
                If GetFirstRowInColumn(WS, 1, True, False) <> 0 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Multiple cell Start 1, 1 Hidden Columns True False"
                End If
                If GetFirstRowInColumn(WS, 1, False, True) <> 1 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Multiple cell Start 1, 1 Hidden Columns False True"
                End If
                If GetFirstRowInColumn(WS, 1, False, False) <> 0 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Multiple cell Start 1, 1 Hidden Columns False False"
                End If
                'Outside
                If GetFirstRowInColumn(WS, 4) <> 0 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Multiple cell Start 1, 1 Hidden Columns Outside more"
                End If
            WS.Columns.Hidden = False
            'Hidden Rows And Columns
            WS.Rows(1).Hidden = True
            WS.Columns(1).Hidden = True
                If GetFirstRowInColumn(WS, 1) <> 1 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Multiple cell Start 1, 1 Hidden Rows And Columns"
                End If
                If GetFirstRowInColumn(WS, 1, True, True) <> 1 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Multiple cell Start 1, 1 Hidden Rows And Columns True True"
                End If
                If GetFirstRowInColumn(WS, 1, True, False) <> 0 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Multiple cell Start 1, 1 Hidden Rows And Columns True False"
                End If
                If GetFirstRowInColumn(WS, 1, False, True) <> 2 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Multiple cell Start 1, 1 Hidden Rows And Columns False True"
                End If
                If GetFirstRowInColumn(WS, 1, False, False) <> 0 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Multiple cell Start 1, 1 Hidden Rows And Columns False False"
                End If
                'Outside
                If GetFirstRowInColumn(WS, 4) <> 0 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Multiple cell Start 1, 1 Hidden Rows And Columns Outside more"
                End If
            WS.Rows.Hidden = False
            WS.Columns.Hidden = False
            'Filter
            WS.Range("A1").Clear
            WS.Range("$A$1:$C$4").AutoFilter Field:=1, Criteria1:="$A$3"
                If GetFirstRowInColumn(WS, 1) <> 2 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Multiple cell Start 1, 1 Filter"
                End If
                If GetFirstRowInColumn(WS, 1, True, True) <> 2 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Multiple cell Start 1, 1 Filter True True"
                End If
                If GetFirstRowInColumn(WS, 1, True, False) <> 2 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Multiple cell Start 1, 1 Filter True False"
                End If
                If GetFirstRowInColumn(WS, 1, False, True) <> 3 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Multiple cell Start 1, 1 Filter False True"
                End If
                If GetFirstRowInColumn(WS, 1, False, False) <> 3 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Multiple cell Start 1, 1 Filter False False"
                End If
                'Outside
                If GetFirstRowInColumn(WS, 4) <> 0 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Multiple cell Start 1, 1 Filter Outside more"
                End If
            WS.AutoFilterMode = False
                
        'Start 2, 2
        WS.Cells.Clear
        WS.Range("B2:D2").Value = "Header"
        WS.Range("B3:D5").Formula = "=ADDRESS(ROW(),COLUMN())"
            'Unhidden
                If GetFirstRowInColumn(WS, 2) <> 2 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Multiple cell Start 2, 2 Unhidden"
                End If
                If GetFirstRowInColumn(WS, 2, True, True) <> 2 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Multiple cell Start 2, 2 Unhidden True True"
                End If
                If GetFirstRowInColumn(WS, 2, True, False) <> 2 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Multiple cell Start 2, 2 Unhidden True False"
                End If
                If GetFirstRowInColumn(WS, 2, False, True) <> 2 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Multiple cell Start 2, 2 Unhidden False True"
                End If
                If GetFirstRowInColumn(WS, 2, False, False) <> 2 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Multiple cell Start 2, 2 Unhidden False False"
                End If
                'Outside
                If GetFirstRowInColumn(WS, 1) <> 0 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Multiple cell Start 2, 2 Unhidden Outside less"
                End If
                If GetFirstRowInColumn(WS, 5) <> 0 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Multiple cell Start 2, 2 Unhidden Outside more"
                End If
            'Hidden Rows
            WS.Rows(2).Hidden = True
                If GetFirstRowInColumn(WS, 2) <> 2 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Multiple cell Start 2, 2 Hidden Rows"
                End If
                If GetFirstRowInColumn(WS, 2, True, True) <> 2 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Multiple cell Start 2, 2 Hidden Rows True True"
                End If
                If GetFirstRowInColumn(WS, 2, True, False) <> 2 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Multiple cell Start 2, 2 Hidden Rows True False"
                End If
                If GetFirstRowInColumn(WS, 2, False, True) <> 3 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Multiple cell Start 2, 2 Hidden Rows False True"
                End If
                If GetFirstRowInColumn(WS, 2, False, False) <> 3 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Multiple cell Start 2, 2 Hidden Rows False False"
                End If
                'Outside
                If GetFirstRowInColumn(WS, 1) <> 0 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Multiple cell Start 2, 2 Hidden Rows Outside less"
                End If
                If GetFirstRowInColumn(WS, 5) <> 0 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Multiple cell Start 2, 2 Hidden  Rows Outside more"
                End If
            WS.Rows.Hidden = False
            'Hidden Columns
            WS.Columns(2).Hidden = True
                If GetFirstRowInColumn(WS, 2) <> 2 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Multiple cell Start 2, 2 Hidden Columns"
                End If
                If GetFirstRowInColumn(WS, 2, True, True) <> 2 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Multiple cell Start 2, 2 Hidden Columns True True"
                End If
                If GetFirstRowInColumn(WS, 2, True, False) <> 0 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Multiple cell Start 2, 2 Hidden Columns True False"
                End If
                If GetFirstRowInColumn(WS, 2, False, True) <> 2 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Multiple cell Start 2, 2 Hidden Columns False True"
                End If
                If GetFirstRowInColumn(WS, 2, False, False) <> 0 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Multiple cell Start 2, 2 Hidden Columns False False"
                End If
                'Outside
                If GetFirstRowInColumn(WS, 1) <> 0 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Multiple cell Start 2, 2 Hidden Columns Outside less"
                End If
                If GetFirstRowInColumn(WS, 5) <> 0 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Multiple cell Start 2, 2 Hidden  Columns Outside more"
                End If
            WS.Columns.Hidden = False
            'Hidden Rows And Columns
            WS.Rows(2).Hidden = True
            WS.Columns(2).Hidden = True
                If GetFirstRowInColumn(WS, 2) <> 2 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Multiple cell Start 2, 2 Hidden Rows And Columns"
                End If
                If GetFirstRowInColumn(WS, 2, True, True) <> 2 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Multiple cell Start 2, 2 Hidden Rows And Columns True True"
                End If
                If GetFirstRowInColumn(WS, 2, True, False) <> 0 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Multiple cell Start 2, 2 Hidden Rows And Columns True False"
                End If
                If GetFirstRowInColumn(WS, 2, False, True) <> 3 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Multiple cell Start 2, 2 Hidden Rows And Columns False True"
                End If
                If GetFirstRowInColumn(WS, 2, False, False) <> 0 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Multiple cell Start 2, 2 Hidden Rows And Columns False False"
                End If
                'Outside
                If GetFirstRowInColumn(WS, 1) <> 0 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Multiple cell Start 2, 2 Hidden Rows And Columns Outside less"
                End If
                If GetFirstRowInColumn(WS, 5) <> 0 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Multiple cell Start 2, 2 Hidden Rows And Columns Outside more"
                End If
            WS.Rows.Hidden = False
            WS.Columns.Hidden = False
            'Filter
            WS.Range("B2").Clear
            WS.Range("$B$2:$D$5").AutoFilter Field:=1, Criteria1:="$B$4"
                If GetFirstRowInColumn(WS, 2) <> 3 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Multiple cell Start 2, 2 Filter"
                End If
                If GetFirstRowInColumn(WS, 2, True, True) <> 3 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Multiple cell Start 2, 2 Filter True True"
                End If
                If GetFirstRowInColumn(WS, 2, True, False) <> 3 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Multiple cell Start 2, 2 Filter True False"
                End If
                If GetFirstRowInColumn(WS, 2, False, True) <> 4 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Multiple cell Start 2, 2 Filter False True"
                End If
                If GetFirstRowInColumn(WS, 2, False, False) <> 4 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Multiple cell Start 2, 2 Filter False False"
                End If
                'Outside
                If GetFirstRowInColumn(WS, 1) <> 0 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Multiple cell Start 2, 2 Filter Outside less"
                End If
                If GetFirstRowInColumn(WS, 5) <> 0 Then
                    TestGetFirstRowInColumn = False
                    Debug.Print "Multiple cell Start 2, 2 Filter Outside more"
                End If
            WS.AutoFilterMode = False
       
    'Dynamic
    WS.Cells.Clear
    Dim i&
    For i = 5 To 1 Step -1
        WS.Cells(i, i).Value = i
        If GetFirstRowInColumn(WS, i) <> i Then
            TestGetFirstRowInColumn = False
            Debug.Print "Dynamic"
        End If
    Next i
    
    Excel.Application.DisplayAlerts = False
    WB.Close
    Excel.Application.DisplayAlerts = True
    
    Excel.Application.ScreenUpdating = True
    
    Debug.Print "TestGetFirstRowInColumn: " & TestGetFirstRowInColumn
    
End Function

Private Function TestGetFirstColumn() As Boolean
    
    TestGetFirstColumn = True

    Excel.Application.ScreenUpdating = False
    
    Dim WB As Workbook
    Set WB = Excel.Application.Workbooks.Add
    
    Dim WS As Worksheet
    Set WS = WB.Worksheets(1)
    
    'Blank
        'Unhidden
            If GetFirstColumn(WS) <> 0 Then
                TestGetFirstColumn = False
                Debug.Print "Blank Unhidden"
            End If
            If GetFirstColumn(WS, True, True) <> 0 Then
                TestGetFirstColumn = False
                Debug.Print "Blank Unhidden True True"
            End If
            If GetFirstColumn(WS, True, False) <> 0 Then
                TestGetFirstColumn = False
                Debug.Print "Blank Unhidden True False"
            End If
            If GetFirstColumn(WS, False, True) <> 0 Then
                TestGetFirstColumn = False
                Debug.Print "Blank Unhidden False True"
            End If
            If GetFirstColumn(WS, False, False) <> 0 Then
                TestGetFirstColumn = False
                Debug.Print "Blank Unhidden False False"
            End If
        'Hidden Row
        WS.Rows(1).Hidden = True
            If GetFirstColumn(WS) <> 0 Then
                TestGetFirstColumn = False
                Debug.Print "Blank Hidden Row"
            End If
            If GetFirstColumn(WS, True, True) <> 0 Then
                TestGetFirstColumn = False
                Debug.Print "Blank Hidden Row True True"
            End If
            If GetFirstColumn(WS, True, False) <> 0 Then
                TestGetFirstColumn = False
                Debug.Print "Blank Hidden Row True False"
            End If
            If GetFirstColumn(WS, False, True) <> 0 Then
                TestGetFirstColumn = False
                Debug.Print "Blank Hidden Row False True"
            End If
            If GetFirstColumn(WS, False, False) <> 0 Then
                TestGetFirstColumn = False
                Debug.Print "Blank Hidden Row False False"
            End If
        WS.Rows.Hidden = False
        'Hidden Column
        WS.Columns(1).Hidden = True
            If GetFirstColumn(WS) <> 0 Then
                TestGetFirstColumn = False
                Debug.Print "Blank Hidden Column"
            End If
            If GetFirstColumn(WS, True, True) <> 0 Then
                TestGetFirstColumn = False
                Debug.Print "Blank Hidden Column True True"
            End If
            If GetFirstColumn(WS, True, False) <> 0 Then
                TestGetFirstColumn = False
                Debug.Print "Blank Hidden Column True False"
            End If
            If GetFirstColumn(WS, False, True) <> 0 Then
                TestGetFirstColumn = False
                Debug.Print "Blank Hidden Column False True"
            End If
            If GetFirstColumn(WS, False, False) <> 0 Then
                TestGetFirstColumn = False
                Debug.Print "Blank Hidden Column False False"
            End If
        WS.Columns.Hidden = False
        'Hidden Row And Column
        WS.Columns(1).Hidden = True
        WS.Rows(1).Hidden = True
            If GetFirstColumn(WS) <> 0 Then
                TestGetFirstColumn = False
                Debug.Print "Blank Hidden Row And Column"
            End If
            If GetFirstColumn(WS, True, True) <> 0 Then
                TestGetFirstColumn = False
                Debug.Print "Blank Hidden Row And Column True True"
            End If
            If GetFirstColumn(WS, True, False) <> 0 Then
                TestGetFirstColumn = False
                Debug.Print "Blank Hidden Row And Column True False"
            End If
            If GetFirstColumn(WS, False, True) <> 0 Then
                TestGetFirstColumn = False
                Debug.Print "Blank Hidden Row And Column False True"
            End If
            If GetFirstColumn(WS, False, False) <> 0 Then
                TestGetFirstColumn = False
                Debug.Print "Blank Hidden Row And Column False False"
            End If
        WS.Rows.Hidden = False
        WS.Columns.Hidden = False
        'Filter
            'Cannot filter blank
            
    'Single cell
        'Start 1, 1
        WS.Cells(1, 1).Value = "Test"
            'Unhidden
                If GetFirstColumn(WS) <> 1 Then
                    TestGetFirstColumn = False
                    Debug.Print "Single cell Start 1, 1 Unhidden"
                End If
                If GetFirstColumn(WS, True, True) <> 1 Then
                    TestGetFirstColumn = False
                    Debug.Print "Single cell Start 1, 1 Unhidden True True"
                End If
                If GetFirstColumn(WS, True, False) <> 1 Then
                    TestGetFirstColumn = False
                    Debug.Print "Single cell Start 1, 1 Unhidden True False"
                End If
                If GetFirstColumn(WS, False, True) <> 1 Then
                    TestGetFirstColumn = False
                    Debug.Print "Single cell Start 1, 1 Unhidden False True"
                End If
                If GetFirstColumn(WS, False, False) <> 1 Then
                    TestGetFirstColumn = False
                    Debug.Print "Single cell Start 1, 1 Unhidden False False"
                End If
            'Hidden Row
            WS.Rows(1).Hidden = True
                If GetFirstColumn(WS) <> 1 Then
                    TestGetFirstColumn = False
                    Debug.Print "Single cell Start 1, 1 Hidden Row"
                End If
                If GetFirstColumn(WS, True, True) <> 1 Then
                    TestGetFirstColumn = False
                    Debug.Print "Single cell Start 1, 1 Hidden Row True True"
                End If
                If GetFirstColumn(WS, True, False) <> 1 Then
                    TestGetFirstColumn = False
                    Debug.Print "Single cell Start 1, 1 Hidden Row True False"
                End If
                If GetFirstColumn(WS, False, True) <> 0 Then
                    TestGetFirstColumn = False
                    Debug.Print "Single cell Start 1, 1 Hidden Row False True"
                End If
                If GetFirstColumn(WS, False, False) <> 0 Then
                    TestGetFirstColumn = False
                    Debug.Print "Single cell Start 1, 1 Hidden Row False False"
                End If
            WS.Rows.Hidden = False
            'Hidden Column
            WS.Columns(1).Hidden = True
                If GetFirstColumn(WS) <> 1 Then
                    TestGetFirstColumn = False
                    Debug.Print "Single cell Start 1, 1 Hidden Column"
                End If
                If GetFirstColumn(WS, True, True) <> 1 Then
                    TestGetFirstColumn = False
                    Debug.Print "Single cell Start 1, 1 Hidden Column True True"
                End If
                If GetFirstColumn(WS, True, False) <> 0 Then
                    TestGetFirstColumn = False
                    Debug.Print "Single cell Start 1, 1 Hidden Column True False"
                End If
                If GetFirstColumn(WS, False, True) <> 1 Then
                    TestGetFirstColumn = False
                    Debug.Print "Single cell Start 1, 1 Hidden Column False True"
                End If
                If GetFirstColumn(WS, False, False) <> 0 Then
                    TestGetFirstColumn = False
                    Debug.Print "Single cell Start 1, 1 Hidden Column False False"
                End If
            WS.Columns.Hidden = False
            'Hidden Row And Column
            WS.Columns(1).Hidden = True
            WS.Rows(1).Hidden = True
                If GetFirstColumn(WS) <> 1 Then
                    TestGetFirstColumn = False
                    Debug.Print "Single cell Start 1, 1 Hidden Row And Column"
                End If
                If GetFirstColumn(WS, True, True) <> 1 Then
                    TestGetFirstColumn = False
                    Debug.Print "Single cell Start 1, 1 Hidden Row And Column True True"
                End If
                If GetFirstColumn(WS, True, False) <> 0 Then
                    TestGetFirstColumn = False
                    Debug.Print "Single cell Start 1, 1 Hidden Row And Column True False"
                End If
                If GetFirstColumn(WS, False, True) <> 0 Then
                    TestGetFirstColumn = False
                    Debug.Print "Single cell Start 1, 1 Hidden Row And Column False True"
                End If
                If GetFirstColumn(WS, False, False) <> 0 Then
                    TestGetFirstColumn = False
                    Debug.Print "Single cell Start 1, 1 Hidden Row And Column False False"
                End If
            WS.Rows.Hidden = False
            WS.Columns.Hidden = False
            'Filter
                'Cannot filter single cell
        'Start 2, 2
        WS.Cells.Clear
        WS.Cells(2, 2).Value = "Test"
            'Unhidden
                If GetFirstColumn(WS) <> 2 Then
                    TestGetFirstColumn = False
                    Debug.Print "Single cell Start 2, 2 Unhidden"
                End If
                If GetFirstColumn(WS, True, True) <> 2 Then
                    TestGetFirstColumn = False
                    Debug.Print "Single cell Start 2, 2 Unhidden True True"
                End If
                If GetFirstColumn(WS, True, False) <> 2 Then
                    TestGetFirstColumn = False
                    Debug.Print "Single cell Start 2, 2 Unhidden True False"
                End If
                If GetFirstColumn(WS, False, True) <> 2 Then
                    TestGetFirstColumn = False
                    Debug.Print "Single cell Start 2, 2 Unhidden False True"
                End If
                If GetFirstColumn(WS, False, False) <> 2 Then
                    TestGetFirstColumn = False
                    Debug.Print "Single cell Start 2, 2 Unhidden False False"
                End If
            'Hidden Row
            WS.Rows(2).Hidden = True
                If GetFirstColumn(WS) <> 2 Then
                    TestGetFirstColumn = False
                    Debug.Print "Single cell Start 2, 2 Hidden Row"
                End If
                If GetFirstColumn(WS, True, True) <> 2 Then
                    TestGetFirstColumn = False
                    Debug.Print "Single cell Start 2, 2 Hidden Row True True"
                End If
                If GetFirstColumn(WS, True, False) <> 2 Then
                    TestGetFirstColumn = False
                    Debug.Print "Single cell Start 2, 2 Hidden Row True False"
                End If
                If GetFirstColumn(WS, False, True) <> 0 Then
                    TestGetFirstColumn = False
                    Debug.Print "Single cell Start 2, 2 Hidden Row False True"
                End If
                If GetFirstColumn(WS, False, False) <> 0 Then
                    TestGetFirstColumn = False
                    Debug.Print "Single cell Start 2, 2 Hidden Row False False"
                End If
            WS.Rows.Hidden = False
            'Hidden Column
            WS.Columns(2).Hidden = True
                If GetFirstColumn(WS) <> 2 Then
                    TestGetFirstColumn = False
                    Debug.Print "Single cell Start 2, 2 Hidden Column"
                End If
                If GetFirstColumn(WS, True, True) <> 2 Then
                    TestGetFirstColumn = False
                    Debug.Print "Single cell Start 2, 2 Hidden Column True True"
                End If
                If GetFirstColumn(WS, True, False) <> 0 Then
                    TestGetFirstColumn = False
                    Debug.Print "Single cell Start 2, 2 Hidden Column True False"
                End If
                If GetFirstColumn(WS, False, True) <> 2 Then
                    TestGetFirstColumn = False
                    Debug.Print "Single cell Start 2, 2 Hidden Column False True"
                End If
                If GetFirstColumn(WS, False, False) <> 0 Then
                    TestGetFirstColumn = False
                    Debug.Print "Single cell Start 2, 2 Hidden Column False False"
                End If
            WS.Columns.Hidden = False
            'Hidden Row And Column
            WS.Columns(2).Hidden = True
            WS.Rows(2).Hidden = True
                If GetFirstColumn(WS) <> 2 Then
                    TestGetFirstColumn = False
                    Debug.Print "Single cell Start 2, 2 Hidden Row And Column"
                End If
                If GetFirstColumn(WS, True, True) <> 2 Then
                    TestGetFirstColumn = False
                    Debug.Print "Single cell Start 2, 2 Hidden Row And Column True True"
                End If
                If GetFirstColumn(WS, True, False) <> 0 Then
                    TestGetFirstColumn = False
                    Debug.Print "Single cell Start 2, 2 Hidden Row And Column True False"
                End If
                If GetFirstColumn(WS, False, True) <> 0 Then
                    TestGetFirstColumn = False
                    Debug.Print "Single cell Start 2, 2 Hidden Row And Column False True"
                End If
                If GetFirstColumn(WS, False, False) <> 0 Then
                    TestGetFirstColumn = False
                    Debug.Print "Single cell Start 2, 2 Hidden Row And Column False False"
                End If
            WS.Rows.Hidden = False
            WS.Columns.Hidden = False
            'Filter
                'Cannot filter single cell
                
    'Muliple cell
        'Start 1, 1
        WS.Cells.Clear
        WS.Range("A1:C1").Value = "Header"
        WS.Range("A2:C4").Formula = "=ADDRESS(ROW(),COLUMN())"
            'Unhidden
                If GetFirstColumn(WS) <> 1 Then
                    TestGetFirstColumn = False
                    Debug.Print "Muliple cell Start 1, 1 Unhidden"
                End If
                If GetFirstColumn(WS, True, True) <> 1 Then
                    TestGetFirstColumn = False
                    Debug.Print "Muliple cell Start 1, 1 Unhidden True True"
                End If
                If GetFirstColumn(WS, True, False) <> 1 Then
                    TestGetFirstColumn = False
                    Debug.Print "Muliple cell Start 1, 1 Unhidden True False"
                End If
                If GetFirstColumn(WS, False, True) <> 1 Then
                    TestGetFirstColumn = False
                    Debug.Print "Muliple cell Start 1, 1 Unhidden False True"
                End If
                If GetFirstColumn(WS, False, False) <> 1 Then
                    TestGetFirstColumn = False
                    Debug.Print "Muliple cell Start 1, 1 Unhidden False False"
                End If
            'Hidden Row
            WS.Rows(1).Hidden = True
                If GetFirstColumn(WS) <> 1 Then
                    TestGetFirstColumn = False
                    Debug.Print "Muliple cell Start 1, 1 Hidden Row"
                End If
                If GetFirstColumn(WS, True, True) <> 1 Then
                    TestGetFirstColumn = False
                    Debug.Print "Muliple cell Start 1, 1 Hidden Row True True"
                End If
                If GetFirstColumn(WS, True, False) <> 1 Then
                    TestGetFirstColumn = False
                    Debug.Print "Muliple cell Start 1, 1 Hidden Row True False"
                End If
                If GetFirstColumn(WS, False, True) <> 1 Then
                    TestGetFirstColumn = False
                    Debug.Print "Muliple cell Start 1, 1 Hidden Row False True"
                End If
                If GetFirstColumn(WS, False, False) <> 1 Then
                    TestGetFirstColumn = False
                    Debug.Print "Muliple cell Start 1, 1 Hidden Row False False"
                End If
            WS.Rows.Hidden = False
            'Hidden Column
            WS.Columns(1).Hidden = True
                If GetFirstColumn(WS) <> 1 Then
                    TestGetFirstColumn = False
                    Debug.Print "Muliple cell Start 1, 1 Hidden Column"
                End If
                If GetFirstColumn(WS, True, True) <> 1 Then
                    TestGetFirstColumn = False
                    Debug.Print "Muliple cell Start 1, 1 Hidden Column True True"
                End If
                If GetFirstColumn(WS, True, False) <> 2 Then
                    TestGetFirstColumn = False
                    Debug.Print "Muliple cell Start 1, 1 Hidden Column True False"
                End If
                If GetFirstColumn(WS, False, True) <> 1 Then
                    TestGetFirstColumn = False
                    Debug.Print "Muliple cell Start 1, 1 Hidden Column False True"
                End If
                If GetFirstColumn(WS, False, False) <> 2 Then
                    TestGetFirstColumn = False
                    Debug.Print "Muliple cell Start 1, 1 Hidden Column False False"
                End If
            WS.Columns.Hidden = False
            'Hidden Row And Column
            WS.Columns(1).Hidden = True
            WS.Rows(1).Hidden = True
                If GetFirstColumn(WS) <> 1 Then
                    TestGetFirstColumn = False
                    Debug.Print "Muliple cell Start 1, 1 Hidden Row And Column"
                End If
                If GetFirstColumn(WS, True, True) <> 1 Then
                    TestGetFirstColumn = False
                    Debug.Print "Muliple cell Start 1, 1 Hidden Row And Column True True"
                End If
                If GetFirstColumn(WS, True, False) <> 2 Then
                    TestGetFirstColumn = False
                    Debug.Print "Muliple cell Start 1, 1 Hidden Row And Column True False"
                End If
                If GetFirstColumn(WS, False, True) <> 1 Then
                    TestGetFirstColumn = False
                    Debug.Print "Muliple cell Start 1, 1 Hidden Row And Column False True"
                End If
                If GetFirstColumn(WS, False, False) <> 2 Then
                    TestGetFirstColumn = False
                    Debug.Print "Muliple cell Start 1, 1 Hidden Row And Column False False"
                End If
            WS.Rows.Hidden = False
            WS.Columns.Hidden = False
            'Filter
            WS.Range("A1,A3:A4").Clear
            WS.Range("$A$1:$C$4").AutoFilter Field:=1, Criteria1:="$A$3"
                If GetFirstColumn(WS) <> 1 Then
                    TestGetFirstColumn = False
                    Debug.Print "Muliple cell Start 1, 1 Filter"
                End If
                If GetFirstColumn(WS, True, True) <> 1 Then
                    TestGetFirstColumn = False
                    Debug.Print "Muliple cell Start 1, 1 Filter True True"
                End If
                If GetFirstColumn(WS, True, False) <> 1 Then
                    TestGetFirstColumn = False
                    Debug.Print "Muliple cell Start 1, 1 Filter True False"
                End If
                If GetFirstColumn(WS, False, True) <> 2 Then
                    TestGetFirstColumn = False
                    Debug.Print "Muliple cell Start 1, 1 Filter False True"
                End If
                If GetFirstColumn(WS, False, False) <> 2 Then
                    TestGetFirstColumn = False
                    Debug.Print "Muliple cell Start 1, 1 Filter False False"
                End If
            WS.AutoFilterMode = False
        'Start 2, 2
        WS.Cells.Clear
        WS.Range("B2:D2").Value = "Header"
        WS.Range("B3:D5").Formula = "=ADDRESS(ROW(),COLUMN())"
            'Unhidden
                If GetFirstColumn(WS) <> 2 Then
                    TestGetFirstColumn = False
                    Debug.Print "Muliple cell Start 2, 2 Unhidden"
                End If
                If GetFirstColumn(WS, True, True) <> 2 Then
                    TestGetFirstColumn = False
                    Debug.Print "Muliple cell Start 2, 2 Unhidden True True"
                End If
                If GetFirstColumn(WS, True, False) <> 2 Then
                    TestGetFirstColumn = False
                    Debug.Print "Muliple cell Start 2, 2 Unhidden True False"
                End If
                If GetFirstColumn(WS, False, True) <> 2 Then
                    TestGetFirstColumn = False
                    Debug.Print "Muliple cell Start 2, 2 Unhidden False True"
                End If
                If GetFirstColumn(WS, False, False) <> 2 Then
                    TestGetFirstColumn = False
                    Debug.Print "Muliple cell Start 2, 2 Unhidden False False"
                End If
            'Hidden Row
            WS.Rows(2).Hidden = True
                If GetFirstColumn(WS) <> 2 Then
                    TestGetFirstColumn = False
                    Debug.Print "Muliple cell Start 2, 2 Hidden Row"
                End If
                If GetFirstColumn(WS, True, True) <> 2 Then
                    TestGetFirstColumn = False
                    Debug.Print "Muliple cell Start 2, 2 Hidden Row True True"
                End If
                If GetFirstColumn(WS, True, False) <> 2 Then
                    TestGetFirstColumn = False
                    Debug.Print "Muliple cell Start 2, 2 Hidden Row True False"
                End If
                If GetFirstColumn(WS, False, True) <> 2 Then
                    TestGetFirstColumn = False
                    Debug.Print "Muliple cell Start 2, 2 Hidden Row False True"
                End If
                If GetFirstColumn(WS, False, False) <> 2 Then
                    TestGetFirstColumn = False
                    Debug.Print "Muliple cell Start 2, 2 Hidden Row False False"
                End If
            WS.Rows.Hidden = False
            'Hidden Column
            WS.Columns(2).Hidden = True
                If GetFirstColumn(WS) <> 2 Then
                    TestGetFirstColumn = False
                    Debug.Print "Muliple cell Start 2, 2 Hidden Column"
                End If
                If GetFirstColumn(WS, True, True) <> 2 Then
                    TestGetFirstColumn = False
                    Debug.Print "Muliple cell Start 2, 2 Hidden Column True True"
                End If
                If GetFirstColumn(WS, True, False) <> 3 Then
                    TestGetFirstColumn = False
                    Debug.Print "Muliple cell Start 2, 2 Hidden Column True False"
                End If
                If GetFirstColumn(WS, False, True) <> 2 Then
                    TestGetFirstColumn = False
                    Debug.Print "Muliple cell Start 2, 2 Hidden Column False True"
                End If
                If GetFirstColumn(WS, False, False) <> 3 Then
                    TestGetFirstColumn = False
                    Debug.Print "Muliple cell Start 2, 2 Hidden Column False False"
                End If
            WS.Columns.Hidden = False
            'Hidden Row And Column
            WS.Columns(2).Hidden = True
            WS.Rows(2).Hidden = True
                If GetFirstColumn(WS) <> 2 Then
                    TestGetFirstColumn = False
                    Debug.Print "Muliple cell Start 2, 2 Hidden Row And Column"
                End If
                If GetFirstColumn(WS, True, True) <> 2 Then
                    TestGetFirstColumn = False
                    Debug.Print "Muliple cell Start 2, 2 Hidden Row And Column True True"
                End If
                If GetFirstColumn(WS, True, False) <> 3 Then
                    TestGetFirstColumn = False
                    Debug.Print "Muliple cell Start 2, 2 Hidden Row And Column True False"
                End If
                If GetFirstColumn(WS, False, True) <> 2 Then
                    TestGetFirstColumn = False
                    Debug.Print "Muliple cell Start 2, 2 Hidden Row And Column False True"
                End If
                If GetFirstColumn(WS, False, False) <> 3 Then
                    TestGetFirstColumn = False
                    Debug.Print "Muliple cell Start 2, 2 Hidden Row And Column False False"
                End If
            WS.Rows.Hidden = False
            WS.Columns.Hidden = False
            'Filter
            WS.Range("B2,B4:B5").Clear
            WS.Range("$B$2:$D$5").AutoFilter Field:=2, Criteria1:="$C$4"
                If GetFirstColumn(WS) <> 2 Then
                    TestGetFirstColumn = False
                    Debug.Print "Muliple cell Start 2, 2 Filter"
                End If
                If GetFirstColumn(WS, True, True) <> 2 Then
                    TestGetFirstColumn = False
                    Debug.Print "Muliple cell Start 2, 2 Filter True True"
                End If
                If GetFirstColumn(WS, True, False) <> 2 Then
                    TestGetFirstColumn = False
                    Debug.Print "Muliple cell Start 2, 2 Filter True False"
                End If
                If GetFirstColumn(WS, False, True) <> 3 Then
                    TestGetFirstColumn = False
                    Debug.Print "Muliple cell Start 2, 2 Filter False True"
                End If
                If GetFirstColumn(WS, False, False) <> 3 Then
                    TestGetFirstColumn = False
                    Debug.Print "Muliple cell Start 2, 2 Filter False False"
                End If
            WS.AutoFilterMode = False
    
    'Dynamic
    WS.Cells.Clear
    Dim i&
    For i = 5 To 1 Step -1
        WS.Cells(i, i).Value = i
        If GetFirstColumn(WS) <> i Then
            TestGetFirstColumn = False
            Debug.Print "Dynamic"
        End If
    Next i
    
    Excel.Application.DisplayAlerts = False
    WB.Close
    Excel.Application.DisplayAlerts = True
    
    Excel.Application.ScreenUpdating = True
    
    Debug.Print "TestGetFirstColumn: " & TestGetFirstColumn
    
End Function

Private Function TestGetFirstColumnInRow() As Boolean
    
    TestGetFirstColumnInRow = True

    Excel.Application.ScreenUpdating = False
    
    Dim WB As Workbook
    Set WB = Excel.Application.Workbooks.Add
    
    Dim WS As Worksheet
    Set WS = WB.Worksheets(1)
            
    'Blank
        'Unhidden
            If GetFirstColumnInRow(WS, 1) <> 0 Then
                TestGetFirstColumnInRow = False
                Debug.Print "Blank Unhidden"
            End If
            If GetFirstColumnInRow(WS, 1, True, True) <> 0 Then
                TestGetFirstColumnInRow = False
                Debug.Print "Blank Unhidden True True"
            End If
            If GetFirstColumnInRow(WS, 1, True, False) <> 0 Then
                TestGetFirstColumnInRow = False
                Debug.Print "Blank Unhidden True False"
            End If
            If GetFirstColumnInRow(WS, 1, False, True) <> 0 Then
                TestGetFirstColumnInRow = False
                Debug.Print "Blank Unhidden False True"
            End If
            If GetFirstColumnInRow(WS, 1, False, False) <> 0 Then
                TestGetFirstColumnInRow = False
                Debug.Print "Blank Unhidden False False"
            End If
        'Hidden Rows
        WS.Rows(1).Hidden = True
            If GetFirstColumnInRow(WS, 1) <> 0 Then
                TestGetFirstColumnInRow = False
                Debug.Print "Blank Hidden Rows"
            End If
            If GetFirstColumnInRow(WS, 1, True, True) <> 0 Then
                TestGetFirstColumnInRow = False
                Debug.Print "Blank Hidden Rows True True"
            End If
            If GetFirstColumnInRow(WS, 1, True, False) <> 0 Then
                TestGetFirstColumnInRow = False
                Debug.Print "Blank Hidden Rows True False"
            End If
            If GetFirstColumnInRow(WS, 1, False, True) <> 0 Then
                TestGetFirstColumnInRow = False
                Debug.Print "Blank Hidden Rows False True"
            End If
            If GetFirstColumnInRow(WS, 1, False, False) <> 0 Then
                TestGetFirstColumnInRow = False
                Debug.Print "Blank Hidden Rows False False"
            End If
        WS.Rows.Hidden = False
        'Hidden Columns
        WS.Columns(1).Hidden = True
            If GetFirstColumnInRow(WS, 1) <> 0 Then
                TestGetFirstColumnInRow = False
                Debug.Print "Blank Hidden Columns"
            End If
            If GetFirstColumnInRow(WS, 1, True, True) <> 0 Then
                TestGetFirstColumnInRow = False
                Debug.Print "Blank Hidden Columns True True"
            End If
            If GetFirstColumnInRow(WS, 1, True, False) <> 0 Then
                TestGetFirstColumnInRow = False
                Debug.Print "Blank Hidden Columns True False"
            End If
            If GetFirstColumnInRow(WS, 1, False, True) <> 0 Then
                TestGetFirstColumnInRow = False
                Debug.Print "Blank Hidden Columns False True"
            End If
            If GetFirstColumnInRow(WS, 1, False, False) <> 0 Then
                TestGetFirstColumnInRow = False
                Debug.Print "Blank Hidden Columns False False"
            End If
        WS.Columns.Hidden = False
        'Hidden Rows And Columns
        WS.Rows(1).Hidden = True
        WS.Columns(1).Hidden = True
            If GetFirstColumnInRow(WS, 1) <> 0 Then
                TestGetFirstColumnInRow = False
                Debug.Print "Blank Hidden Rows And Columns"
            End If
            If GetFirstColumnInRow(WS, 1, True, True) <> 0 Then
                TestGetFirstColumnInRow = False
                Debug.Print "Blank Hidden Rows And Columns True True"
            End If
            If GetFirstColumnInRow(WS, 1, True, False) <> 0 Then
                TestGetFirstColumnInRow = False
                Debug.Print "Blank Hidden Rows And Columns True False"
            End If
            If GetFirstColumnInRow(WS, 1, False, True) <> 0 Then
                TestGetFirstColumnInRow = False
                Debug.Print "Blank Hidden Rows And Columns False True"
            End If
            If GetFirstColumnInRow(WS, 1, False, False) <> 0 Then
                TestGetFirstColumnInRow = False
                Debug.Print "Blank Hidden Rows And Columns False False"
            End If
        WS.Rows.Hidden = False
        WS.Columns.Hidden = False
        'Filter
            'Cannot filter blank
            
    'Single cell
        'Start 1, 1
        WS.Cells(1, 1).Value = "Test"
            'Unhidden
                If GetFirstColumnInRow(WS, 1) <> 1 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Single cell Start 1, 1 Unhidden"
                End If
                If GetFirstColumnInRow(WS, 1, True, True) <> 1 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Single cell Start 1, 1 Unhidden True True"
                End If
                If GetFirstColumnInRow(WS, 1, True, False) <> 1 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Single cell Start 1, 1 Unhidden True False"
                End If
                If GetFirstColumnInRow(WS, 1, False, True) <> 1 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Single cell Start 1, 1 Unhidden False True"
                End If
                If GetFirstColumnInRow(WS, 1, False, False) <> 1 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Single cell Start 1, 1 Unhidden False False"
                End If
                'Outside
                If GetFirstColumnInRow(WS, 4) <> 0 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Single cell Start 1, 1 Unhidden Outside more"
                End If
            'Hidden Rows
            WS.Rows(1).Hidden = True
                If GetFirstColumnInRow(WS, 1) <> 1 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Single cell Start 1, 1 Hidden Rows"
                End If
                If GetFirstColumnInRow(WS, 1, True, True) <> 1 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Single cell Start 1, 1 Hidden Rows True True"
                End If
                If GetFirstColumnInRow(WS, 1, True, False) <> 1 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Single cell Start 1, 1 Hidden Rows True False"
                End If
                If GetFirstColumnInRow(WS, 1, False, True) <> 0 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Single cell Start 1, 1 Hidden Rows False True"
                End If
                If GetFirstColumnInRow(WS, 1, False, False) <> 0 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Single cell Start 1, 1 Hidden Rows False False"
                End If
                'Outside
                If GetFirstColumnInRow(WS, 2) <> 0 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Single cell Start 1, 1 Hidden Rows Outside more"
                End If
            WS.Rows.Hidden = False
            'Hidden Columns
            WS.Columns(1).Hidden = True
                If GetFirstColumnInRow(WS, 1) <> 1 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Single cell Start 1, 1 Hidden Columns"
                End If
                If GetFirstColumnInRow(WS, 1, True, True) <> 1 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Single cell Start 1, 1 Hidden Columns True True"
                End If
                If GetFirstColumnInRow(WS, 1, True, False) <> 0 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Single cell Start 1, 1 Hidden Columns True False"
                End If
                If GetFirstColumnInRow(WS, 1, False, True) <> 1 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Single cell Start 1, 1 Hidden Columns False True"
                End If
                If GetFirstColumnInRow(WS, 1, False, False) <> 0 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Single cell Start 1, 1 Hidden Columns False False"
                End If
                'Outside
                If GetFirstColumnInRow(WS, 2) <> 0 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Single cell Start 1, 1 Hidden Columns Outside more"
                End If
            WS.Columns.Hidden = False
            'Hidden Rows And Columns
            WS.Rows(1).Hidden = True
            WS.Columns(1).Hidden = True
                If GetFirstColumnInRow(WS, 1) <> 1 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Single cell Start 1, 1 Hidden Rows And Columns"
                End If
                If GetFirstColumnInRow(WS, 1, True, True) <> 1 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Single cell Start 1, 1 Hidden Rows And Columns True True"
                End If
                If GetFirstColumnInRow(WS, 1, True, False) <> 0 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Single cell Start 1, 1 Hidden Rows And Columns True False"
                End If
                If GetFirstColumnInRow(WS, 1, False, True) <> 0 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Single cell Start 1, 1 Hidden Rows And Columns False True"
                End If
                If GetFirstColumnInRow(WS, 1, False, False) <> 0 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Single cell Start 1, 1 Hidden Rows And Columns False False"
                End If
                'Outside
                If GetFirstColumnInRow(WS, 2) <> 0 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Single cell Start 1, 1 Hidden Rows And Columns Outside more"
                End If
            WS.Rows.Hidden = False
            WS.Columns.Hidden = False
            'Filter
                'Cannot filter single cell

        'Start 2, 2
        WS.Cells.Clear
        WS.Cells(2, 2).Value = "Test"
            'Unhidden
                If GetFirstColumnInRow(WS, 2) <> 2 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Single cell Start 2, 2 Unhidden"
                End If
                If GetFirstColumnInRow(WS, 2, True, True) <> 2 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Single cell Start 2, 2 Unhidden True True"
                End If
                If GetFirstColumnInRow(WS, 2, True, False) <> 2 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Single cell Start 2, 2 Unhidden True False"
                End If
                If GetFirstColumnInRow(WS, 2, False, True) <> 2 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Single cell Start 2, 2 Unhidden False True"
                End If
                If GetFirstColumnInRow(WS, 2, False, False) <> 2 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Single cell Start 2, 2 Unhidden False False"
                End If
                'Outside
                If GetFirstColumnInRow(WS, 1) <> 0 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Single cell Start 2, 2 Unhidden Outside less"
                End If
                If GetFirstColumnInRow(WS, 3) <> 0 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Single cell Start 2, 2 Unhidden Outside more"
                End If
            'Hidden Rows
            WS.Rows(2).Hidden = True
                If GetFirstColumnInRow(WS, 2) <> 2 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Single cell Start 2, 2 Hidden Rows"
                End If
                If GetFirstColumnInRow(WS, 2, True, True) <> 2 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Single cell Start 2, 2 Hidden Rows True True"
                End If
                If GetFirstColumnInRow(WS, 2, True, False) <> 2 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Single cell Start 2, 2 Hidden Rows True False"
                End If
                If GetFirstColumnInRow(WS, 2, False, True) <> 0 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Single cell Start 2, 2 Hidden Rows False True"
                End If
                If GetFirstColumnInRow(WS, 2, False, False) <> 0 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Single cell Start 2, 2 Hidden Rows False False"
                End If
                'Outside
                If GetFirstColumnInRow(WS, 1) <> 0 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Single cell Start 2, 2 Hidden Rows Outside less"
                End If
                If GetFirstColumnInRow(WS, 3) <> 0 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Single cell Start 2, 2 Hidden Rows Outside more"
                End If
            WS.Rows.Hidden = False
            'Hidden Columns
            WS.Columns(2).Hidden = True
                If GetFirstColumnInRow(WS, 2) <> 2 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Single cell Start 2, 2 Hidden Columns"
                End If
                If GetFirstColumnInRow(WS, 2, True, True) <> 2 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Single cell Start 2, 2 Hidden Columns True True"
                End If
                If GetFirstColumnInRow(WS, 2, True, False) <> 0 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Single cell Start 2, 2 Hidden Columns True False"
                End If
                If GetFirstColumnInRow(WS, 2, False, True) <> 2 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Single cell Start 2, 2 Hidden Columns False True"
                End If
                If GetFirstColumnInRow(WS, 2, False, False) <> 0 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Single cell Start 2, 2 Hidden Columns False False"
                End If
                'Outside
                If GetFirstColumnInRow(WS, 1) <> 0 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Single cell Start 2, 2 Hidden Columns Outside less"
                End If
                If GetFirstColumnInRow(WS, 3) <> 0 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Single cell Start 2, 2 Hidden Columns Outside more"
                End If
            WS.Columns.Hidden = False
            'Hidden Rows And Columns
            WS.Rows(2).Hidden = True
            WS.Columns(2).Hidden = True
                If GetFirstColumnInRow(WS, 2) <> 2 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Single cell Start 2, 2 Hidden Rows And Columns"
                End If
                If GetFirstColumnInRow(WS, 2, True, True) <> 2 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Single cell Start 2, 2 Hidden Rows And Columns True True"
                End If
                If GetFirstColumnInRow(WS, 2, True, False) <> 0 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Single cell Start 2, 2 Hidden Rows And Columns True False"
                End If
                If GetFirstColumnInRow(WS, 2, False, True) <> 0 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Single cell Start 2, 2 Hidden Rows And Columns False True"
                End If
                If GetFirstColumnInRow(WS, 2, False, False) <> 0 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Single cell Start 2, 2 Hidden Rows And Columns False False"
                End If
                'Outside
                If GetFirstColumnInRow(WS, 1) <> 0 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Single cell Start 2, 2 Hidden Rows And Columns Outside less"
                End If
                If GetFirstColumnInRow(WS, 3) <> 0 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Single cell Start 2, 2 Hidden Rows And Columns Outside more"
                End If
            WS.Rows.Hidden = False
            WS.Columns.Hidden = False
            'Filter
                'Cannot filter single cell
    
    'Multiple cell
        'Start 1, 1
        WS.Cells.Clear
        WS.Range("A1:C1").Value = "Header"
        WS.Range("A2:C4").Formula = "=ADDRESS(ROW(),COLUMN())"
            'Unhidden
                If GetFirstColumnInRow(WS, 1) <> 1 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Multiple cell Start 1, 1 Unhidden"
                End If
                If GetFirstColumnInRow(WS, 1, True, True) <> 1 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Multiple cell Start 1, 1 Unhidden True True"
                End If
                If GetFirstColumnInRow(WS, 1, True, False) <> 1 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Multiple cell Start 1, 1 Unhidden True False"
                End If
                If GetFirstColumnInRow(WS, 1, False, True) <> 1 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Multiple cell Start 1, 1 Unhidden False True"
                End If
                If GetFirstColumnInRow(WS, 1, False, False) <> 1 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Multiple cell Start 1, 1 Unhidden False False"
                End If
                'Outside
                If GetFirstColumnInRow(WS, 5) <> 0 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Multiple cell Start 1, 1 Unhidden Outside more"
                End If
            'Hidden Rows
            WS.Rows(1).Hidden = True
                If GetFirstColumnInRow(WS, 1) <> 1 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Multiple cell Start 1, 1 Hidden Rows"
                End If
                If GetFirstColumnInRow(WS, 1, True, True) <> 1 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Multiple cell Start 1, 1 Hidden Rows True True"
                End If
                If GetFirstColumnInRow(WS, 1, True, False) <> 1 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Multiple cell Start 1, 1 Hidden Rows True False"
                End If
                If GetFirstColumnInRow(WS, 1, False, True) <> 0 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Multiple cell Start 1, 1 Hidden Rows False True"
                End If
                If GetFirstColumnInRow(WS, 1, False, False) <> 0 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Multiple cell Start 1, 1 Hidden Rows False False"
                End If
                'Outside
                If GetFirstColumnInRow(WS, 5) <> 0 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Multiple cell Start 1, 1 Hidden Rows Outside more"
                End If
            WS.Rows.Hidden = False
            'Hidden Columns
            WS.Columns(1).Hidden = True
                If GetFirstColumnInRow(WS, 1) <> 1 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Multiple cell Start 1, 1 Hidden Columns"
                End If
                If GetFirstColumnInRow(WS, 1, True, True) <> 1 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Multiple cell Start 1, 1 Hidden Columns True True"
                End If
                If GetFirstColumnInRow(WS, 1, True, False) <> 2 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Multiple cell Start 1, 1 Hidden Columns True False"
                End If
                If GetFirstColumnInRow(WS, 1, False, True) <> 1 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Multiple cell Start 1, 1 Hidden Columns False True"
                End If
                If GetFirstColumnInRow(WS, 1, False, False) <> 2 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Multiple cell Start 1, 1 Hidden Columns False False"
                End If
                'Outside
                If GetFirstColumnInRow(WS, 5) <> 0 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Multiple cell Start 1, 1 Hidden Columns Outside more"
                End If
            WS.Columns.Hidden = False
            'Hidden Rows And Columns
            WS.Rows(1).Hidden = True
            WS.Columns(1).Hidden = True
                If GetFirstColumnInRow(WS, 1) <> 1 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Multiple cell Start 1, 1 Hidden Rows And Columns"
                End If
                If GetFirstColumnInRow(WS, 1, True, True) <> 1 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Multiple cell Start 1, 1 Hidden Rows And Columns True True"
                End If
                If GetFirstColumnInRow(WS, 1, True, False) <> 2 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Multiple cell Start 1, 1 Hidden Rows And Columns True False"
                End If
                If GetFirstColumnInRow(WS, 1, False, True) <> 0 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Multiple cell Start 1, 1 Hidden Rows And Columns False True"
                End If
                If GetFirstColumnInRow(WS, 1, False, False) <> 0 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Multiple cell Start 1, 1 Hidden Rows And Columns False False"
                End If
                'Outside
                If GetFirstColumnInRow(WS, 5) <> 0 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Multiple cell Start 1, 1 Hidden Rows And Columns Outside more"
                End If
            WS.Rows.Hidden = False
            WS.Columns.Hidden = False
            'Filter
            WS.Range("$A$1:$C$4").AutoFilter Field:=1, Criteria1:="$A$2"
                If GetFirstColumnInRow(WS, 3) <> 1 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Multiple cell Start 1, 1 Filter"
                End If
                If GetFirstColumnInRow(WS, 3, True, True) <> 1 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Multiple cell Start 1, 1 Filter True True"
                End If
                If GetFirstColumnInRow(WS, 3, True, False) <> 1 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Multiple cell Start 1, 1 Filter True False"
                End If
                If GetFirstColumnInRow(WS, 3, False, True) <> 0 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Multiple cell Start 1, 1 Filter False True"
                End If
                If GetFirstColumnInRow(WS, 3, False, False) <> 0 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Multiple cell Start 1, 1 Filter False False"
                End If
                'Outside
                If GetFirstColumnInRow(WS, 5) <> 0 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Multiple cell Start 1, 1 Filter Outside more"
                End If
            WS.AutoFilterMode = False
                
        'Start 2, 2
        WS.Cells.Clear
        WS.Range("B2:D2").Value = "Header"
        WS.Range("B3:D5").Formula = "=ADDRESS(ROW(),COLUMN())"
            'Unhidden
                If GetFirstColumnInRow(WS, 2) <> 2 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Multiple cell Start 2, 2 Unhidden"
                End If
                If GetFirstColumnInRow(WS, 2, True, True) <> 2 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Multiple cell Start 2, 2 Unhidden True True"
                End If
                If GetFirstColumnInRow(WS, 2, True, False) <> 2 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Multiple cell Start 2, 2 Unhidden True False"
                End If
                If GetFirstColumnInRow(WS, 2, False, True) <> 2 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Multiple cell Start 2, 2 Unhidden False True"
                End If
                If GetFirstColumnInRow(WS, 2, False, False) <> 2 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Multiple cell Start 2, 2 Unhidden False False"
                End If
                'Outside
                If GetFirstColumnInRow(WS, 1) <> 0 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Multiple cell Start 2, 2 Unhidden Outside less"
                End If
                If GetFirstColumnInRow(WS, 6) <> 0 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Multiple cell Start 2, 2 Unhidden Outside more"
                End If
            'Hidden Rows
            WS.Rows(2).Hidden = True
                If GetFirstColumnInRow(WS, 2) <> 2 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Multiple cell Start 2, 2 Hidden Rows"
                End If
                If GetFirstColumnInRow(WS, 2, True, True) <> 2 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Multiple cell Start 2, 2 Hidden Rows True True"
                End If
                If GetFirstColumnInRow(WS, 2, True, False) <> 2 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Multiple cell Start 2, 2 Hidden Rows True False"
                End If
                If GetFirstColumnInRow(WS, 2, False, True) <> 0 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Multiple cell Start 2, 2 Hidden Rows False True"
                End If
                If GetFirstColumnInRow(WS, 2, False, False) <> 0 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Multiple cell Start 2, 2 Hidden Rows False False"
                End If
                'Outside
                If GetFirstColumnInRow(WS, 1) <> 0 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Multiple cell Start 2, 2 Hidden Rows Outside less"
                End If
                If GetFirstColumnInRow(WS, 6) <> 0 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Multiple cell Start 2, 2 Hidden  Rows Outside more"
                End If
            WS.Rows.Hidden = False
            'Hidden Columns
            WS.Columns(2).Hidden = True
                If GetFirstColumnInRow(WS, 2) <> 2 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Multiple cell Start 2, 2 Hidden Columns"
                End If
                If GetFirstColumnInRow(WS, 2, True, True) <> 2 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Multiple cell Start 2, 2 Hidden Columns True True"
                End If
                If GetFirstColumnInRow(WS, 2, True, False) <> 3 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Multiple cell Start 2, 2 Hidden Columns True False"
                End If
                If GetFirstColumnInRow(WS, 2, False, True) <> 2 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Multiple cell Start 2, 2 Hidden Columns False True"
                End If
                If GetFirstColumnInRow(WS, 2, False, False) <> 3 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Multiple cell Start 2, 2 Hidden Columns False False"
                End If
                'Outside
                If GetFirstColumnInRow(WS, 1) <> 0 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Multiple cell Start 2, 2 Hidden Columns Outside less"
                End If
                If GetFirstColumnInRow(WS, 6) <> 0 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Multiple cell Start 2, 2 Hidden  Columns Outside more"
                End If
            WS.Columns.Hidden = False
            'Hidden Rows And Columns
            WS.Rows(2).Hidden = True
            WS.Columns(2).Hidden = True
                If GetFirstColumnInRow(WS, 2) <> 2 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Multiple cell Start 2, 2 Hidden Rows And Columns"
                End If
                If GetFirstColumnInRow(WS, 2, True, True) <> 2 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Multiple cell Start 2, 2 Hidden Rows And Columns True True"
                End If
                If GetFirstColumnInRow(WS, 2, True, False) <> 3 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Multiple cell Start 2, 2 Hidden Rows And Columns True False"
                End If
                If GetFirstColumnInRow(WS, 2, False, True) <> 0 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Multiple cell Start 2, 2 Hidden Rows And Columns False True"
                End If
                If GetFirstColumnInRow(WS, 2, False, False) <> 0 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Multiple cell Start 2, 2 Hidden Rows And Columns False False"
                End If
                'Outside
                If GetFirstColumnInRow(WS, 1) <> 0 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Multiple cell Start 2, 2 Hidden Rows And Columns Outside less"
                End If
                If GetFirstColumnInRow(WS, 6) <> 0 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Multiple cell Start 2, 2 Hidden Rows And Columns Outside more"
                End If
            WS.Rows.Hidden = False
            WS.Columns.Hidden = False
            'Filter
            WS.Range("$B$2:$D$5").AutoFilter Field:=1, Criteria1:="$B$3"
                If GetFirstColumnInRow(WS, 4) <> 2 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Multiple cell Start 2, 2 Filter"
                End If
                If GetFirstColumnInRow(WS, 4, True, True) <> 2 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Multiple cell Start 2, 2 Filter True True"
                End If
                If GetFirstColumnInRow(WS, 4, True, False) <> 2 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Multiple cell Start 2, 2 Filter True False"
                End If
                If GetFirstColumnInRow(WS, 4, False, True) <> 0 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Multiple cell Start 2, 2 Filter False True"
                End If
                If GetFirstColumnInRow(WS, 4, False, False) <> 0 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Multiple cell Start 2, 2 Filter False False"
                End If
                'Outside
                If GetFirstColumnInRow(WS, 1) <> 0 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Multiple cell Start 2, 2 Filter Outside less"
                End If
                If GetFirstColumnInRow(WS, 6) <> 0 Then
                    TestGetFirstColumnInRow = False
                    Debug.Print "Multiple cell Start 2, 2 Filter Outside more"
                End If
            WS.AutoFilterMode = False
            
    'Dynamic
    WS.Cells.Clear
    Dim i&
    For i = 5 To 1 Step -1
        WS.Cells(i, i).Value = i
        If GetFirstColumnInRow(WS, i) <> i Then
            TestGetFirstColumnInRow = False
            Debug.Print "Dynamic"
        End If
    Next i
    
    Excel.Application.DisplayAlerts = False
    WB.Close
    Excel.Application.DisplayAlerts = True
    
    Excel.Application.ScreenUpdating = True
    
    Debug.Print "TestGetFirstColumnInRow: " & TestGetFirstColumnInRow
    
End Function

Private Function TestGetLastRow() As Boolean
    
    TestGetLastRow = True

    Excel.Application.ScreenUpdating = False

    Dim WB As Workbook
    Set WB = Excel.Application.Workbooks.Add
    
    Dim WS As Worksheet
    Set WS = WB.Worksheets(1)
            
    'Blank
        'Unhidden
            If GetLastRow(WS) <> 0 Then
                TestGetLastRow = False
                Debug.Print "Blank Unhidden"
            End If
            If GetLastRow(WS, True, True) <> 0 Then
                TestGetLastRow = False
                Debug.Print "Blank Unhidden True True"
            End If
            If GetLastRow(WS, True, False) <> 0 Then
                TestGetLastRow = False
                Debug.Print "Blank Unhidden True False"
            End If
            If GetLastRow(WS, False, True) <> 0 Then
                TestGetLastRow = False
                Debug.Print "Blank Unhidden False True"
            End If
            If GetLastRow(WS, False, False) <> 0 Then
                TestGetLastRow = False
                Debug.Print "Blank Unhidden False False"
            End If
        'Hidden Row
        WS.Rows(1).Hidden = True
            If GetLastRow(WS) <> 0 Then
                TestGetLastRow = False
                Debug.Print "Blank Hidden Row"
            End If
            If GetLastRow(WS, True, True) <> 0 Then
                TestGetLastRow = False
                Debug.Print "Blank Hidden Row True True"
            End If
            If GetLastRow(WS, True, False) <> 0 Then
                TestGetLastRow = False
                Debug.Print "Blank Hidden Row True False"
            End If
            If GetLastRow(WS, False, True) <> 0 Then
                TestGetLastRow = False
                Debug.Print "Blank Hidden Row False True"
            End If
            If GetLastRow(WS, False, False) <> 0 Then
                TestGetLastRow = False
                Debug.Print "Blank Hidden Row False False"
            End If
        WS.Rows.Hidden = False
        'Hidden Column
        WS.Columns(1).Hidden = True
            If GetLastRow(WS) <> 0 Then
                TestGetLastRow = False
                Debug.Print "Blank Hidden Column"
            End If
            If GetLastRow(WS, True, True) <> 0 Then
                TestGetLastRow = False
                Debug.Print "Blank Hidden Column True True"
            End If
            If GetLastRow(WS, True, False) <> 0 Then
                TestGetLastRow = False
                Debug.Print "Blank Hidden Column True False"
            End If
            If GetLastRow(WS, False, True) <> 0 Then
                TestGetLastRow = False
                Debug.Print "Blank Hidden Column False True"
            End If
            If GetLastRow(WS, False, False) <> 0 Then
                TestGetLastRow = False
                Debug.Print "Blank Hidden Column False False"
            End If
        WS.Columns.Hidden = False
        'Hidden Row And Column
        WS.Columns(1).Hidden = True
        WS.Rows(1).Hidden = True
            If GetLastRow(WS) <> 0 Then
                TestGetLastRow = False
                Debug.Print "Blank Hidden Row And Column"
            End If
            If GetLastRow(WS, True, True) <> 0 Then
                TestGetLastRow = False
                Debug.Print "Blank Hidden Row And Column True True"
            End If
            If GetLastRow(WS, True, False) <> 0 Then
                TestGetLastRow = False
                Debug.Print "Blank Hidden Row And Column True False"
            End If
            If GetLastRow(WS, False, True) <> 0 Then
                TestGetLastRow = False
                Debug.Print "Blank Hidden Row And Column False True"
            End If
            If GetLastRow(WS, False, False) <> 0 Then
                TestGetLastRow = False
                Debug.Print "Blank Hidden Row And Column False False"
            End If
        WS.Rows.Hidden = False
        WS.Columns.Hidden = False
        'Filter
            'Cannot filter blank
            
    'Single cell
        'Start 1, 1
        WS.Cells(1, 1).Value = "Test"
            'Unhidden
                If GetLastRow(WS) <> 1 Then
                    TestGetLastRow = False
                    Debug.Print "Single cell Start 1, 1 Unhidden"
                End If
                If GetLastRow(WS, True, True) <> 1 Then
                    TestGetLastRow = False
                    Debug.Print "Single cell Start 1, 1 Unhidden True True"
                End If
                If GetLastRow(WS, True, False) <> 1 Then
                    TestGetLastRow = False
                    Debug.Print "Single cell Start 1, 1 Unhidden True False"
                End If
                If GetLastRow(WS, False, True) <> 1 Then
                    TestGetLastRow = False
                    Debug.Print "Single cell Start 1, 1 Unhidden False True"
                End If
                If GetLastRow(WS, False, False) <> 1 Then
                    TestGetLastRow = False
                    Debug.Print "Single cell Start 1, 1 Unhidden False False"
                End If
            'Hidden Row
            WS.Rows(1).Hidden = True
                If GetLastRow(WS) <> 1 Then
                    TestGetLastRow = False
                    Debug.Print "Single cell Start 1, 1 Hidden Row"
                End If
                If GetLastRow(WS, True, True) <> 1 Then
                    TestGetLastRow = False
                    Debug.Print "Single cell Start 1, 1 Hidden Row True True"
                End If
                If GetLastRow(WS, True, False) <> 1 Then
                    TestGetLastRow = False
                    Debug.Print "Single cell Start 1, 1 Hidden Row True False"
                End If
                If GetLastRow(WS, False, True) <> 0 Then
                    TestGetLastRow = False
                    Debug.Print "Single cell Start 1, 1 Hidden Row False True"
                End If
                If GetLastRow(WS, False, False) <> 0 Then
                    TestGetLastRow = False
                    Debug.Print "Single cell Start 1, 1 Hidden Row False False"
                End If
            WS.Rows.Hidden = False
            'Hidden Column
            WS.Columns(1).Hidden = True
                If GetLastRow(WS) <> 1 Then
                    TestGetLastRow = False
                    Debug.Print "Single cell Start 1, 1 Hidden Column"
                End If
                If GetLastRow(WS, True, True) <> 1 Then
                    TestGetLastRow = False
                    Debug.Print "Single cell Start 1, 1 Hidden Column True True"
                End If
                If GetLastRow(WS, True, False) <> 0 Then
                    TestGetLastRow = False
                    Debug.Print "Single cell Start 1, 1 Hidden Column True False"
                End If
                If GetLastRow(WS, False, True) <> 1 Then
                    TestGetLastRow = False
                    Debug.Print "Single cell Start 1, 1 Hidden Column False True"
                End If
                If GetLastRow(WS, False, False) <> 0 Then
                    TestGetLastRow = False
                    Debug.Print "Single cell Start 1, 1 Hidden Column False False"
                End If
            WS.Columns.Hidden = False
            'Hidden Row And Column
            WS.Columns(1).Hidden = True
            WS.Rows(1).Hidden = True
                If GetLastRow(WS) <> 1 Then
                    TestGetLastRow = False
                    Debug.Print "Single cell Start 1, 1 Hidden Row And Column"
                End If
                If GetLastRow(WS, True, True) <> 1 Then
                    TestGetLastRow = False
                    Debug.Print "Single cell Start 1, 1 Hidden Row And Column True True"
                End If
                If GetLastRow(WS, True, False) <> 0 Then
                    TestGetLastRow = False
                    Debug.Print "Single cell Start 1, 1 Hidden Row And Column True False"
                End If
                If GetLastRow(WS, False, True) <> 0 Then
                    TestGetLastRow = False
                    Debug.Print "Single cell Start 1, 1 Hidden Row And Column False True"
                End If
                If GetLastRow(WS, False, False) <> 0 Then
                    TestGetLastRow = False
                    Debug.Print "Single cell Start 1, 1 Hidden Row And Column False False"
                End If
            WS.Rows.Hidden = False
            WS.Columns.Hidden = False
            'Filter
                'Cannot filter single cell
        'Start 2, 2
        WS.Cells.Clear
        WS.Cells(2, 2).Value = "Test"
            'Unhidden
                If GetLastRow(WS) <> 2 Then
                    TestGetLastRow = False
                    Debug.Print "Single cell Start 2, 2 Unhidden"
                End If
                If GetLastRow(WS, True, True) <> 2 Then
                    TestGetLastRow = False
                    Debug.Print "Single cell Start 2, 2 Unhidden True True"
                End If
                If GetLastRow(WS, True, False) <> 2 Then
                    TestGetLastRow = False
                    Debug.Print "Single cell Start 2, 2 Unhidden True False"
                End If
                If GetLastRow(WS, False, True) <> 2 Then
                    TestGetLastRow = False
                    Debug.Print "Single cell Start 2, 2 Unhidden False True"
                End If
                If GetLastRow(WS, False, False) <> 2 Then
                    TestGetLastRow = False
                    Debug.Print "Single cell Start 2, 2 Unhidden False False"
                End If
            'Hidden Row
            WS.Rows(2).Hidden = True
                If GetLastRow(WS) <> 2 Then
                    TestGetLastRow = False
                    Debug.Print "Single cell Start 2, 2 Hidden Row"
                End If
                If GetLastRow(WS, True, True) <> 2 Then
                    TestGetLastRow = False
                    Debug.Print "Single cell Start 2, 2 Hidden Row True True"
                End If
                If GetLastRow(WS, True, False) <> 2 Then
                    TestGetLastRow = False
                    Debug.Print "Single cell Start 2, 2 Hidden Row True False"
                End If
                If GetLastRow(WS, False, True) <> 0 Then
                    TestGetLastRow = False
                    Debug.Print "Single cell Start 2, 2 Hidden Row False True"
                End If
                If GetLastRow(WS, False, False) <> 0 Then
                    TestGetLastRow = False
                    Debug.Print "Single cell Start 2, 2 Hidden Row False False"
                End If
            WS.Rows.Hidden = False
            'Hidden Column
            WS.Columns(2).Hidden = True
                If GetLastRow(WS) <> 2 Then
                    TestGetLastRow = False
                    Debug.Print "Single cell Start 2, 2 Hidden Column"
                End If
                If GetLastRow(WS, True, True) <> 2 Then
                    TestGetLastRow = False
                    Debug.Print "Single cell Start 2, 2 Hidden Column True True"
                End If
                If GetLastRow(WS, True, False) <> 0 Then
                    TestGetLastRow = False
                    Debug.Print "Single cell Start 2, 2 Hidden Column True False"
                End If
                If GetLastRow(WS, False, True) <> 2 Then
                    TestGetLastRow = False
                    Debug.Print "Single cell Start 2, 2 Hidden Column False True"
                End If
                If GetLastRow(WS, False, False) <> 0 Then
                    TestGetLastRow = False
                    Debug.Print "Single cell Start 2, 2 Hidden Column False False"
                End If
            WS.Columns.Hidden = False
            'Hidden Row And Column
            WS.Columns(2).Hidden = True
            WS.Rows(2).Hidden = True
                If GetLastRow(WS) <> 2 Then
                    TestGetLastRow = False
                    Debug.Print "Single cell Start 2, 2 Hidden Row And Column"
                End If
                If GetLastRow(WS, True, True) <> 2 Then
                    TestGetLastRow = False
                    Debug.Print "Single cell Start 2, 2 Hidden Row And Column True True"
                End If
                If GetLastRow(WS, True, False) <> 0 Then
                    TestGetLastRow = False
                    Debug.Print "Single cell Start 2, 2 Hidden Row And Column True False"
                End If
                If GetLastRow(WS, False, True) <> 0 Then
                    TestGetLastRow = False
                    Debug.Print "Single cell Start 2, 2 Hidden Row And Column False True"
                End If
                If GetLastRow(WS, False, False) <> 0 Then
                    TestGetLastRow = False
                    Debug.Print "Single cell Start 2, 2 Hidden Row And Column False False"
                End If
            WS.Rows.Hidden = False
            WS.Columns.Hidden = False
            'Filter
                'Cannot filter single cell
                
    'Muliple cell
        'Start 1, 1
        WS.Cells.Clear
        WS.Range("A1:C1").Value = "Header"
        WS.Range("A2:C4").Formula = "=ADDRESS(ROW(),COLUMN())"
            'Unhidden
                If GetLastRow(WS) <> 4 Then
                    TestGetLastRow = False
                    Debug.Print "Muliple cell Start 1, 1 Unhidden"
                End If
                If GetLastRow(WS, True, True) <> 4 Then
                    TestGetLastRow = False
                    Debug.Print "Muliple cell Start 1, 1 Unhidden True True"
                End If
                If GetLastRow(WS, True, False) <> 4 Then
                    TestGetLastRow = False
                    Debug.Print "Muliple cell Start 1, 1 Unhidden True False"
                End If
                If GetLastRow(WS, False, True) <> 4 Then
                    TestGetLastRow = False
                    Debug.Print "Muliple cell Start 1, 1 Unhidden False True"
                End If
                If GetLastRow(WS, False, False) <> 4 Then
                    TestGetLastRow = False
                    Debug.Print "Muliple cell Start 1, 1 Unhidden False False"
                End If
            'Hidden Row
            WS.Rows(4).Hidden = True
                If GetLastRow(WS) <> 4 Then
                    TestGetLastRow = False
                    Debug.Print "Muliple cell Start 1, 1 Hidden Row"
                End If
                If GetLastRow(WS, True, True) <> 4 Then
                    TestGetLastRow = False
                    Debug.Print "Muliple cell Start 1, 1 Hidden Row True True"
                End If
                If GetLastRow(WS, True, False) <> 4 Then
                    TestGetLastRow = False
                    Debug.Print "Muliple cell Start 1, 1 Hidden Row True False"
                End If
                If GetLastRow(WS, False, True) <> 3 Then
                    TestGetLastRow = False
                    Debug.Print "Muliple cell Start 1, 1 Hidden Row False True"
                End If
                If GetLastRow(WS, False, False) <> 3 Then
                    TestGetLastRow = False
                    Debug.Print "Muliple cell Start 1, 1 Hidden Row False False"
                End If
            WS.Rows.Hidden = False
            'Hidden Column
            WS.Columns(1).Hidden = True
                If GetLastRow(WS) <> 4 Then
                    TestGetLastRow = False
                    Debug.Print "Muliple cell Start 1, 1 Hidden Column"
                End If
                If GetLastRow(WS, True, True) <> 4 Then
                    TestGetLastRow = False
                    Debug.Print "Muliple cell Start 1, 1 Hidden Column True True"
                End If
                If GetLastRow(WS, True, False) <> 4 Then
                    TestGetLastRow = False
                    Debug.Print "Muliple cell Start 1, 1 Hidden Column True False"
                End If
                If GetLastRow(WS, False, True) <> 4 Then
                    TestGetLastRow = False
                    Debug.Print "Muliple cell Start 1, 1 Hidden Column False True"
                End If
                If GetLastRow(WS, False, False) <> 4 Then
                    TestGetLastRow = False
                    Debug.Print "Muliple cell Start 1, 1 Hidden Column False False"
                End If
            WS.Columns.Hidden = False
            'Hidden Row And Column
            WS.Columns(1).Hidden = True
            WS.Rows(4).Hidden = True
                If GetLastRow(WS) <> 4 Then
                    TestGetLastRow = False
                    Debug.Print "Muliple cell Start 1, 1 Hidden Row And Column"
                End If
                If GetLastRow(WS, True, True) <> 4 Then
                    TestGetLastRow = False
                    Debug.Print "Muliple cell Start 1, 1 Hidden Row And Column True True"
                End If
                If GetLastRow(WS, True, False) <> 4 Then
                    TestGetLastRow = False
                    Debug.Print "Muliple cell Start 1, 1 Hidden Row And Column True False"
                End If
                If GetLastRow(WS, False, True) <> 3 Then
                    TestGetLastRow = False
                    Debug.Print "Muliple cell Start 1, 1 Hidden Row And Column False True"
                End If
                If GetLastRow(WS, False, False) <> 3 Then
                    TestGetLastRow = False
                    Debug.Print "Muliple cell Start 1, 1 Hidden Row And Column False False"
                End If
            WS.Rows.Hidden = False
            WS.Columns.Hidden = False
            'Filter
            WS.Range("$A$1:$C$4").AutoFilter Field:=1, Criteria1:="$A$2"
                If GetLastRow(WS) <> 4 Then
                    TestGetLastRow = False
                    Debug.Print "Muliple cell Start 1, 1 Filter"
                End If
                If GetLastRow(WS, True, True) <> 4 Then
                    TestGetLastRow = False
                    Debug.Print "Muliple cell Start 1, 1 Filter True True"
                End If
                If GetLastRow(WS, True, False) <> 4 Then
                    TestGetLastRow = False
                    Debug.Print "Muliple cell Start 1, 1 Filter True False"
                End If
                If GetLastRow(WS, False, True) <> 2 Then
                    TestGetLastRow = False
                    Debug.Print "Muliple cell Start 1, 1 Filter False True"
                End If
                If GetLastRow(WS, False, False) <> 2 Then
                    TestGetLastRow = False
                    Debug.Print "Muliple cell Start 1, 1 Filter False False"
                End If
            WS.AutoFilterMode = False
        'Start 2, 2
        WS.Cells.Clear
        WS.Range("B2:D2").Value = "Header"
        WS.Range("B3:D5").Formula = "=ADDRESS(ROW(),COLUMN())"
            'Unhidden
                If GetLastRow(WS) <> 5 Then
                    TestGetLastRow = False
                    Debug.Print "Muliple cell Start 2, 2 Unhidden"
                End If
                If GetLastRow(WS, True, True) <> 5 Then
                    TestGetLastRow = False
                    Debug.Print "Muliple cell Start 2, 2 Unhidden True True"
                End If
                If GetLastRow(WS, True, False) <> 5 Then
                    TestGetLastRow = False
                    Debug.Print "Muliple cell Start 2, 2 Unhidden True False"
                End If
                If GetLastRow(WS, False, True) <> 5 Then
                    TestGetLastRow = False
                    Debug.Print "Muliple cell Start 2, 2 Unhidden False True"
                End If
                If GetLastRow(WS, False, False) <> 5 Then
                    TestGetLastRow = False
                    Debug.Print "Muliple cell Start 2, 2 Unhidden False False"
                End If
            'Hidden Row
            WS.Rows(5).Hidden = True
                If GetLastRow(WS) <> 5 Then
                    TestGetLastRow = False
                    Debug.Print "Muliple cell Start 2, 2 Hidden Row"
                End If
                If GetLastRow(WS, True, True) <> 5 Then
                    TestGetLastRow = False
                    Debug.Print "Muliple cell Start 2, 2 Hidden Row True True"
                End If
                If GetLastRow(WS, True, False) <> 5 Then
                    TestGetLastRow = False
                    Debug.Print "Muliple cell Start 2, 2 Hidden Row True False"
                End If
                If GetLastRow(WS, False, True) <> 4 Then
                    TestGetLastRow = False
                    Debug.Print "Muliple cell Start 2, 2 Hidden Row False True"
                End If
                If GetLastRow(WS, False, False) <> 4 Then
                    TestGetLastRow = False
                    Debug.Print "Muliple cell Start 2, 2 Hidden Row False False"
                End If
            WS.Rows.Hidden = False
            'Hidden Column
            WS.Columns(2).Hidden = True
                If GetLastRow(WS) <> 5 Then
                    TestGetLastRow = False
                    Debug.Print "Muliple cell Start 2, 2 Hidden Column"
                End If
                If GetLastRow(WS, True, True) <> 5 Then
                    TestGetLastRow = False
                    Debug.Print "Muliple cell Start 2, 2 Hidden Column True True"
                End If
                If GetLastRow(WS, True, False) <> 5 Then
                    TestGetLastRow = False
                    Debug.Print "Muliple cell Start 2, 2 Hidden Column True False"
                End If
                If GetLastRow(WS, False, True) <> 5 Then
                    TestGetLastRow = False
                    Debug.Print "Muliple cell Start 2, 2 Hidden Column False True"
                End If
                If GetLastRow(WS, False, False) <> 5 Then
                    TestGetLastRow = False
                    Debug.Print "Muliple cell Start 2, 2 Hidden Column False False"
                End If
            WS.Columns.Hidden = False
            'Hidden Row And Column
            WS.Columns(2).Hidden = True
            WS.Rows(5).Hidden = True
                If GetLastRow(WS) <> 5 Then
                    TestGetLastRow = False
                    Debug.Print "Muliple cell Start 2, 2 Hidden Row And Column"
                End If
                If GetLastRow(WS, True, True) <> 5 Then
                    TestGetLastRow = False
                    Debug.Print "Muliple cell Start 2, 2 Hidden Row And Column True True"
                End If
                If GetLastRow(WS, True, False) <> 5 Then
                    TestGetLastRow = False
                    Debug.Print "Muliple cell Start 2, 2 Hidden Row And Column True False"
                End If
                If GetLastRow(WS, False, True) <> 4 Then
                    TestGetLastRow = False
                    Debug.Print "Muliple cell Start 2, 2 Hidden Row And Column False True"
                End If
                If GetLastRow(WS, False, False) <> 4 Then
                    TestGetLastRow = False
                    Debug.Print "Muliple cell Start 2, 2 Hidden Row And Column False False"
                End If
            WS.Rows.Hidden = False
            WS.Columns.Hidden = False
            'Filter
            WS.Range("$B$2:$D$5").AutoFilter Field:=1, Criteria1:="$B$3"
                If GetLastRow(WS) <> 5 Then
                    TestGetLastRow = False
                    Debug.Print "Muliple cell Start 2, 2 Filter"
                End If
                If GetLastRow(WS, True, True) <> 5 Then
                    TestGetLastRow = False
                    Debug.Print "Muliple cell Start 2, 2 Filter True True"
                End If
                If GetLastRow(WS, True, False) <> 5 Then
                    TestGetLastRow = False
                    Debug.Print "Muliple cell Start 2, 2 Filter True False"
                End If
                If GetLastRow(WS, False, True) <> 3 Then
                    TestGetLastRow = False
                    Debug.Print "Muliple cell Start 2, 2 Filter False True"
                End If
                If GetLastRow(WS, False, False) <> 3 Then
                    TestGetLastRow = False
                    Debug.Print "Muliple cell Start 2, 2 Filter False False"
                End If
            WS.AutoFilterMode = False
            
    'Dynamic
    WS.Cells.Clear
    Dim i&
    For i = 1 To 5
        WS.Cells(i, i).Value = i
        If GetLastRow(WS) <> i Then
            TestGetLastRow = False
            Debug.Print "Dynamic"
        End If
    Next i
    
    Excel.Application.DisplayAlerts = False
    WB.Close
    Excel.Application.DisplayAlerts = True
    
    Excel.Application.ScreenUpdating = True
    
    Debug.Print "TestGetLastRow: " & TestGetLastRow
    
End Function

Private Function TestGetLastRowInColumn() As Boolean
    
    TestGetLastRowInColumn = True

    Excel.Application.ScreenUpdating = False

    Dim WB As Workbook
    Set WB = Excel.Application.Workbooks.Add
    
    Dim WS As Worksheet
    Set WS = WB.Worksheets(1)

    'Blank
        'Unhidden
            If GetLastRowInColumn(WS, 1) <> 0 Then
                TestGetLastRowInColumn = False
                Debug.Print "Blank Unhidden"
            End If
            If GetLastRowInColumn(WS, 1, True, True) <> 0 Then
                TestGetLastRowInColumn = False
                Debug.Print "Blank Unhidden True True"
            End If
            If GetLastRowInColumn(WS, 1, True, False) <> 0 Then
                TestGetLastRowInColumn = False
                Debug.Print "Blank Unhidden True False"
            End If
            If GetLastRowInColumn(WS, 1, False, True) <> 0 Then
                TestGetLastRowInColumn = False
                Debug.Print "Blank Unhidden False True"
            End If
            If GetLastRowInColumn(WS, 1, False, False) <> 0 Then
                TestGetLastRowInColumn = False
                Debug.Print "Blank Unhidden False False"
            End If
        'Hidden Rows
        WS.Rows(1).Hidden = True
            If GetLastRowInColumn(WS, 1) <> 0 Then
                TestGetLastRowInColumn = False
                Debug.Print "Blank Hidden Rows"
            End If
            If GetLastRowInColumn(WS, 1, True, True) <> 0 Then
                TestGetLastRowInColumn = False
                Debug.Print "Blank Hidden Rows True True"
            End If
            If GetLastRowInColumn(WS, 1, True, False) <> 0 Then
                TestGetLastRowInColumn = False
                Debug.Print "Blank Hidden Rows True False"
            End If
            If GetLastRowInColumn(WS, 1, False, True) <> 0 Then
                TestGetLastRowInColumn = False
                Debug.Print "Blank Hidden Rows False True"
            End If
            If GetLastRowInColumn(WS, 1, False, False) <> 0 Then
                TestGetLastRowInColumn = False
                Debug.Print "Blank Hidden Rows False False"
            End If
        WS.Rows.Hidden = False
        'Hidden Columns
        WS.Columns(1).Hidden = True
            If GetLastRowInColumn(WS, 1) <> 0 Then
                TestGetLastRowInColumn = False
                Debug.Print "Blank Hidden Columns"
            End If
            If GetLastRowInColumn(WS, 1, True, True) <> 0 Then
                TestGetLastRowInColumn = False
                Debug.Print "Blank Hidden Columns True True"
            End If
            If GetLastRowInColumn(WS, 1, True, False) <> 0 Then
                TestGetLastRowInColumn = False
                Debug.Print "Blank Hidden Columns True False"
            End If
            If GetLastRowInColumn(WS, 1, False, True) <> 0 Then
                TestGetLastRowInColumn = False
                Debug.Print "Blank Hidden Columns False True"
            End If
            If GetLastRowInColumn(WS, 1, False, False) <> 0 Then
                TestGetLastRowInColumn = False
                Debug.Print "Blank Hidden Columns False False"
            End If
        WS.Columns.Hidden = False
        'Hidden Rows And Columns
        WS.Rows(1).Hidden = True
        WS.Columns(1).Hidden = True
            If GetLastRowInColumn(WS, 1) <> 0 Then
                TestGetLastRowInColumn = False
                Debug.Print "Blank Hidden Rows And Columns"
            End If
            If GetLastRowInColumn(WS, 1, True, True) <> 0 Then
                TestGetLastRowInColumn = False
                Debug.Print "Blank Hidden Rows And Columns True True"
            End If
            If GetLastRowInColumn(WS, 1, True, False) <> 0 Then
                TestGetLastRowInColumn = False
                Debug.Print "Blank Hidden Rows And Columns True False"
            End If
            If GetLastRowInColumn(WS, 1, False, True) <> 0 Then
                TestGetLastRowInColumn = False
                Debug.Print "Blank Hidden Rows And Columns False True"
            End If
            If GetLastRowInColumn(WS, 1, False, False) <> 0 Then
                TestGetLastRowInColumn = False
                Debug.Print "Blank Hidden Rows And Columns False False"
            End If
        WS.Rows.Hidden = False
        WS.Columns.Hidden = False
        'Filter
            'Cannot filter blank
            
    'Single cell
        'Start 1, 1
        WS.Cells(1, 1).Value = "Test"
            'Unhidden
                If GetLastRowInColumn(WS, 1) <> 1 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Single cell Start 1, 1 Unhidden"
                End If
                If GetLastRowInColumn(WS, 1, True, True) <> 1 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Single cell Start 1, 1 Unhidden True True"
                End If
                If GetLastRowInColumn(WS, 1, True, False) <> 1 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Single cell Start 1, 1 Unhidden True False"
                End If
                If GetLastRowInColumn(WS, 1, False, True) <> 1 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Single cell Start 1, 1 Unhidden False True"
                End If
                If GetLastRowInColumn(WS, 1, False, False) <> 1 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Single cell Start 1, 1 Unhidden False False"
                End If
                'Outside
                If GetLastRowInColumn(WS, 2) <> 0 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Single cell Start 1, 1 Unhidden Outside more"
                End If
            'Hidden Rows
            WS.Rows(1).Hidden = True
                If GetLastRowInColumn(WS, 1) <> 1 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Single cell Start 1, 1 Hidden Rows"
                End If
                If GetLastRowInColumn(WS, 1, True, True) <> 1 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Single cell Start 1, 1 Hidden Rows True True"
                End If
                If GetLastRowInColumn(WS, 1, True, False) <> 1 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Single cell Start 1, 1 Hidden Rows True False"
                End If
                If GetLastRowInColumn(WS, 1, False, True) <> 0 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Single cell Start 1, 1 Hidden Rows False True"
                End If
                If GetLastRowInColumn(WS, 1, False, False) <> 0 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Single cell Start 1, 1 Hidden Rows False False"
                End If
                'Outside
                If GetLastRowInColumn(WS, 2) <> 0 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Single cell Start 1, 1 Hidden Rows Outside more"
                End If
            WS.Rows.Hidden = False
            'Hidden Columns
            WS.Columns(1).Hidden = True
                If GetLastRowInColumn(WS, 1) <> 1 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Single cell Start 1, 1 Hidden Columns"
                End If
                If GetLastRowInColumn(WS, 1, True, True) <> 1 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Single cell Start 1, 1 Hidden Columns True True"
                End If
                If GetLastRowInColumn(WS, 1, True, False) <> 0 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Single cell Start 1, 1 Hidden Columns True False"
                End If
                If GetLastRowInColumn(WS, 1, False, True) <> 1 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Single cell Start 1, 1 Hidden Columns False True"
                End If
                If GetLastRowInColumn(WS, 1, False, False) <> 0 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Single cell Start 1, 1 Hidden Columns False False"
                End If
                'Outside
                If GetLastRowInColumn(WS, 2) <> 0 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Single cell Start 1, 1 Hidden Columns Outside more"
                End If
            WS.Columns.Hidden = False
            'Hidden Rows And Columns
            WS.Rows(1).Hidden = True
            WS.Columns(1).Hidden = True
                If GetLastRowInColumn(WS, 1) <> 1 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Single cell Start 1, 1 Hidden Rows And Columns"
                End If
                If GetLastRowInColumn(WS, 1, True, True) <> 1 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Single cell Start 1, 1 Hidden Rows And Columns True True"
                End If
                If GetLastRowInColumn(WS, 1, True, False) <> 0 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Single cell Start 1, 1 Hidden Rows And Columns True False"
                End If
                If GetLastRowInColumn(WS, 1, False, True) <> 0 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Single cell Start 1, 1 Hidden Rows And Columns False True"
                End If
                If GetLastRowInColumn(WS, 1, False, False) <> 0 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Single cell Start 1, 1 Hidden Rows And Columns False False"
                End If
                'Outside
                If GetLastRowInColumn(WS, 2) <> 0 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Single cell Start 1, 1 Hidden Rows And Columns Outside more"
                End If
            WS.Rows.Hidden = False
            WS.Columns.Hidden = False
            'Filter
                'Cannot filter single cell

        'Start 2, 2
        WS.Cells.Clear
        WS.Cells(2, 2).Value = "Test"
            'Unhidden
                If GetLastRowInColumn(WS, 2) <> 2 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Single cell Start 2, 2 Unhidden"
                End If
                If GetLastRowInColumn(WS, 2, True, True) <> 2 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Single cell Start 2, 2 Unhidden True True"
                End If
                If GetLastRowInColumn(WS, 2, True, False) <> 2 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Single cell Start 2, 2 Unhidden True False"
                End If
                If GetLastRowInColumn(WS, 2, False, True) <> 2 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Single cell Start 2, 2 Unhidden False True"
                End If
                If GetLastRowInColumn(WS, 2, False, False) <> 2 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Single cell Start 2, 2 Unhidden False False"
                End If
                'Outside
                If GetLastRowInColumn(WS, 1) <> 0 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Single cell Start 2, 2 Unhidden Outside less"
                End If
                If GetLastRowInColumn(WS, 3) <> 0 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Single cell Start 2, 2 Unhidden Outside more"
                End If
            'Hidden Rows
            WS.Rows(2).Hidden = True
                If GetLastRowInColumn(WS, 2) <> 2 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Single cell Start 2, 2 Hidden Rows"
                End If
                If GetLastRowInColumn(WS, 2, True, True) <> 2 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Single cell Start 2, 2 Hidden Rows True True"
                End If
                If GetLastRowInColumn(WS, 2, True, False) <> 2 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Single cell Start 2, 2 Hidden Rows True False"
                End If
                If GetLastRowInColumn(WS, 2, False, True) <> 0 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Single cell Start 2, 2 Hidden Rows False True"
                End If
                If GetLastRowInColumn(WS, 2, False, False) <> 0 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Single cell Start 2, 2 Hidden Rows False False"
                End If
                'Outside
                If GetLastRowInColumn(WS, 1) <> 0 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Single cell Start 2, 2 Hidden Rows Outside less"
                End If
                If GetLastRowInColumn(WS, 3) <> 0 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Single cell Start 2, 2 Hidden Rows Outside more"
                End If
            WS.Rows.Hidden = False
            'Hidden Columns
            WS.Columns(2).Hidden = True
                If GetLastRowInColumn(WS, 2) <> 2 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Single cell Start 2, 2 Hidden Columns"
                End If
                If GetLastRowInColumn(WS, 2, True, True) <> 2 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Single cell Start 2, 2 Hidden Columns True True"
                End If
                If GetLastRowInColumn(WS, 2, True, False) <> 0 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Single cell Start 2, 2 Hidden Columns True False"
                End If
                If GetLastRowInColumn(WS, 2, False, True) <> 2 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Single cell Start 2, 2 Hidden Columns False True"
                End If
                If GetLastRowInColumn(WS, 2, False, False) <> 0 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Single cell Start 2, 2 Hidden Columns False False"
                End If
                'Outside
                If GetLastRowInColumn(WS, 1) <> 0 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Single cell Start 2, 2 Hidden Columns Outside less"
                End If
                If GetLastRowInColumn(WS, 3) <> 0 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Single cell Start 2, 2 Hidden Columns Outside more"
                End If
            WS.Columns.Hidden = False
            'Hidden Rows And Columns
            WS.Rows(2).Hidden = True
            WS.Columns(2).Hidden = True
                If GetLastRowInColumn(WS, 2) <> 2 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Single cell Start 2, 2 Hidden Rows And Columns"
                End If
                If GetLastRowInColumn(WS, 2, True, True) <> 2 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Single cell Start 2, 2 Hidden Rows And Columns True True"
                End If
                If GetLastRowInColumn(WS, 2, True, False) <> 0 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Single cell Start 2, 2 Hidden Rows And Columns True False"
                End If
                If GetLastRowInColumn(WS, 2, False, True) <> 0 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Single cell Start 2, 2 Hidden Rows And Columns False True"
                End If
                If GetLastRowInColumn(WS, 2, False, False) <> 0 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Single cell Start 2, 2 Hidden Rows And Columns False False"
                End If
                'Outside
                If GetLastRowInColumn(WS, 1) <> 0 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Single cell Start 2, 2 Hidden Rows And Columns Outside less"
                End If
                If GetLastRowInColumn(WS, 3) <> 0 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Single cell Start 2, 2 Hidden Rows And Columns Outside more"
                End If
            WS.Rows.Hidden = False
            WS.Columns.Hidden = False
            'Filter
                'Cannot filter single cell
    
    'Multiple cell
        'Start 1, 1
        WS.Cells.Clear
        WS.Range("A1:C1").Value = "Header"
        WS.Range("A2:C4").Formula = "=ADDRESS(ROW(),COLUMN())"
            'Unhidden
                If GetLastRowInColumn(WS, 1) <> 4 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Multiple cell Start 1, 1 Unhidden"
                End If
                If GetLastRowInColumn(WS, 1, True, True) <> 4 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Multiple cell Start 1, 1 Unhidden True True"
                End If
                If GetLastRowInColumn(WS, 1, True, False) <> 4 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Multiple cell Start 1, 1 Unhidden True False"
                End If
                If GetLastRowInColumn(WS, 1, False, True) <> 4 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Multiple cell Start 1, 1 Unhidden False True"
                End If
                If GetLastRowInColumn(WS, 1, False, False) <> 4 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Multiple cell Start 1, 1 Unhidden False False"
                End If
                'Outside
                If GetLastRowInColumn(WS, 4) <> 0 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Multiple cell Start 1, 1 Unhidden Outside more"
                End If
            'Hidden Rows
            WS.Rows(4).Hidden = True
                If GetLastRowInColumn(WS, 1) <> 4 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Multiple cell Start 1, 1 Hidden Rows"
                End If
                If GetLastRowInColumn(WS, 1, True, True) <> 4 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Multiple cell Start 1, 1 Hidden Rows True True"
                End If
                If GetLastRowInColumn(WS, 1, True, False) <> 4 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Multiple cell Start 1, 1 Hidden Rows True False"
                End If
                If GetLastRowInColumn(WS, 1, False, True) <> 3 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Multiple cell Start 1, 1 Hidden Rows False True"
                End If
                If GetLastRowInColumn(WS, 1, False, False) <> 3 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Multiple cell Start 1, 1 Hidden Rows False False"
                End If
                'Outside
                If GetLastRowInColumn(WS, 5) <> 0 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Multiple cell Start 1, 1 Hidden Rows Outside more"
                End If
            WS.Rows.Hidden = False
            'Hidden Columns
            WS.Columns(1).Hidden = True
                If GetLastRowInColumn(WS, 1) <> 4 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Multiple cell Start 1, 1 Hidden Columns"
                End If
                If GetLastRowInColumn(WS, 1, True, True) <> 4 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Multiple cell Start 1, 1 Hidden Columns True True"
                End If
                If GetLastRowInColumn(WS, 1, True, False) <> 0 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Multiple cell Start 1, 1 Hidden Columns True False"
                End If
                If GetLastRowInColumn(WS, 1, False, True) <> 4 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Multiple cell Start 1, 1 Hidden Columns False True"
                End If
                If GetLastRowInColumn(WS, 1, False, False) <> 0 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Multiple cell Start 1, 1 Hidden Columns False False"
                End If
                'Outside
                If GetLastRowInColumn(WS, 5) <> 0 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Multiple cell Start 1, 1 Hidden Columns Outside more"
                End If
            WS.Columns.Hidden = False
            'Hidden Rows And Columns
            WS.Rows(4).Hidden = True
            WS.Columns(1).Hidden = True
                If GetLastRowInColumn(WS, 1) <> 4 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Multiple cell Start 1, 1 Hidden Rows And Columns"
                End If
                If GetLastRowInColumn(WS, 1, True, True) <> 4 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Multiple cell Start 1, 1 Hidden Rows And Columns True True"
                End If
                If GetLastRowInColumn(WS, 1, True, False) <> 0 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Multiple cell Start 1, 1 Hidden Rows And Columns True False"
                End If
                If GetLastRowInColumn(WS, 1, False, True) <> 3 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Multiple cell Start 1, 1 Hidden Rows And Columns False True"
                End If
                If GetLastRowInColumn(WS, 1, False, False) <> 0 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Multiple cell Start 1, 1 Hidden Rows And Columns False False"
                End If
                'Outside
                If GetLastRowInColumn(WS, 5) <> 0 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Multiple cell Start 1, 1 Hidden Rows And Columns Outside more"
                End If
            WS.Rows.Hidden = False
            WS.Columns.Hidden = False
            'Filter
            WS.Range("$A$1:$C$4").AutoFilter Field:=1, Criteria1:="$A$2"
                If GetLastRowInColumn(WS, 1) <> 4 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Multiple cell Start 1, 1 Filter"
                End If
                If GetLastRowInColumn(WS, 1, True, True) <> 4 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Multiple cell Start 1, 1 Filter True True"
                End If
                If GetLastRowInColumn(WS, 1, True, False) <> 4 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Multiple cell Start 1, 1 Filter True False"
                End If
                If GetLastRowInColumn(WS, 1, False, True) <> 2 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Multiple cell Start 1, 1 Filter False True"
                End If
                If GetLastRowInColumn(WS, 1, False, False) <> 2 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Multiple cell Start 1, 1 Filter False False"
                End If
                'Outside
                If GetLastRowInColumn(WS, 5) <> 0 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Multiple cell Start 1, 1 Filter Outside more"
                End If
            WS.AutoFilterMode = False
                
        'Start 2, 2
        WS.Cells.Clear
        WS.Range("B2:D2").Value = "Header"
        WS.Range("B3:D5").Formula = "=ADDRESS(ROW(),COLUMN())"
            'Unhidden
                If GetLastRowInColumn(WS, 2) <> 5 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Multiple cell Start 2, 2 Unhidden"
                End If
                If GetLastRowInColumn(WS, 2, True, True) <> 5 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Multiple cell Start 2, 2 Unhidden True True"
                End If
                If GetLastRowInColumn(WS, 2, True, False) <> 5 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Multiple cell Start 2, 2 Unhidden True False"
                End If
                If GetLastRowInColumn(WS, 2, False, True) <> 5 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Multiple cell Start 2, 2 Unhidden False True"
                End If
                If GetLastRowInColumn(WS, 2, False, False) <> 5 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Multiple cell Start 2, 2 Unhidden False False"
                End If
                'Outside
                If GetLastRowInColumn(WS, 1) <> 0 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Multiple cell Start 2, 2 Unhidden Outside less"
                End If
                If GetLastRowInColumn(WS, 6) <> 0 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Multiple cell Start 2, 2 Unhidden Outside more"
                End If
            'Hidden Rows
            WS.Rows(5).Hidden = True
                If GetLastRowInColumn(WS, 2) <> 5 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Multiple cell Start 2, 2 Hidden Rows"
                End If
                If GetLastRowInColumn(WS, 2, True, True) <> 5 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Multiple cell Start 2, 2 Hidden Rows True True"
                End If
                If GetLastRowInColumn(WS, 2, True, False) <> 5 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Multiple cell Start 2, 2 Hidden Rows True False"
                End If
                If GetLastRowInColumn(WS, 2, False, True) <> 4 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Multiple cell Start 2, 2 Hidden Rows False True"
                End If
                If GetLastRowInColumn(WS, 2, False, False) <> 4 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Multiple cell Start 2, 2 Hidden Rows False False"
                End If
                'Outside
                If GetLastRowInColumn(WS, 1) <> 0 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Multiple cell Start 2, 2 Hidden Rows Outside less"
                End If
                If GetLastRowInColumn(WS, 6) <> 0 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Multiple cell Start 2, 2 Hidden  Rows Outside more"
                End If
            WS.Rows.Hidden = False
            'Hidden Columns
            WS.Columns(2).Hidden = True
                If GetLastRowInColumn(WS, 2) <> 5 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Multiple cell Start 2, 2 Hidden Columns"
                End If
                If GetLastRowInColumn(WS, 2, True, True) <> 5 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Multiple cell Start 2, 2 Hidden Columns True True"
                End If
                If GetLastRowInColumn(WS, 2, True, False) <> 0 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Multiple cell Start 2, 2 Hidden Columns True False"
                End If
                If GetLastRowInColumn(WS, 2, False, True) <> 5 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Multiple cell Start 2, 2 Hidden Columns False True"
                End If
                If GetLastRowInColumn(WS, 2, False, False) <> 0 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Multiple cell Start 2, 2 Hidden Columns False False"
                End If
                'Outside
                If GetLastRowInColumn(WS, 1) <> 0 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Multiple cell Start 2, 2 Hidden Columns Outside less"
                End If
                If GetLastRowInColumn(WS, 6) <> 0 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Multiple cell Start 2, 2 Hidden  Columns Outside more"
                End If
            WS.Columns.Hidden = False
            'Hidden Rows And Columns
            WS.Rows(5).Hidden = True
            WS.Columns(2).Hidden = True
                If GetLastRowInColumn(WS, 2) <> 5 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Multiple cell Start 2, 2 Hidden Rows And Columns"
                End If
                If GetLastRowInColumn(WS, 2, True, True) <> 5 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Multiple cell Start 2, 2 Hidden Rows And Columns True True"
                End If
                If GetLastRowInColumn(WS, 2, True, False) <> 0 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Multiple cell Start 2, 2 Hidden Rows And Columns True False"
                End If
                If GetLastRowInColumn(WS, 2, False, True) <> 4 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Multiple cell Start 2, 2 Hidden Rows And Columns False True"
                End If
                If GetLastRowInColumn(WS, 2, False, False) <> 0 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Multiple cell Start 2, 2 Hidden Rows And Columns False False"
                End If
                'Outside
                If GetLastRowInColumn(WS, 1) <> 0 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Multiple cell Start 2, 2 Hidden Rows And Columns Outside less"
                End If
                If GetLastRowInColumn(WS, 6) <> 0 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Multiple cell Start 2, 2 Hidden Rows And Columns Outside more"
                End If
            WS.Rows.Hidden = False
            WS.Columns.Hidden = False
            'Filter
            WS.Range("$B$2:$D$5").AutoFilter Field:=1, Criteria1:="$B$3"
                If GetLastRowInColumn(WS, 2) <> 5 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Multiple cell Start 2, 2 Filter"
                End If
                If GetLastRowInColumn(WS, 2, True, True) <> 5 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Multiple cell Start 2, 2 Filter True True"
                End If
                If GetLastRowInColumn(WS, 2, True, False) <> 5 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Multiple cell Start 2, 2 Filter True False"
                End If
                If GetLastRowInColumn(WS, 2, False, True) <> 3 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Multiple cell Start 2, 2 Filter False True"
                End If
                If GetLastRowInColumn(WS, 2, False, False) <> 3 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Multiple cell Start 2, 2 Filter False False"
                End If
                'Outside
                If GetLastRowInColumn(WS, 1) <> 0 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Multiple cell Start 2, 2 Filter Outside less"
                End If
                If GetLastRowInColumn(WS, 6) <> 0 Then
                    TestGetLastRowInColumn = False
                    Debug.Print "Multiple cell Start 2, 2 Filter Outside more"
                End If
            WS.AutoFilterMode = False
            
    'Dynamic
    WS.Cells.Clear
    Dim i&
    For i = 1 To 5
        WS.Cells(i, i).Value = i
        If GetLastRowInColumn(WS, i) <> i Then
            TestGetLastRowInColumn = False
            Debug.Print "Dynamic"
        End If
    Next i
    
    Excel.Application.DisplayAlerts = False
    WB.Close
    Excel.Application.DisplayAlerts = True
    
    Excel.Application.ScreenUpdating = True
    
    Debug.Print "TestGetLastRowInColumn: " & TestGetLastRowInColumn
    
End Function

Private Function TestGetLastColumn() As Boolean
    
    TestGetLastColumn = True
    
    Excel.Application.ScreenUpdating = False
    
    Dim WB As Workbook
    Set WB = Excel.Application.Workbooks.Add
    
    Dim WS As Worksheet
    Set WS = WB.Worksheets(1)
    
    'Blank
        'Unhidden
            If GetLastColumn(WS) <> 0 Then
                TestGetLastColumn = False
                Debug.Print "Blank Unhidden"
            End If
            If GetLastColumn(WS, True, True) <> 0 Then
                TestGetLastColumn = False
                Debug.Print "Blank Unhidden True True"
            End If
            If GetLastColumn(WS, True, False) <> 0 Then
                TestGetLastColumn = False
                Debug.Print "Blank Unhidden True False"
            End If
            If GetLastColumn(WS, False, True) <> 0 Then
                TestGetLastColumn = False
                Debug.Print "Blank Unhidden False True"
            End If
            If GetLastColumn(WS, False, False) <> 0 Then
                TestGetLastColumn = False
                Debug.Print "Blank Unhidden False False"
            End If
        'Hidden Row
        WS.Rows(1).Hidden = True
            If GetLastColumn(WS) <> 0 Then
                TestGetLastColumn = False
                Debug.Print "Blank Hidden Row"
            End If
            If GetLastColumn(WS, True, True) <> 0 Then
                TestGetLastColumn = False
                Debug.Print "Blank Hidden Row True True"
            End If
            If GetLastColumn(WS, True, False) <> 0 Then
                TestGetLastColumn = False
                Debug.Print "Blank Hidden Row True False"
            End If
            If GetLastColumn(WS, False, True) <> 0 Then
                TestGetLastColumn = False
                Debug.Print "Blank Hidden Row False True"
            End If
            If GetLastColumn(WS, False, False) <> 0 Then
                TestGetLastColumn = False
                Debug.Print "Blank Hidden Row False False"
            End If
        WS.Rows.Hidden = False
        'Hidden Column
        WS.Columns(1).Hidden = True
            If GetLastColumn(WS) <> 0 Then
                TestGetLastColumn = False
                Debug.Print "Blank Hidden Column"
            End If
            If GetLastColumn(WS, True, True) <> 0 Then
                TestGetLastColumn = False
                Debug.Print "Blank Hidden Column True True"
            End If
            If GetLastColumn(WS, True, False) <> 0 Then
                TestGetLastColumn = False
                Debug.Print "Blank Hidden Column True False"
            End If
            If GetLastColumn(WS, False, True) <> 0 Then
                TestGetLastColumn = False
                Debug.Print "Blank Hidden Column False True"
            End If
            If GetLastColumn(WS, False, False) <> 0 Then
                TestGetLastColumn = False
                Debug.Print "Blank Hidden Column False False"
            End If
        WS.Columns.Hidden = False
        'Hidden Row And Column
        WS.Columns(1).Hidden = True
        WS.Rows(1).Hidden = True
            If GetLastColumn(WS) <> 0 Then
                TestGetLastColumn = False
                Debug.Print "Blank Hidden Row And Column"
            End If
            If GetLastColumn(WS, True, True) <> 0 Then
                TestGetLastColumn = False
                Debug.Print "Blank Hidden Row And Column True True"
            End If
            If GetLastColumn(WS, True, False) <> 0 Then
                TestGetLastColumn = False
                Debug.Print "Blank Hidden Row And Column True False"
            End If
            If GetLastColumn(WS, False, True) <> 0 Then
                TestGetLastColumn = False
                Debug.Print "Blank Hidden Row And Column False True"
            End If
            If GetLastColumn(WS, False, False) <> 0 Then
                TestGetLastColumn = False
                Debug.Print "Blank Hidden Row And Column False False"
            End If
        WS.Rows.Hidden = False
        WS.Columns.Hidden = False
        'Filter
            'Cannot filter blank
            
    'Single cell
        'Start 1, 1
        WS.Cells(1, 1).Value = "Test"
            'Unhidden
                If GetLastColumn(WS) <> 1 Then
                    TestGetLastColumn = False
                    Debug.Print "Single cell Start 1, 1 Unhidden"
                End If
                If GetLastColumn(WS, True, True) <> 1 Then
                    TestGetLastColumn = False
                    Debug.Print "Single cell Start 1, 1 Unhidden True True"
                End If
                If GetLastColumn(WS, True, False) <> 1 Then
                    TestGetLastColumn = False
                    Debug.Print "Single cell Start 1, 1 Unhidden True False"
                End If
                If GetLastColumn(WS, False, True) <> 1 Then
                    TestGetLastColumn = False
                    Debug.Print "Single cell Start 1, 1 Unhidden False True"
                End If
                If GetLastColumn(WS, False, False) <> 1 Then
                    TestGetLastColumn = False
                    Debug.Print "Single cell Start 1, 1 Unhidden False False"
                End If
            'Hidden Row
            WS.Rows(1).Hidden = True
                If GetLastColumn(WS) <> 1 Then
                    TestGetLastColumn = False
                    Debug.Print "Single cell Start 1, 1 Hidden Row"
                End If
                If GetLastColumn(WS, True, True) <> 1 Then
                    TestGetLastColumn = False
                    Debug.Print "Single cell Start 1, 1 Hidden Row True True"
                End If
                If GetLastColumn(WS, True, False) <> 1 Then
                    TestGetLastColumn = False
                    Debug.Print "Single cell Start 1, 1 Hidden Row True False"
                End If
                If GetLastColumn(WS, False, True) <> 0 Then
                    TestGetLastColumn = False
                    Debug.Print "Single cell Start 1, 1 Hidden Row False True"
                End If
                If GetLastColumn(WS, False, False) <> 0 Then
                    TestGetLastColumn = False
                    Debug.Print "Single cell Start 1, 1 Hidden Row False False"
                End If
            WS.Rows.Hidden = False
            'Hidden Column
            WS.Columns(1).Hidden = True
                If GetLastColumn(WS) <> 1 Then
                    TestGetLastColumn = False
                    Debug.Print "Single cell Start 1, 1 Hidden Column"
                End If
                If GetLastColumn(WS, True, True) <> 1 Then
                    TestGetLastColumn = False
                    Debug.Print "Single cell Start 1, 1 Hidden Column True True"
                End If
                If GetLastColumn(WS, True, False) <> 0 Then
                    TestGetLastColumn = False
                    Debug.Print "Single cell Start 1, 1 Hidden Column True False"
                End If
                If GetLastColumn(WS, False, True) <> 1 Then
                    TestGetLastColumn = False
                    Debug.Print "Single cell Start 1, 1 Hidden Column False True"
                End If
                If GetLastColumn(WS, False, False) <> 0 Then
                    TestGetLastColumn = False
                    Debug.Print "Single cell Start 1, 1 Hidden Column False False"
                End If
            WS.Columns.Hidden = False
            'Hidden Row And Column
            WS.Columns(1).Hidden = True
            WS.Rows(1).Hidden = True
                If GetLastColumn(WS) <> 1 Then
                    TestGetLastColumn = False
                    Debug.Print "Single cell Start 1, 1 Hidden Row And Column"
                End If
                If GetLastColumn(WS, True, True) <> 1 Then
                    TestGetLastColumn = False
                    Debug.Print "Single cell Start 1, 1 Hidden Row And Column True True"
                End If
                If GetLastColumn(WS, True, False) <> 0 Then
                    TestGetLastColumn = False
                    Debug.Print "Single cell Start 1, 1 Hidden Row And Column True False"
                End If
                If GetLastColumn(WS, False, True) <> 0 Then
                    TestGetLastColumn = False
                    Debug.Print "Single cell Start 1, 1 Hidden Row And Column False True"
                End If
                If GetLastColumn(WS, False, False) <> 0 Then
                    TestGetLastColumn = False
                    Debug.Print "Single cell Start 1, 1 Hidden Row And Column False False"
                End If
            WS.Rows.Hidden = False
            WS.Columns.Hidden = False
            'Filter
                'Cannot filter single cell
        'Start 2, 2
        WS.Cells.Clear
        WS.Cells(2, 2).Value = "Test"
            'Unhidden
                If GetLastColumn(WS) <> 2 Then
                    TestGetLastColumn = False
                    Debug.Print "Single cell Start 2, 2 Unhidden"
                End If
                If GetLastColumn(WS, True, True) <> 2 Then
                    TestGetLastColumn = False
                    Debug.Print "Single cell Start 2, 2 Unhidden True True"
                End If
                If GetLastColumn(WS, True, False) <> 2 Then
                    TestGetLastColumn = False
                    Debug.Print "Single cell Start 2, 2 Unhidden True False"
                End If
                If GetLastColumn(WS, False, True) <> 2 Then
                    TestGetLastColumn = False
                    Debug.Print "Single cell Start 2, 2 Unhidden False True"
                End If
                If GetLastColumn(WS, False, False) <> 2 Then
                    TestGetLastColumn = False
                    Debug.Print "Single cell Start 2, 2 Unhidden False False"
                End If
            'Hidden Row
            WS.Rows(2).Hidden = True
                If GetLastColumn(WS) <> 2 Then
                    TestGetLastColumn = False
                    Debug.Print "Single cell Start 2, 2 Hidden Row"
                End If
                If GetLastColumn(WS, True, True) <> 2 Then
                    TestGetLastColumn = False
                    Debug.Print "Single cell Start 2, 2 Hidden Row True True"
                End If
                If GetLastColumn(WS, True, False) <> 2 Then
                    TestGetLastColumn = False
                    Debug.Print "Single cell Start 2, 2 Hidden Row True False"
                End If
                If GetLastColumn(WS, False, True) <> 0 Then
                    TestGetLastColumn = False
                    Debug.Print "Single cell Start 2, 2 Hidden Row False True"
                End If
                If GetLastColumn(WS, False, False) <> 0 Then
                    TestGetLastColumn = False
                    Debug.Print "Single cell Start 2, 2 Hidden Row False False"
                End If
            WS.Rows.Hidden = False
            'Hidden Column
            WS.Columns(2).Hidden = True
                If GetLastColumn(WS) <> 2 Then
                    TestGetLastColumn = False
                    Debug.Print "Single cell Start 2, 2 Hidden Column"
                End If
                If GetLastColumn(WS, True, True) <> 2 Then
                    TestGetLastColumn = False
                    Debug.Print "Single cell Start 2, 2 Hidden Column True True"
                End If
                If GetLastColumn(WS, True, False) <> 0 Then
                    TestGetLastColumn = False
                    Debug.Print "Single cell Start 2, 2 Hidden Column True False"
                End If
                If GetLastColumn(WS, False, True) <> 2 Then
                    TestGetLastColumn = False
                    Debug.Print "Single cell Start 2, 2 Hidden Column False True"
                End If
                If GetLastColumn(WS, False, False) <> 0 Then
                    TestGetLastColumn = False
                    Debug.Print "Single cell Start 2, 2 Hidden Column False False"
                End If
            WS.Columns.Hidden = False
            'Hidden Row And Column
            WS.Columns(2).Hidden = True
            WS.Rows(2).Hidden = True
                If GetLastColumn(WS) <> 2 Then
                    TestGetLastColumn = False
                    Debug.Print "Single cell Start 2, 2 Hidden Row And Column"
                End If
                If GetLastColumn(WS, True, True) <> 2 Then
                    TestGetLastColumn = False
                    Debug.Print "Single cell Start 2, 2 Hidden Row And Column True True"
                End If
                If GetLastColumn(WS, True, False) <> 0 Then
                    TestGetLastColumn = False
                    Debug.Print "Single cell Start 2, 2 Hidden Row And Column True False"
                End If
                If GetLastColumn(WS, False, True) <> 0 Then
                    TestGetLastColumn = False
                    Debug.Print "Single cell Start 2, 2 Hidden Row And Column False True"
                End If
                If GetLastColumn(WS, False, False) <> 0 Then
                    TestGetLastColumn = False
                    Debug.Print "Single cell Start 2, 2 Hidden Row And Column False False"
                End If
            WS.Rows.Hidden = False
            WS.Columns.Hidden = False
            'Filter
                'Cannot filter single cell
                
    'Muliple cell
        'Start 1, 1
        WS.Cells.Clear
        WS.Range("A1:C1").Value = "Header"
        WS.Range("A2:C4").Formula = "=ADDRESS(ROW(),COLUMN())"
            'Unhidden
                If GetLastColumn(WS) <> 3 Then
                    TestGetLastColumn = False
                    Debug.Print "Muliple cell Start 1, 1 Unhidden"
                End If
                If GetLastColumn(WS, True, True) <> 3 Then
                    TestGetLastColumn = False
                    Debug.Print "Muliple cell Start 1, 1 Unhidden True True"
                End If
                If GetLastColumn(WS, True, False) <> 3 Then
                    TestGetLastColumn = False
                    Debug.Print "Muliple cell Start 1, 1 Unhidden True False"
                End If
                If GetLastColumn(WS, False, True) <> 3 Then
                    TestGetLastColumn = False
                    Debug.Print "Muliple cell Start 1, 1 Unhidden False True"
                End If
                If GetLastColumn(WS, False, False) <> 3 Then
                    TestGetLastColumn = False
                    Debug.Print "Muliple cell Start 1, 1 Unhidden False False"
                End If
            'Hidden Row
            WS.Rows(1).Hidden = True
                If GetLastColumn(WS) <> 3 Then
                    TestGetLastColumn = False
                    Debug.Print "Muliple cell Start 1, 1 Hidden Row"
                End If
                If GetLastColumn(WS, True, True) <> 3 Then
                    TestGetLastColumn = False
                    Debug.Print "Muliple cell Start 1, 1 Hidden Row True True"
                End If
                If GetLastColumn(WS, True, False) <> 3 Then
                    TestGetLastColumn = False
                    Debug.Print "Muliple cell Start 1, 1 Hidden Row True False"
                End If
                If GetLastColumn(WS, False, True) <> 3 Then
                    TestGetLastColumn = False
                    Debug.Print "Muliple cell Start 1, 1 Hidden Row False True"
                End If
                If GetLastColumn(WS, False, False) <> 3 Then
                    TestGetLastColumn = False
                    Debug.Print "Muliple cell Start 1, 1 Hidden Row False False"
                End If
            WS.Rows.Hidden = False
            'Hidden Column
            WS.Columns(3).Hidden = True
                If GetLastColumn(WS) <> 3 Then
                    TestGetLastColumn = False
                    Debug.Print "Muliple cell Start 1, 1 Hidden Column"
                End If
                If GetLastColumn(WS, True, True) <> 3 Then
                    TestGetLastColumn = False
                    Debug.Print "Muliple cell Start 1, 1 Hidden Column True True"
                End If
                If GetLastColumn(WS, True, False) <> 2 Then
                    TestGetLastColumn = False
                    Debug.Print "Muliple cell Start 1, 1 Hidden Column True False"
                End If
                If GetLastColumn(WS, False, True) <> 3 Then
                    TestGetLastColumn = False
                    Debug.Print "Muliple cell Start 1, 1 Hidden Column False True"
                End If
                If GetLastColumn(WS, False, False) <> 2 Then
                    TestGetLastColumn = False
                    Debug.Print "Muliple cell Start 1, 1 Hidden Column False False"
                End If
            WS.Columns.Hidden = False
            'Hidden Row And Column
            WS.Columns(3).Hidden = True
            WS.Rows(1).Hidden = True
                If GetLastColumn(WS) <> 3 Then
                    TestGetLastColumn = False
                    Debug.Print "Muliple cell Start 1, 1 Hidden Row And Column"
                End If
                If GetLastColumn(WS, True, True) <> 3 Then
                    TestGetLastColumn = False
                    Debug.Print "Muliple cell Start 1, 1 Hidden Row And Column True True"
                End If
                If GetLastColumn(WS, True, False) <> 2 Then
                    TestGetLastColumn = False
                    Debug.Print "Muliple cell Start 1, 1 Hidden Row And Column True False"
                End If
                If GetLastColumn(WS, False, True) <> 3 Then
                    TestGetLastColumn = False
                    Debug.Print "Muliple cell Start 1, 1 Hidden Row And Column False True"
                End If
                If GetLastColumn(WS, False, False) <> 2 Then
                    TestGetLastColumn = False
                    Debug.Print "Muliple cell Start 1, 1 Hidden Row And Column False False"
                End If
            WS.Rows.Hidden = False
            WS.Columns.Hidden = False
            'Filter
            WS.Range("C1:C3").Clear
            WS.Range("$A$1:$C$4").AutoFilter Field:=1, Criteria1:="$A$2"
                If GetLastColumn(WS) <> 3 Then
                    TestGetLastColumn = False
                    Debug.Print "Muliple cell Start 1, 1 Filter"
                End If
                If GetLastColumn(WS, True, True) <> 3 Then
                    TestGetLastColumn = False
                    Debug.Print "Muliple cell Start 1, 1 Filter True True"
                End If
                If GetLastColumn(WS, True, False) <> 3 Then
                    TestGetLastColumn = False
                    Debug.Print "Muliple cell Start 1, 1 Filter True False"
                End If
                If GetLastColumn(WS, False, True) <> 2 Then
                    TestGetLastColumn = False
                    Debug.Print "Muliple cell Start 1, 1 Filter False True"
                End If
                If GetLastColumn(WS, False, False) <> 2 Then
                    TestGetLastColumn = False
                    Debug.Print "Muliple cell Start 1, 1 Filter False False"
                End If
            WS.AutoFilterMode = False
        'Start 2, 2
        WS.Cells.Clear
        WS.Range("B2:D2").Value = "Header"
        WS.Range("B3:D5").Formula = "=ADDRESS(ROW(),COLUMN())"
            'Unhidden
                If GetLastColumn(WS) <> 4 Then
                    TestGetLastColumn = False
                    Debug.Print "Muliple cell Start 2, 2 Unhidden"
                End If
                If GetLastColumn(WS, True, True) <> 4 Then
                    TestGetLastColumn = False
                    Debug.Print "Muliple cell Start 2, 2 Unhidden True True"
                End If
                If GetLastColumn(WS, True, False) <> 4 Then
                    TestGetLastColumn = False
                    Debug.Print "Muliple cell Start 2, 2 Unhidden True False"
                End If
                If GetLastColumn(WS, False, True) <> 4 Then
                    TestGetLastColumn = False
                    Debug.Print "Muliple cell Start 2, 2 Unhidden False True"
                End If
                If GetLastColumn(WS, False, False) <> 4 Then
                    TestGetLastColumn = False
                    Debug.Print "Muliple cell Start 2, 2 Unhidden False False"
                End If
            'Hidden Row
            WS.Rows(2).Hidden = True
                If GetLastColumn(WS) <> 4 Then
                    TestGetLastColumn = False
                    Debug.Print "Muliple cell Start 2, 2 Hidden Row"
                End If
                If GetLastColumn(WS, True, True) <> 4 Then
                    TestGetLastColumn = False
                    Debug.Print "Muliple cell Start 2, 2 Hidden Row True True"
                End If
                If GetLastColumn(WS, True, False) <> 4 Then
                    TestGetLastColumn = False
                    Debug.Print "Muliple cell Start 2, 2 Hidden Row True False"
                End If
                If GetLastColumn(WS, False, True) <> 4 Then
                    TestGetLastColumn = False
                    Debug.Print "Muliple cell Start 2, 2 Hidden Row False True"
                End If
                If GetLastColumn(WS, False, False) <> 4 Then
                    TestGetLastColumn = False
                    Debug.Print "Muliple cell Start 2, 2 Hidden Row False False"
                End If
            WS.Rows.Hidden = False
            'Hidden Column
            WS.Columns(4).Hidden = True
                If GetLastColumn(WS) <> 4 Then
                    TestGetLastColumn = False
                    Debug.Print "Muliple cell Start 2, 2 Hidden Column"
                End If
                If GetLastColumn(WS, True, True) <> 4 Then
                    TestGetLastColumn = False
                    Debug.Print "Muliple cell Start 2, 2 Hidden Column True True"
                End If
                If GetLastColumn(WS, True, False) <> 3 Then
                    TestGetLastColumn = False
                    Debug.Print "Muliple cell Start 2, 2 Hidden Column True False"
                End If
                If GetLastColumn(WS, False, True) <> 4 Then
                    TestGetLastColumn = False
                    Debug.Print "Muliple cell Start 2, 2 Hidden Column False True"
                End If
                If GetLastColumn(WS, False, False) <> 3 Then
                    TestGetLastColumn = False
                    Debug.Print "Muliple cell Start 2, 2 Hidden Column False False"
                End If
            WS.Columns.Hidden = False
            'Hidden Row And Column
            WS.Columns(4).Hidden = True
            WS.Rows(2).Hidden = True
                If GetLastColumn(WS) <> 4 Then
                    TestGetLastColumn = False
                    Debug.Print "Muliple cell Start 2, 2 Hidden Row And Column"
                End If
                If GetLastColumn(WS, True, True) <> 4 Then
                    TestGetLastColumn = False
                    Debug.Print "Muliple cell Start 2, 2 Hidden Row And Column True True"
                End If
                If GetLastColumn(WS, True, False) <> 3 Then
                    TestGetLastColumn = False
                    Debug.Print "Muliple cell Start 2, 2 Hidden Row And Column True False"
                End If
                If GetLastColumn(WS, False, True) <> 4 Then
                    TestGetLastColumn = False
                    Debug.Print "Muliple cell Start 2, 2 Hidden Row And Column False True"
                End If
                If GetLastColumn(WS, False, False) <> 3 Then
                    TestGetLastColumn = False
                    Debug.Print "Muliple cell Start 2, 2 Hidden Row And Column False False"
                End If
            WS.Rows.Hidden = False
            WS.Columns.Hidden = False
            'Filter
            WS.Range("D2:D4").Clear
            WS.Range("$B$2:$D$5").AutoFilter Field:=2, Criteria1:="$B$3"
                If GetLastColumn(WS) <> 4 Then
                    TestGetLastColumn = False
                    Debug.Print "Muliple cell Start 2, 2 Filter"
                End If
                If GetLastColumn(WS, True, True) <> 4 Then
                    TestGetLastColumn = False
                    Debug.Print "Muliple cell Start 2, 2 Filter True True"
                End If
                If GetLastColumn(WS, True, False) <> 4 Then
                    TestGetLastColumn = False
                    Debug.Print "Muliple cell Start 2, 2 Filter True False"
                End If
                If GetLastColumn(WS, False, True) <> 3 Then
                    TestGetLastColumn = False
                    Debug.Print "Muliple cell Start 2, 2 Filter False True"
                End If
                If GetLastColumn(WS, False, False) <> 3 Then
                    TestGetLastColumn = False
                    Debug.Print "Muliple cell Start 2, 2 Filter False False"
                End If
            WS.AutoFilterMode = False
     
    'Dynamic
    WS.Cells.Clear
    Dim i&
    For i = 1 To 5
        WS.Cells(i, i).Value = i
        If GetLastColumn(WS) <> i Then
            TestGetLastColumn = False
            Debug.Print "Dynamic"
        End If
    Next i
    
    Excel.Application.DisplayAlerts = False
    WB.Close
    Excel.Application.DisplayAlerts = True
    
    Excel.Application.ScreenUpdating = True
    
    Debug.Print "TestGetLastColumn: " & TestGetLastColumn
    
End Function

Private Function TestGetLastColumnInRow() As Boolean
    
    TestGetLastColumnInRow = True
    
    Excel.Application.ScreenUpdating = False
    
    Dim WB As Workbook
    Set WB = Excel.Application.Workbooks.Add
    
    Dim WS As Worksheet
    Set WS = WB.Worksheets(1)
            
    'Blank
        'Unhidden
            If GetLastColumnInRow(WS, 1) <> 0 Then
                TestGetLastColumnInRow = False
                Debug.Print "Blank Unhidden"
            End If
            If GetLastColumnInRow(WS, 1, True, True) <> 0 Then
                TestGetLastColumnInRow = False
                Debug.Print "Blank Unhidden True True"
            End If
            If GetLastColumnInRow(WS, 1, True, False) <> 0 Then
                TestGetLastColumnInRow = False
                Debug.Print "Blank Unhidden True False"
            End If
            If GetLastColumnInRow(WS, 1, False, True) <> 0 Then
                TestGetLastColumnInRow = False
                Debug.Print "Blank Unhidden False True"
            End If
            If GetLastColumnInRow(WS, 1, False, False) <> 0 Then
                TestGetLastColumnInRow = False
                Debug.Print "Blank Unhidden False False"
            End If
        'Hidden Rows
        WS.Rows(1).Hidden = True
            If GetLastColumnInRow(WS, 1) <> 0 Then
                TestGetLastColumnInRow = False
                Debug.Print "Blank Hidden Rows"
            End If
            If GetLastColumnInRow(WS, 1, True, True) <> 0 Then
                TestGetLastColumnInRow = False
                Debug.Print "Blank Hidden Rows True True"
            End If
            If GetLastColumnInRow(WS, 1, True, False) <> 0 Then
                TestGetLastColumnInRow = False
                Debug.Print "Blank Hidden Rows True False"
            End If
            If GetLastColumnInRow(WS, 1, False, True) <> 0 Then
                TestGetLastColumnInRow = False
                Debug.Print "Blank Hidden Rows False True"
            End If
            If GetLastColumnInRow(WS, 1, False, False) <> 0 Then
                TestGetLastColumnInRow = False
                Debug.Print "Blank Hidden Rows False False"
            End If
        WS.Rows.Hidden = False
        'Hidden Columns
        WS.Columns(1).Hidden = True
            If GetLastColumnInRow(WS, 1) <> 0 Then
                TestGetLastColumnInRow = False
                Debug.Print "Blank Hidden Columns"
            End If
            If GetLastColumnInRow(WS, 1, True, True) <> 0 Then
                TestGetLastColumnInRow = False
                Debug.Print "Blank Hidden Columns True True"
            End If
            If GetLastColumnInRow(WS, 1, True, False) <> 0 Then
                TestGetLastColumnInRow = False
                Debug.Print "Blank Hidden Columns True False"
            End If
            If GetLastColumnInRow(WS, 1, False, True) <> 0 Then
                TestGetLastColumnInRow = False
                Debug.Print "Blank Hidden Columns False True"
            End If
            If GetLastColumnInRow(WS, 1, False, False) <> 0 Then
                TestGetLastColumnInRow = False
                Debug.Print "Blank Hidden Columns False False"
            End If
        WS.Columns.Hidden = False
        'Hidden Rows And Columns
        WS.Rows(1).Hidden = True
        WS.Columns(1).Hidden = True
            If GetLastColumnInRow(WS, 1) <> 0 Then
                TestGetLastColumnInRow = False
                Debug.Print "Blank Hidden Rows And Columns"
            End If
            If GetLastColumnInRow(WS, 1, True, True) <> 0 Then
                TestGetLastColumnInRow = False
                Debug.Print "Blank Hidden Rows And Columns True True"
            End If
            If GetLastColumnInRow(WS, 1, True, False) <> 0 Then
                TestGetLastColumnInRow = False
                Debug.Print "Blank Hidden Rows And Columns True False"
            End If
            If GetLastColumnInRow(WS, 1, False, True) <> 0 Then
                TestGetLastColumnInRow = False
                Debug.Print "Blank Hidden Rows And Columns False True"
            End If
            If GetLastColumnInRow(WS, 1, False, False) <> 0 Then
                TestGetLastColumnInRow = False
                Debug.Print "Blank Hidden Rows And Columns False False"
            End If
        WS.Rows.Hidden = False
        WS.Columns.Hidden = False
        'Filter
            'Cannot filter blank
            
    'Single cell
        'Start 1, 1
        WS.Cells(1, 1).Value = "Test"
            'Unhidden
                If GetLastColumnInRow(WS, 1) <> 1 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Single cell Start 1, 1 Unhidden"
                End If
                If GetLastColumnInRow(WS, 1, True, True) <> 1 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Single cell Start 1, 1 Unhidden True True"
                End If
                If GetLastColumnInRow(WS, 1, True, False) <> 1 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Single cell Start 1, 1 Unhidden True False"
                End If
                If GetLastColumnInRow(WS, 1, False, True) <> 1 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Single cell Start 1, 1 Unhidden False True"
                End If
                If GetLastColumnInRow(WS, 1, False, False) <> 1 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Single cell Start 1, 1 Unhidden False False"
                End If
                'Outside
                If GetLastColumnInRow(WS, 2) <> 0 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Single cell Start 1, 1 Unhidden Outside more"
                End If
            'Hidden Rows
            WS.Rows(1).Hidden = True
                If GetLastColumnInRow(WS, 1) <> 1 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Single cell Start 1, 1 Hidden Rows"
                End If
                If GetLastColumnInRow(WS, 1, True, True) <> 1 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Single cell Start 1, 1 Hidden Rows True True"
                End If
                If GetLastColumnInRow(WS, 1, True, False) <> 1 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Single cell Start 1, 1 Hidden Rows True False"
                End If
                If GetLastColumnInRow(WS, 1, False, True) <> 0 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Single cell Start 1, 1 Hidden Rows False True"
                End If
                If GetLastColumnInRow(WS, 1, False, False) <> 0 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Single cell Start 1, 1 Hidden Rows False False"
                End If
                'Outside
                If GetLastColumnInRow(WS, 2) <> 0 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Single cell Start 1, 1 Hidden Rows Outside more"
                End If
            WS.Rows.Hidden = False
            'Hidden Columns
            WS.Columns(1).Hidden = True
                If GetLastColumnInRow(WS, 1) <> 1 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Single cell Start 1, 1 Hidden Columns"
                End If
                If GetLastColumnInRow(WS, 1, True, True) <> 1 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Single cell Start 1, 1 Hidden Columns True True"
                End If
                If GetLastColumnInRow(WS, 1, True, False) <> 0 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Single cell Start 1, 1 Hidden Columns True False"
                End If
                If GetLastColumnInRow(WS, 1, False, True) <> 1 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Single cell Start 1, 1 Hidden Columns False True"
                End If
                If GetLastColumnInRow(WS, 1, False, False) <> 0 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Single cell Start 1, 1 Hidden Columns False False"
                End If
                'Outside
                If GetLastColumnInRow(WS, 2) <> 0 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Single cell Start 1, 1 Hidden Columns Outside more"
                End If
            WS.Columns.Hidden = False
            'Hidden Rows And Columns
            WS.Rows(1).Hidden = True
            WS.Columns(1).Hidden = True
                If GetLastColumnInRow(WS, 1) <> 1 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Single cell Start 1, 1 Hidden Rows And Columns"
                End If
                If GetLastColumnInRow(WS, 1, True, True) <> 1 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Single cell Start 1, 1 Hidden Rows And Columns True True"
                End If
                If GetLastColumnInRow(WS, 1, True, False) <> 0 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Single cell Start 1, 1 Hidden Rows And Columns True False"
                End If
                If GetLastColumnInRow(WS, 1, False, True) <> 0 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Single cell Start 1, 1 Hidden Rows And Columns False True"
                End If
                If GetLastColumnInRow(WS, 1, False, False) <> 0 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Single cell Start 1, 1 Hidden Rows And Columns False False"
                End If
                'Outside
                If GetLastColumnInRow(WS, 2) <> 0 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Single cell Start 1, 1 Hidden Rows And Columns Outside more"
                End If
            WS.Rows.Hidden = False
            WS.Columns.Hidden = False
            'Filter
                'Cannot filter single cell

        'Start 2, 2
        WS.Cells.Clear
        WS.Cells(2, 2).Value = "Test"
            'Unhidden
                If GetLastColumnInRow(WS, 2) <> 2 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Single cell Start 2, 2 Unhidden"
                End If
                If GetLastColumnInRow(WS, 2, True, True) <> 2 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Single cell Start 2, 2 Unhidden True True"
                End If
                If GetLastColumnInRow(WS, 2, True, False) <> 2 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Single cell Start 2, 2 Unhidden True False"
                End If
                If GetLastColumnInRow(WS, 2, False, True) <> 2 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Single cell Start 2, 2 Unhidden False True"
                End If
                If GetLastColumnInRow(WS, 2, False, False) <> 2 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Single cell Start 2, 2 Unhidden False False"
                End If
                'Outside
                If GetLastColumnInRow(WS, 1) <> 0 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Single cell Start 2, 2 Unhidden Outside less"
                End If
                If GetLastColumnInRow(WS, 3) <> 0 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Single cell Start 2, 2 Unhidden Outside more"
                End If
            'Hidden Rows
            WS.Rows(2).Hidden = True
                If GetLastColumnInRow(WS, 2) <> 2 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Single cell Start 2, 2 Hidden Rows"
                End If
                If GetLastColumnInRow(WS, 2, True, True) <> 2 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Single cell Start 2, 2 Hidden Rows True True"
                End If
                If GetLastColumnInRow(WS, 2, True, False) <> 2 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Single cell Start 2, 2 Hidden Rows True False"
                End If
                If GetLastColumnInRow(WS, 2, False, True) <> 0 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Single cell Start 2, 2 Hidden Rows False True"
                End If
                If GetLastColumnInRow(WS, 2, False, False) <> 0 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Single cell Start 2, 2 Hidden Rows False False"
                End If
                'Outside
                If GetLastColumnInRow(WS, 1) <> 0 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Single cell Start 2, 2 Hidden Rows Outside less"
                End If
                If GetLastColumnInRow(WS, 3) <> 0 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Single cell Start 2, 2 Hidden Rows Outside more"
                End If
            WS.Rows.Hidden = False
            'Hidden Columns
            WS.Columns(2).Hidden = True
                If GetLastColumnInRow(WS, 2) <> 2 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Single cell Start 2, 2 Hidden Columns"
                End If
                If GetLastColumnInRow(WS, 2, True, True) <> 2 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Single cell Start 2, 2 Hidden Columns True True"
                End If
                If GetLastColumnInRow(WS, 2, True, False) <> 0 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Single cell Start 2, 2 Hidden Columns True False"
                End If
                If GetLastColumnInRow(WS, 2, False, True) <> 2 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Single cell Start 2, 2 Hidden Columns False True"
                End If
                If GetLastColumnInRow(WS, 2, False, False) <> 0 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Single cell Start 2, 2 Hidden Columns False False"
                End If
                'Outside
                If GetLastColumnInRow(WS, 1) <> 0 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Single cell Start 2, 2 Hidden Columns Outside less"
                End If
                If GetLastColumnInRow(WS, 3) <> 0 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Single cell Start 2, 2 Hidden Columns Outside more"
                End If
            WS.Columns.Hidden = False
            'Hidden Rows And Columns
            WS.Rows(2).Hidden = True
            WS.Columns(2).Hidden = True
                If GetLastColumnInRow(WS, 2) <> 2 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Single cell Start 2, 2 Hidden Rows And Columns"
                End If
                If GetLastColumnInRow(WS, 2, True, True) <> 2 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Single cell Start 2, 2 Hidden Rows And Columns True True"
                End If
                If GetLastColumnInRow(WS, 2, True, False) <> 0 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Single cell Start 2, 2 Hidden Rows And Columns True False"
                End If
                If GetLastColumnInRow(WS, 2, False, True) <> 0 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Single cell Start 2, 2 Hidden Rows And Columns False True"
                End If
                If GetLastColumnInRow(WS, 2, False, False) <> 0 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Single cell Start 2, 2 Hidden Rows And Columns False False"
                End If
                'Outside
                If GetLastColumnInRow(WS, 1) <> 0 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Single cell Start 2, 2 Hidden Rows And Columns Outside less"
                End If
                If GetLastColumnInRow(WS, 3) <> 0 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Single cell Start 2, 2 Hidden Rows And Columns Outside more"
                End If
            WS.Rows.Hidden = False
            WS.Columns.Hidden = False
            'Filter
                'Cannot filter single cell
    
    'Multiple cell
        'Start 1, 1
        WS.Cells.Clear
        WS.Range("A1:C1").Value = "Header"
        WS.Range("A2:C4").Formula = "=ADDRESS(ROW(),COLUMN())"
            'Unhidden
                If GetLastColumnInRow(WS, 1) <> 3 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Multiple cell Start 1, 1 Unhidden"
                End If
                If GetLastColumnInRow(WS, 1, True, True) <> 3 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Multiple cell Start 1, 1 Unhidden True True"
                End If
                If GetLastColumnInRow(WS, 1, True, False) <> 3 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Multiple cell Start 1, 1 Unhidden True False"
                End If
                If GetLastColumnInRow(WS, 1, False, True) <> 3 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Multiple cell Start 1, 1 Unhidden False True"
                End If
                If GetLastColumnInRow(WS, 1, False, False) <> 3 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Multiple cell Start 1, 1 Unhidden False False"
                End If
                'Outside
                If GetLastColumnInRow(WS, 5) <> 0 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Multiple cell Start 1, 1 Unhidden Outside more"
                End If
            'Hidden Rows
            WS.Rows(1).Hidden = True
                If GetLastColumnInRow(WS, 1) <> 3 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Multiple cell Start 1, 1 Hidden Rows"
                End If
                If GetLastColumnInRow(WS, 1, True, True) <> 3 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Multiple cell Start 1, 1 Hidden Rows True True"
                End If
                If GetLastColumnInRow(WS, 1, True, False) <> 3 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Multiple cell Start 1, 1 Hidden Rows True False"
                End If
                If GetLastColumnInRow(WS, 1, False, True) <> 0 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Multiple cell Start 1, 1 Hidden Rows False True"
                End If
                If GetLastColumnInRow(WS, 1, False, False) <> 0 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Multiple cell Start 1, 1 Hidden Rows False False"
                End If
                'Outside
                If GetLastColumnInRow(WS, 5) <> 0 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Multiple cell Start 1, 1 Hidden Rows Outside more"
                End If
            WS.Rows.Hidden = False
            'Hidden Columns
            WS.Columns(3).Hidden = True
                If GetLastColumnInRow(WS, 1) <> 3 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Multiple cell Start 1, 1 Hidden Columns"
                End If
                If GetLastColumnInRow(WS, 1, True, True) <> 3 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Multiple cell Start 1, 1 Hidden Columns True True"
                End If
                If GetLastColumnInRow(WS, 1, True, False) <> 2 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Multiple cell Start 1, 1 Hidden Columns True False"
                End If
                If GetLastColumnInRow(WS, 1, False, True) <> 3 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Multiple cell Start 1, 1 Hidden Columns False True"
                End If
                If GetLastColumnInRow(WS, 1, False, False) <> 2 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Multiple cell Start 1, 1 Hidden Columns False False"
                End If
                'Outside
                If GetLastColumnInRow(WS, 5) <> 0 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Multiple cell Start 1, 1 Hidden Columns Outside more"
                End If
            WS.Columns.Hidden = False
            'Hidden Rows And Columns
            WS.Rows(4).Hidden = True
            WS.Columns(3).Hidden = True
                If GetLastColumnInRow(WS, 4) <> 3 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Multiple cell Start 1, 1 Hidden Rows And Columns"
                End If
                If GetLastColumnInRow(WS, 4, True, True) <> 3 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Multiple cell Start 1, 1 Hidden Rows And Columns True True"
                End If
                If GetLastColumnInRow(WS, 4, True, False) <> 2 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Multiple cell Start 1, 1 Hidden Rows And Columns True False"
                End If
                If GetLastColumnInRow(WS, 4, False, True) <> 0 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Multiple cell Start 1, 1 Hidden Rows And Columns False True"
                End If
                If GetLastColumnInRow(WS, 4, False, False) <> 0 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Multiple cell Start 1, 1 Hidden Rows And Columns False False"
                End If
                'Outside
                If GetLastColumnInRow(WS, 5) <> 0 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Multiple cell Start 1, 1 Hidden Rows And Columns Outside more"
                End If
            WS.Rows.Hidden = False
            WS.Columns.Hidden = False
            'Filter
            WS.Range("$A$1:$C$4").AutoFilter Field:=1, Criteria1:="$A$2"
                If GetLastColumnInRow(WS, 4) <> 3 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Multiple cell Start 1, 1 Filter"
                End If
                If GetLastColumnInRow(WS, 4, True, True) <> 3 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Multiple cell Start 1, 1 Filter True True"
                End If
                If GetLastColumnInRow(WS, 4, True, False) <> 3 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Multiple cell Start 1, 1 Filter True False"
                End If
                If GetLastColumnInRow(WS, 4, False, True) <> 0 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Multiple cell Start 1, 1 Filter False True"
                End If
                If GetLastColumnInRow(WS, 4, False, False) <> 0 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Multiple cell Start 1, 1 Filter False False"
                End If
                'Outside
                If GetLastColumnInRow(WS, 5) <> 0 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Multiple cell Start 1, 1 Filter Outside more"
                End If
            WS.AutoFilterMode = False
                
        'Start 2, 2
        WS.Cells.Clear
        WS.Range("B2:D2").Value = "Header"
        WS.Range("B3:D5").Formula = "=ADDRESS(ROW(),COLUMN())"
            'Unhidden
                If GetLastColumnInRow(WS, 5) <> 4 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Multiple cell Start 2, 2 Unhidden"
                End If
                If GetLastColumnInRow(WS, 2, True, True) <> 4 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Multiple cell Start 2, 2 Unhidden True True"
                End If
                If GetLastColumnInRow(WS, 2, True, False) <> 4 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Multiple cell Start 2, 2 Unhidden True False"
                End If
                If GetLastColumnInRow(WS, 2, False, True) <> 4 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Multiple cell Start 2, 2 Unhidden False True"
                End If
                If GetLastColumnInRow(WS, 2, False, False) <> 4 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Multiple cell Start 2, 2 Unhidden False False"
                End If
                'Outside
                If GetLastColumnInRow(WS, 1) <> 0 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Multiple cell Start 2, 2 Unhidden Outside less"
                End If
                If GetLastColumnInRow(WS, 6) <> 0 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Multiple cell Start 2, 2 Unhidden Outside more"
                End If
            'Hidden Rows
            WS.Rows(5).Hidden = True
                If GetLastColumnInRow(WS, 5) <> 4 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Multiple cell Start 2, 2 Hidden Rows"
                End If
                If GetLastColumnInRow(WS, 5, True, True) <> 4 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Multiple cell Start 2, 2 Hidden Rows True True"
                End If
                If GetLastColumnInRow(WS, 5, True, False) <> 4 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Multiple cell Start 2, 2 Hidden Rows True False"
                End If
                If GetLastColumnInRow(WS, 5, False, True) <> 0 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Multiple cell Start 2, 2 Hidden Rows False True"
                End If
                If GetLastColumnInRow(WS, 5, False, False) <> 0 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Multiple cell Start 2, 2 Hidden Rows False False"
                End If
                'Outside
                If GetLastColumnInRow(WS, 1) <> 0 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Multiple cell Start 2, 2 Hidden Rows Outside less"
                End If
                If GetLastColumnInRow(WS, 6) <> 0 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Multiple cell Start 2, 2 Hidden  Rows Outside more"
                End If
            WS.Rows.Hidden = False
            'Hidden Columns
            WS.Columns(4).Hidden = True
                If GetLastColumnInRow(WS, 2) <> 4 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Multiple cell Start 2, 2 Hidden Columns"
                End If
                If GetLastColumnInRow(WS, 2, True, True) <> 4 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Multiple cell Start 2, 2 Hidden Columns True True"
                End If
                If GetLastColumnInRow(WS, 2, True, False) <> 3 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Multiple cell Start 2, 2 Hidden Columns True False"
                End If
                If GetLastColumnInRow(WS, 2, False, True) <> 4 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Multiple cell Start 2, 2 Hidden Columns False True"
                End If
                If GetLastColumnInRow(WS, 2, False, False) <> 3 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Multiple cell Start 2, 2 Hidden Columns False False"
                End If
                'Outside
                If GetLastColumnInRow(WS, 1) <> 0 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Multiple cell Start 2, 2 Hidden Columns Outside less"
                End If
                If GetLastColumnInRow(WS, 6) <> 0 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Multiple cell Start 2, 2 Hidden  Columns Outside more"
                End If
            WS.Columns.Hidden = False
            'Hidden Rows And Columns
            WS.Rows(5).Hidden = True
            WS.Columns(4).Hidden = True
                If GetLastColumnInRow(WS, 5) <> 4 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Multiple cell Start 2, 2 Hidden Rows And Columns"
                End If
                If GetLastColumnInRow(WS, 5, True, True) <> 4 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Multiple cell Start 2, 2 Hidden Rows And Columns True True"
                End If
                If GetLastColumnInRow(WS, 5, True, False) <> 3 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Multiple cell Start 2, 2 Hidden Rows And Columns True False"
                End If
                If GetLastColumnInRow(WS, 5, False, True) <> 0 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Multiple cell Start 2, 2 Hidden Rows And Columns False True"
                End If
                If GetLastColumnInRow(WS, 5, False, False) <> 0 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Multiple cell Start 2, 2 Hidden Rows And Columns False False"
                End If
                'Outside
                If GetLastColumnInRow(WS, 1) <> 0 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Multiple cell Start 2, 2 Hidden Rows And Columns Outside less"
                End If
                If GetLastColumnInRow(WS, 6) <> 0 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Multiple cell Start 2, 2 Hidden Rows And Columns Outside more"
                End If
            WS.Rows.Hidden = False
            WS.Columns.Hidden = False
            'Filter
            WS.Range("$B$2:$D$5").AutoFilter Field:=1, Criteria1:="$B$3"
                If GetLastColumnInRow(WS, 5) <> 4 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Multiple cell Start 2, 2 Filter"
                End If
                If GetLastColumnInRow(WS, 5, True, True) <> 4 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Multiple cell Start 2, 2 Filter True True"
                End If
                If GetLastColumnInRow(WS, 5, True, False) <> 4 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Multiple cell Start 2, 2 Filter True False"
                End If
                If GetLastColumnInRow(WS, 5, False, True) <> 0 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Multiple cell Start 2, 2 Filter False True"
                End If
                If GetLastColumnInRow(WS, 5, False, False) <> 0 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Multiple cell Start 2, 2 Filter False False"
                End If
                'Outside
                If GetLastColumnInRow(WS, 1) <> 0 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Multiple cell Start 2, 2 Filter Outside less"
                End If
                If GetLastColumnInRow(WS, 6) <> 0 Then
                    TestGetLastColumnInRow = False
                    Debug.Print "Multiple cell Start 2, 2 Filter Outside more"
                End If
            WS.AutoFilterMode = False
            
    'Dynamic
    WS.Cells.Clear
    Dim i&
    For i = 1 To 5
        WS.Cells(i, i).Value = i
        If GetLastColumnInRow(WS, i) <> i Then
            TestGetLastColumnInRow = False
            Debug.Print "Dynamic"
        End If
    Next i
    
    Excel.Application.DisplayAlerts = False
    WB.Close
    Excel.Application.DisplayAlerts = True
    
    Excel.Application.ScreenUpdating = True
    
    Debug.Print "TestGetLastColumnInRow: " & TestGetLastColumnInRow
    
End Function

Private Function TestGetWholeRange() As Boolean
    
    TestGetWholeRange = True
    
    Excel.Application.ScreenUpdating = False
    
    Dim WB As Workbook
    Set WB = Excel.Application.Workbooks.Add
    
    Dim WS As Worksheet
    Set WS = WB.Worksheets(1)
        
    'Blank
    If Not GetWholeRange(WS) Is Nothing Then
        TestGetWholeRange = False
        Debug.Print "Blank"
    End If
    
    'Single cell
        'Start 1, 1
        WS.Cells(1, 1).Value = "Test"
        If GetWholeRange(WS).Address <> "$A$1" Then
            TestGetWholeRange = False
            Debug.Print "Single cell Start 1, 1"
        End If
        'Start 2, 2
        WS.Cells.Clear
        WS.Cells(2, 2).Value = "Test"
        If GetWholeRange(WS).Address <> "$A$1:$B$2" Then
            TestGetWholeRange = False
            Debug.Print "Single cell Start 2, 2"
        End If

    'Multiple cell
        'Start 1, 1
        WS.Cells.Clear
        WS.Range("A1:C1").Value = "Header"
        WS.Range("A2:C4").Formula = "=ADDRESS(ROW(),COLUMN())"
        If GetWholeRange(WS).Address <> "$A$1:$C$4" Then
            TestGetWholeRange = False
            Debug.Print "Multiple cell Start 1, 1"
        End If
        'Start 2, 2
        WS.Cells.Clear
        WS.Range("B2:D2").Value = "Header"
        WS.Range("B3:D5").Formula = "=ADDRESS(ROW(),COLUMN())"
        If GetWholeRange(WS).Address <> "$A$1:$D$5" Then
            TestGetWholeRange = False
            Debug.Print "Multiple cell Start 2, 2"
        End If
        
    'Hidden
    WS.Cells.Clear
    WS.Range("A1:C1").Value = "Header"
    WS.Range("A2:C4").Formula = "=ADDRESS(ROW(),COLUMN())"
    WS.Rows("1:4").Hidden = True
    WS.Columns("A:C").Hidden = True
    If GetWholeRange(WS).Address <> "$A$1:$C$4" Then
        TestGetWholeRange = False
        Debug.Print "Hidden"
    End If
    
    'Filtered
    WS.Rows.Hidden = False
    WS.Columns.Hidden = False
    WS.Cells.Clear
    WS.Range("A1:C1").Value = "Header"
    WS.Range("A2:C4").Formula = "=ADDRESS(ROW(),COLUMN())"
    WS.Range("$A$1:$C$4").AutoFilter Field:=1, Criteria1:="$A$2"
    If GetWholeRange(WS).Address <> "$A$1:$C$4" Then
        TestGetWholeRange = False
        Debug.Print "Hidden"
    End If
    
    Excel.Application.DisplayAlerts = False
    WB.Close
    Excel.Application.DisplayAlerts = True
    
    Excel.Application.ScreenUpdating = True
    
    Debug.Print "TestGetWholeRange: " & TestGetWholeRange
    
End Function

Private Function TestGetDataRange() As Boolean
    
    TestGetDataRange = True
    
    Excel.Application.ScreenUpdating = False
    
    Dim WB As Workbook
    Set WB = Excel.Application.Workbooks.Add
    
    Dim WS As Worksheet
    Set WS = WB.Worksheets(1)
            
    'Blank
    If Not GetDataRange(WS) Is Nothing Then
        TestGetDataRange = False
        Debug.Print "Blank"
    End If
    
    'Single cell
        'Start 1, 1
        WS.Cells(1, 1).Value = "Test"
        If GetDataRange(WS).Address <> "$A$1" Then
            TestGetDataRange = False
            Debug.Print "Single cell Start 1, 1"
        End If
        'Start 2, 2
        WS.Cells.Clear
        WS.Cells(2, 2).Value = "Test"
        If GetDataRange(WS).Address <> "$B$2" Then
            TestGetDataRange = False
            Debug.Print "Single cell Start 2, 2"
        End If

    'Multiple cell
        'Start 1, 1
        WS.Cells.Clear
        WS.Range("A1:C1").Value = "Header"
        WS.Range("A2:C4").Formula = "=ADDRESS(ROW(),COLUMN())"
        If GetDataRange(WS).Address <> "$A$1:$C$4" Then
            TestGetDataRange = False
            Debug.Print "Multiple cell Start 1, 1"
        End If
        'Start 2, 2
        WS.Cells.Clear
        WS.Range("B2:D2").Value = "Header"
        WS.Range("B3:D5").Formula = "=ADDRESS(ROW(),COLUMN())"
        If GetDataRange(WS).Address <> "$B$2:$D$5" Then
            TestGetDataRange = False
            Debug.Print "Multiple cell Start 2, 2"
        End If
        
    'Hidden
    WS.Cells.Clear
    WS.Range("A1:C1").Value = "Header"
    WS.Range("A2:C4").Formula = "=ADDRESS(ROW(),COLUMN())"
    WS.Rows("1:4").Hidden = True
    WS.Columns("A:C").Hidden = True
    If GetDataRange(WS).Address <> "$A$1:$C$4" Then
        TestGetDataRange = False
        Debug.Print "Hidden"
    End If
    
    'Filtered
    WS.Rows.Hidden = False
    WS.Columns.Hidden = False
    WS.Cells.Clear
    WS.Range("A1:C1").Value = "Header"
    WS.Range("A2:C4").Formula = "=ADDRESS(ROW(),COLUMN())"
    WS.Range("$A$1:$C$4").AutoFilter Field:=1, Criteria1:="$A$2"
    If GetDataRange(WS).Address <> "$A$1:$C$4" Then
        TestGetDataRange = False
        Debug.Print "Hidden"
    End If
            
    Excel.Application.DisplayAlerts = False
    WB.Close
    Excel.Application.DisplayAlerts = True
    
    Excel.Application.ScreenUpdating = True
    
    Debug.Print "TestGetDataRange: " & TestGetDataRange
    
End Function

Private Function TestGetCharacterDictionary() As Boolean
    
    ' * Requires Microsoft Scripting Runtime Library
    
    TestGetCharacterDictionary = True
    
    Excel.Application.ScreenUpdating = False
    
    Dim WB As Workbook
    Set WB = Excel.Application.Workbooks.Add
    
    Dim WS As Worksheet
    Set WS = WB.Worksheets(1)
            
    Dim CD As Object
            
    'Blank
    Set CD = GetCharacterDictionary(WS)
    If CD.Count <> 0 Then
        TestGetCharacterDictionary = False
        Debug.Print "Blank"
    End If
    
    'Single
    WS.Cells(1, 1).Value = "A"
    Set CD = GetCharacterDictionary(WS)
    If Join(CD.Keys(), "") <> "A" Then
        TestGetCharacterDictionary = False
        Debug.Print "Single keys"
    End If
    If CD("A") <> 1 Then
        TestGetCharacterDictionary = False
        Debug.Print "Single items"
    End If
    If CD.Count <> 1 Then
        TestGetCharacterDictionary = False
        Debug.Print "Single count"
    End If
    
    'Multiple
    WS.Cells(2, 1).Value = "ABC"
    Set CD = GetCharacterDictionary(WS)
    If Join(CD.Keys(), "") <> "ABC" Then
        TestGetCharacterDictionary = False
        Debug.Print "Multiple keys"
    End If
    If CD("A") <> 2 Then
        TestGetCharacterDictionary = False
        Debug.Print "Multiple items A"
    End If
    If CD("B") <> 1 Then
        TestGetCharacterDictionary = False
        Debug.Print "Multiple items B"
    End If
    If CD("C") <> 1 Then
        TestGetCharacterDictionary = False
        Debug.Print "Multiple items C"
    End If
    If CD.Count <> 3 Then
        TestGetCharacterDictionary = False
        Debug.Print "Multiple count"
    End If
    
    '32-127
    WS.Cells.Clear
    Dim i&
    Dim j&
    Dim s$
    For i = 32 To 127
        j = j + 1
        WS.Cells(j, 1).Formula = "=CHAR(" & i & ")"
    Next i
    Set CD = GetCharacterDictionary(WS)
    For i = 32 To 127
        If Not CD.Exists(Chr(i)) Then
            TestGetCharacterDictionary = False
            Debug.Print i
            Debug.Print "32-127 chars"
            Exit For
        End If
    Next i
    If CD.Count <> 96 Then
        Debug.Print CD.Count
        TestGetCharacterDictionary = False
        Debug.Print "32-127 count"
    End If
    
    Excel.Application.DisplayAlerts = False
    WB.Close
    Excel.Application.DisplayAlerts = True
    
    Excel.Application.ScreenUpdating = True
    
    Debug.Print "TestGetCharacterDictionary: " & TestGetCharacterDictionary
    
End Function

Private Function TestCreateWorkbookFromWorksheet() As Boolean
    
    TestCreateWorkbookFromWorksheet = True
    
    Excel.Application.ScreenUpdating = False
    
    Dim WB As Workbook
    Set WB = Excel.Application.Workbooks.Add
    
    Dim WS As Worksheet
    Set WS = WB.Worksheets(1)
       
    WS.Cells(1, 1).Value = "Test"
    
    'Workbook created
    Dim NB As Workbook
    Set NB = CreateWorkbookFromWorksheet(WS)
    If NB Is Nothing Then
        TestCreateWorkbookFromWorksheet = False
        Debug.Print "Workbook created"
    End If
    
    'Worksheet is copy
    If NB.Worksheets(1).Cells(1, 1).Value <> "Test" Then
        TestCreateWorkbookFromWorksheet = False
        Debug.Print "Worksheet is copy"
    End If
    
    Excel.Application.DisplayAlerts = False
    NB.Close
    WB.Close
    Excel.Application.DisplayAlerts = True
    
    Excel.Application.ScreenUpdating = True
    
    Debug.Print "TestCreateWorkbookFromWorksheet: " & TestCreateWorkbookFromWorksheet
    
End Function

Private Function TestJoinRange() As Boolean
    
    TestJoinRange = True
    
    Excel.Application.ScreenUpdating = False
    
    Dim WB As Workbook
    Set WB = Excel.Application.Workbooks.Add
    
    Dim WS As Worksheet
    Set WS = WB.Worksheets(1)
            
    Dim Rng As Range
    
    'Nothing
    Set Rng = JoinRange(Rng, Rng)
    If Not Rng Is Nothing Then
        TestJoinRange = False
        Debug.Print "Nothing"
    End If
    
    'Single
    Set Rng = JoinRange(Rng, WS.Range("A1"))
    If Rng.Address <> "$A$1" Then
        TestJoinRange = False
        Debug.Print "Single"
    End If
    
    'Single same range
    Set Rng = JoinRange(Rng, WS.Range("A1"))
    If Rng.Address <> "$A$1" Then
        TestJoinRange = False
        Debug.Print "Single same range"
    End If
    
    'Single overlapping range
    Set Rng = JoinRange(Rng, WS.Range("A1:B1"))
    If Rng.Address <> "$A$1:$B$1" Then
        TestJoinRange = False
        Debug.Print "Single overlapping range"
    End If
    
    'Multiple
        'Contiguous
        Set Rng = Nothing
        Set Rng = JoinRange(Rng, WS.Range("A1:C3"))
        Set Rng = JoinRange(Rng, WS.Range("A4:C8"))
        If Rng.Address <> "$A$1:$C$8" Then
            TestJoinRange = False
            Debug.Print "Multiple Contiguous"
        End If
        'Noncontiguous
        Set Rng = Nothing
        Set Rng = JoinRange(Rng, WS.Range("A1:C3"))
        Set Rng = JoinRange(Rng, WS.Range("A5:C8"))
        If Rng.Address <> "$A$1:$C$3,$A$5:$C$8" Then
            TestJoinRange = False
            Debug.Print "Multiple Noncontiguous"
        End If
        
    'Entire rows
        'Contiguous
        Set Rng = Nothing
        Set Rng = JoinRange(Rng, WS.Range("1:1"))
        Set Rng = JoinRange(Rng, WS.Range("2:2"))
        If Rng.Address <> "$1:$2" Then
            TestJoinRange = False
            Debug.Print "Entire rows Contiguous"
        End If
        'Noncontiguous
        Set Rng = Nothing
        Set Rng = JoinRange(Rng, WS.Range("1:1"))
        Set Rng = JoinRange(Rng, WS.Range("3:3"))
        If Rng.Address <> "$1:$1,$3:$3" Then
            TestJoinRange = False
            Debug.Print "Entire rows Noncontiguous"
        End If
        
    'Entire columns
        'Contiguous
        Set Rng = Nothing
        Set Rng = JoinRange(Rng, WS.Range("A:A"))
        Set Rng = JoinRange(Rng, WS.Range("B:B"))
        If Rng.Address <> "$A:$B" Then
            TestJoinRange = False
            Debug.Print "Entire columns Contiguous"
        End If
        'Noncontiguous
        Set Rng = Nothing
        Set Rng = JoinRange(Rng, WS.Range("A:A"))
        Set Rng = JoinRange(Rng, WS.Range("C:C"))
        If Rng.Address <> "$A:$A,$C:$C" Then
            TestJoinRange = False
            Debug.Print "Entire columns Noncontiguous"
        End If
        
    'Entire rows and columns
        'Contiguous
        Set Rng = Nothing
        Set Rng = JoinRange(Rng, WS.Range("A:A"))
        Set Rng = JoinRange(Rng, WS.Range("1:1"))
        If Rng.Address <> "$A:$A,$1:$1" Then
            TestJoinRange = False
            Debug.Print "Entire rows and columns Contiguous"
        End If
        'Noncontiguous
        Set Rng = Nothing
        Set Rng = JoinRange(Rng, WS.Range("A:A"))
        Set Rng = JoinRange(Rng, WS.Range("3:3"))
        If Rng.Address <> "$A:$A,$3:$3" Then
            TestJoinRange = False
            Debug.Print "Entire rows and columns Noncontiguous"
        End If

    'Whole sheet
    Set Rng = Nothing
    Set Rng = JoinRange(Rng, WS.Cells)
    If Rng.Address <> "$1:$" & WS.Rows.Count Then
        TestJoinRange = False
        Debug.Print "Whole sheet"
    End If

    'Different sheets
    Dim NS As Worksheet
    Set NS = WB.Worksheets.Add
    Set Rng = Nothing
    Set Rng = JoinRange(Rng, WS.Range("A:A"))
    On Error Resume Next
    Set Rng = JoinRange(Rng, NS.Range("B:B"))
    If Err.Number <> 1004 Then
        TestJoinRange = False
        Debug.Print "Different sheets"
    End If
    On Error GoTo 0

    'Different workbooks
    Dim NB As Workbook
    Set NB = Excel.Application.Workbooks.Add
    Set NS = NB.Worksheets(1)
    Set Rng = Nothing
    Set Rng = JoinRange(Rng, WS.Range("A:A"))
    On Error Resume Next
    Set Rng = JoinRange(Rng, NS.Range("B:B"))
    If Err.Number <> 1004 Then
        TestJoinRange = False
        Debug.Print "Different workbooks"
    End If
    On Error GoTo 0
    
    Excel.Application.DisplayAlerts = False
    NB.Close
    WB.Close
    Excel.Application.DisplayAlerts = True
    
    Excel.Application.ScreenUpdating = True
    
    Debug.Print "TestJoinRange: " & TestJoinRange
    
End Function

Private Function TestNameWorksheet() As Boolean
    
    TestNameWorksheet = True
    
    Excel.Application.ScreenUpdating = False
    
    Dim WB As Workbook
    Set WB = Excel.Application.Workbooks.Add
    
    Dim WS As Worksheet
    Set WS = WB.Worksheets(1)
           
    Dim OriginalName As String
    OriginalName = WS.Name
    
    'Original sheet original name
    If NameWorksheet(WS, OriginalName) <> OriginalName Then
        TestNameWorksheet = False
        Debug.Print "Original sheet original name"
    End If
    
    'New sheet original name
    Dim NS1 As Worksheet
    Set NS1 = WB.Worksheets.Add
    If NameWorksheet(NS1, OriginalName) <> OriginalName & " (1)" Then
        TestNameWorksheet = False
        Debug.Print "New sheet original name"
    End If
    
    'Same sheet original name
    If NameWorksheet(NS1, OriginalName) <> OriginalName & " (1)" Then
        TestNameWorksheet = False
        Debug.Print "Same sheet original name"
    End If
            
    'Another new sheet original name
    Dim NS2 As Worksheet
    Set NS2 = WB.Worksheets.Add
    If NameWorksheet(NS2, OriginalName) <> OriginalName & " (2)" Then
        TestNameWorksheet = False
        Debug.Print "Another new sheet original name"
    End If
    
    Excel.Application.DisplayAlerts = False
    WB.Close
    Excel.Application.DisplayAlerts = True
    
    Excel.Application.ScreenUpdating = True
    
    Debug.Print "TestNameWorksheet: " & TestNameWorksheet
    
End Function

Private Function TestFunc$(s$)
    TestFunc = s & "Test"
End Function

Private Function TestJoinRangeText() As Boolean

    TestJoinRangeText = True
    
    Excel.Application.ScreenUpdating = False
    
    Dim WB As Workbook
    Set WB = Excel.Application.Workbooks.Add
    
    Dim WS As Worksheet
    Set WS = WB.Worksheets(1)
    
    WS.Range("A1:C3").Formula = "=SUBSTITUTE(ADDRESS(ROW(),COLUMN()),""$"","""")"
    WS.Range("B2").Clear
    
    'Row
        'Func
            'Ignore Blanks
            If JoinRangeText(WS.Range("A1:C3"), ",", True, "TestFunc") <> _
            "A1Test,B1Test,C1Test,A2Test,C2Test,A3Test,B3Test,C3Test" Then
                TestJoinRangeText = False
                Debug.Print "Row Func Ignore Blanks"
            End If
            'Dont Ignore Blanks
            If JoinRangeText(WS.Range("A1:C3"), ",", False, "TestFunc") <> _
            "A1Test,B1Test,C1Test,A2Test,Test,C2Test,A3Test,B3Test,C3Test" Then
                TestJoinRangeText = False
                Debug.Print "Row Func Dont Ignore Blanks"
            End If
        'No Func
            'Ignore Blanks
            If JoinRangeText(WS.Range("A1:C3"), ",", True) <> _
            "A1,B1,C1,A2,C2,A3,B3,C3" Then
                TestJoinRangeText = False
                Debug.Print "Row No Func Ignore Blanks"
            End If
            'Dont Ignore Blanks
            If JoinRangeText(WS.Range("A1:C3"), ",", False) <> _
            "A1,B1,C1,A2,,C2,A3,B3,C3" Then
                TestJoinRangeText = False
                Debug.Print "Row No Func Dont Ignore Blanks"
            End If
    
    'Column
        'Func
            'Ignore Blanks
            If JoinRangeText(WS.Range("A1:C3"), ",", True, "TestFunc", "Column") <> _
            "A1Test,A2Test,A3Test,B1Test,B3Test,C1Test,C2Test,C3Test" Then
                TestJoinRangeText = False
                Debug.Print "Column Func Ignore Blanks"
            End If
            'Dont Ignore Blanks
            If JoinRangeText(WS.Range("A1:C3"), ",", False, "TestFunc", "Column") <> _
            "A1Test,A2Test,A3Test,B1Test,Test,B3Test,C1Test,C2Test,C3Test" Then
                TestJoinRangeText = False
                Debug.Print "Column Func Dont Ignore Blanks"
            End If
        'No Func
            'Ignore Blanks
            If JoinRangeText(WS.Range("A1:C3"), ",", True, , "Column") <> _
            "A1,A2,A3,B1,B3,C1,C2,C3" Then
                TestJoinRangeText = False
                Debug.Print "Column No Func Ignore Blanks"
            End If
            'Dont Ignore Blanks
            If JoinRangeText(WS.Range("A1:C3"), ",", False, , "Column") <> _
            "A1,A2,A3,B1,,B3,C1,C2,C3" Then
                TestJoinRangeText = False
                Debug.Print "Column No Func Dont Ignore Blanks"
            End If
            
    Excel.Application.DisplayAlerts = False
    WB.Close
    Excel.Application.DisplayAlerts = True
    
    Excel.Application.ScreenUpdating = True
    
    Debug.Print "TestJoinRangeText: " & TestJoinRangeText
    
End Function
