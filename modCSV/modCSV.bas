Attribute VB_Name = "modCSV"
Attribute VB_Description = "Module to facilitate working with CSV."
Option Explicit

'Meta Data=============================================================
'======================================================================

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'  Copyright © 2021 Peter D Roach. All Rights Reserved.
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
'  Module Name: modCSV
'  Module Description: Facilitates working with CSV text and files.
'  Module Version: 1.0
'  Module License: MIT
'  Module Author: Peter Roach; PeterRoach.Code@gmail.com
'  Module Contents:
'  ----------------------------------------
'    Private Procedures:
'       GetNewLineCharacter
'       GetQuoteCharacter
'       TrimNewLineCharacters
'       ReplaceDelimiters
'       ResolveQuoteCharacters
'       SplitByDelimiter
'       GetDelimiterPositions *Not used
'       SplitByDelimiterPosition *Not used
'       ReadAllFromTextFile
'       WriteBufferToFile
'    Public API:
'       CSVToJaggedArray
'       CSVToStringArray
'       CSVFileToJaggedArray
'       CSVFileToStringArray
'       JaggedArrayToCSVString
'       StringArrayToCSVString
'       JaggedArrayToCSVFile
'       StringArrayToCSVFile
'    Unit Tests:
'       TestmodCSV
'       TestGetNewLineCharacter
'       TestGetQuoteCharacter
'       TestTrimNewLineCharacters
'       TestReplaceDelimiters
'       TestResolveQuoteCharacters
'       TestSplitByDelimiter
'       TestGetDelimiterPositions
'       TestSplitByDelimiterPosition
'       TestReadAllFromTextFileAndWriteBufferToFile
'       TestCSVToJaggedArray
'       TestCSVToStringArray
'       TestCSVFileToJaggedArray
'       TestCSVFileToStringArray
'       TestJaggedArrayToCSVString
'       TestStringArrayToCSVString
'       TestJaggedArrayToCSVFile
'       TestStringArrayToCSVFile
'  ----------------------------------------

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'  Example Usage:
'
'  Public Sub ExampleCSV()
'
'      'Jagged Array'''''''''''''''''''''''''''''''''''''''''''''''''''
'
'      Dim Arr()
'
'      'From file
'      Arr = CSVFileToJaggedArray("C:\example.csv", ",", _
'          NewLineCharacterCRLF, QuoteCharacterDouble)
'
'      'From string
'      Arr = CSVToJaggedArray("A,B,C", ",", _
'          NewLineCharacterCRLF, QuoteCharacterDouble)
'
'      'To string
'      Dim S1$
'      S1 = JaggedArrayToCSVString(Arr, ",", NewLineCharacterCRLF, _
'          QuoteCharacterDouble, QuoteOptionQuoteEmbedded)
'
'      'To file
'      JaggedArrayToCSVFile _
'          Arr, "C:\example1.csv", ",", _
'          NewLineCharacterCRLF, QuoteCharacterDouble, _
'          QuoteOptionQuoteEmbedded
'
'
'      'String Array'''''''''''''''''''''''''''''''''''''''''''''''''''
'
'      'From file
'      Dim SArr$()
'      SArr = CSVFileToStringArray("C:\example.csv", ",", _
'          NewLineCharacterCRLF, QuoteCharacterDouble)
'
'      'From string
'      SArr = CSVToStringArray("A,B,C", ",", _
'          NewLineCharacterCRLF, QuoteCharacterDouble)
'
'      'To string
'      Dim S2$
'      S2 = StringArrayToCSVString(SArr, ",", NewLineCharacterCRLF, _
'          QuoteCharacterDouble, QuoteOptionQuoteEmbedded)
'
'      'To file
'      StringArrayToCSVFile _
'          SArr, "C:\example1.csv", ",", _
'          NewLineCharacterCRLF, QuoteCharacterDouble, _
'          QuoteOptionQuoteEmbedded
'
'  End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Public Constants======================================================
'======================================================================

Public Enum NewLineCharacter
    NewLineCharacterDefault = 1
    NewLineCharacterLF = 2
    NewLineCharacterCR = 3
    NewLineCharacterCRLF = 4
End Enum

Public Enum QuoteCharacter
    QuoteCharacterNone = 1
    QuoteCharacterSingle = 2
    QuoteCharacterDouble = 3
End Enum

Public Enum QuoteOption
    QuoteOptionQuoteNone = 1
    QuoteOptionQuoteEmbedded = 2
    QuoteOptionQuoteNonNumeric = 3
    QuoteOptionQuoteAll = 4
End Enum


'Private Constants=====================================================
'======================================================================

Private Enum CSVState
    CSVStateNormal = 1
    CSVStateQuoted = 2
End Enum

Private Const REPLACEMENT_CHARACTER As String = "" 'Unit Separator Chr(31)


'Private Functions=====================================================
'======================================================================

Private Function GetNewLineCharacter$(NewLineChar As NewLineCharacter)
    Select Case NewLineChar
        Case NewLineCharacterDefault: GetNewLineCharacter = vbNewLine
        Case NewLineCharacterLF: GetNewLineCharacter = vbLf
        Case NewLineCharacterCR: GetNewLineCharacter = vbCr
        Case NewLineCharacterCRLF: GetNewLineCharacter = vbCrLf
    End Select
End Function

Private Function GetQuoteCharacter$(QuoteChar As QuoteCharacter)
    Select Case QuoteChar
        Case QuoteCharacterNone: GetQuoteCharacter = ""
        Case QuoteCharacterSingle: GetQuoteCharacter = "'"
        Case QuoteCharacterDouble: GetQuoteCharacter = """"
    End Select
End Function

Private Function TrimNewLineCharacters$(Text$, NLChar$)
    Dim NLLen&
    Dim L&
    Dim R&
    NLLen = Len(NLChar)
    L = 1
    R = Len(Text)
    If R < NLLen Then
        TrimNewLineCharacters = Text
        Exit Function
    End If
    If Replace(Text, NLChar, "") = "" Then
        TrimNewLineCharacters = ""
        Exit Function
    End If
    Do While Mid$(Text, R - NLLen + 1, NLLen) = NLChar
        R = R - NLLen
    Loop
    Do While Mid$(Text, L, NLLen) = NLChar
        L = L + NLLen
    Loop
    TrimNewLineCharacters = Mid$(Text, L, R - L + 1)
End Function

Private Sub ReplaceDelimiters(Text$, Delimiter$, QuoteChar$)
    Dim State As CSVState: State = CSVStateNormal
    Dim i&
    For i = 1 To Len(Text)
        Select Case State
            Case CSVStateNormal
                Select Case Mid$(Text, i, 1)
                    Case Delimiter: Mid$(Text, i, 1) = _
                                    REPLACEMENT_CHARACTER
                    Case QuoteChar: State = CSVStateQuoted
                End Select
            Case CSVStateQuoted
                Select Case Mid$(Text, i, 1)
                    Case QuoteChar: State = CSVStateNormal
                End Select
        End Select
    Next i
End Sub

Private Sub ResolveQuoteCharacters(Units$(), QuoteChar$)
    Dim Unit$
    Dim i&
    If QuoteChar <> "" Then
        For i = LBound(Units) To UBound(Units)
            Unit = Units(i)
            If Len(Unit) > 1 Then
                If Left$(Unit, 1) = QuoteChar And _
                Right$(Unit, 1) = QuoteChar Then
                    If Len(Unit) > 2 Then
                        Units(i) = _
                        Mid$(Unit, 2, Len(Unit) - 2)
                    Else
                        Units(i) = ""
                    End If
                End If
                Units(i) = _
                Replace$(Units(i), _
                String$(2, QuoteChar), QuoteChar)
            End If
        Next i
    End If
End Sub

Private Function SplitByDelimiter(Text$, Delimiter$, QuoteChar$) _
As String()
    Dim Units$()
    If Len(Text) = 0 Then
        ReDim Units$(0 To 0)
        SplitByDelimiter = Units
        Exit Function
    End If
    Call ReplaceDelimiters(Text, Delimiter, QuoteChar)
    Units = Split(Text, REPLACEMENT_CHARACTER)
    Call ResolveQuoteCharacters(Units, QuoteChar)
    SplitByDelimiter = Units
End Function

Private Function GetDelimiterPositions(Text$, Delimiter$, QuoteChar$) _
As Collection
    Dim State As CSVState
    Dim i&
    State = CSVStateNormal
    Set GetDelimiterPositions = New Collection
    For i = 1 To Len(Text)
        Select Case State
            Case CSVStateNormal
                Select Case Mid$(Text, i, 1)
                    Case Delimiter: GetDelimiterPositions.Add i
                    Case QuoteChar: State = CSVStateQuoted
                End Select
            Case CSVStateQuoted
                Select Case Mid$(Text, i, 1)
                    Case QuoteChar: State = CSVStateNormal
                End Select
        End Select
    Next i
    GetDelimiterPositions.Add Len(Text) + 1
End Function

Private Function SplitByDelimiterPosition(Text$, Delimiter$, QuoteChar$) _
As String()
    Dim Units$()
    Dim i&
    Dim c&
    Dim DelimiterPositions As Collection
    Set DelimiterPositions = _
    GetDelimiterPositions(Text, Delimiter, QuoteChar)
    ReDim Units(0 To DelimiterPositions.Count - 1) As String
    c = 1
    For i = 1 To DelimiterPositions.Count
        Units(i - 1) = Mid$(Text, c, DelimiterPositions.Item(i) - c)
        c = DelimiterPositions.Item(i) + 1
    Next i
    Set DelimiterPositions = Nothing
    ResolveQuoteCharacters Units, QuoteChar
    SplitByDelimiterPosition = Units
End Function

Private Function ReadAllFromTextFile$(FilePath$)
    If Dir(FilePath) = "" Then
        Err.Raise 53
    End If
    Dim FF As Integer
    FF = FreeFile()
    Open FilePath For Binary Access Read Lock Write As #FF
    ReadAllFromTextFile = Space$(LOF(FF))
    Get #FF, , ReadAllFromTextFile
    Close #FF
End Function

Private Function WriteBufferToFile(FilePath$, Buffer$)
    Dim FF As Integer
    FF = FreeFile
    Open FilePath For Binary Access Write Lock Read Write As #FF
    Put #FF, , Buffer
    Close #FF
End Function


'Public Functions======================================================
'======================================================================

Public Function CSVToJaggedArray(CSVData$, _
Optional Delimiter$ = ",", _
Optional NewLineChar As NewLineCharacter = NewLineCharacterCRLF, _
Optional QuoteChar As QuoteCharacter = QuoteCharacterDouble) As Variant()
Attribute CSVToJaggedArray.VB_Description = "Converts a CSV string to a jagged array."
    
    Dim OutArr() As Variant
    
    If Len(Delimiter) <> 1 Then
        Err.Raise 5
    End If
    
    If Len(CSVData) = 0 Then
        ReDim OutArr(0 To 0) As Variant
        Dim EmptyArr$(0 To 0)
        OutArr(0) = EmptyArr
        CSVToJaggedArray = OutArr
        Exit Function
    End If
    
    Dim NLChar$
    NLChar = GetNewLineCharacter(NewLineChar)
    
    Dim QChar As String * 1
    QChar = GetQuoteCharacter(QuoteChar)
    
    Dim CSVData1$
    CSVData1 = TrimNewLineCharacters(CSVData, NLChar)
    
    Dim Lines$()
    Lines = Split(CSVData1, NLChar)
    
    Dim LineCount&
    LineCount = UBound(Lines) - LBound(Lines) + 1
    
    ReDim OutArr(0 To LineCount - 1)
    
    Dim i&
    If QuoteChar = QuoteCharacterNone Then
        For i = LBound(Lines) To UBound(Lines)
            OutArr(i) = Split((Lines(i)), Delimiter)
        Next i
    Else
        For i = LBound(Lines) To UBound(Lines)
            OutArr(i) = _
            SplitByDelimiter((Lines(i)), Delimiter, QChar)
        Next i
    End If
    
    CSVToJaggedArray = OutArr
    
End Function

Public Function CSVToStringArray(CSVData$, _
Optional Delimiter$ = ",", _
Optional NewLineChar As NewLineCharacter = NewLineCharacterCRLF, _
Optional QuoteChar As QuoteCharacter = QuoteCharacterDouble) As String()
Attribute CSVToStringArray.VB_Description = "Converts a CSV string to a 2D string array."
    
    Dim OutArr$()
    
    If Len(Delimiter) <> 1 Then
        Err.Raise 5
    End If
    
    If Len(CSVData) = 0 Then
        ReDim OutArr$(0 To 0, 0 To 0)
        OutArr(0, 0) = ""
        CSVToStringArray = OutArr
        Exit Function
    End If
    
    Dim NLChar$
    NLChar = GetNewLineCharacter(NewLineChar)
    
    Dim QChar As String * 1
    QChar = GetQuoteCharacter(QuoteChar)
    
    Dim CSVData1$
    CSVData1 = TrimNewLineCharacters(CSVData, NLChar)
    
    Dim Lines$()
    Lines = Split(CSVData1, NLChar)
    
    Dim LineCount&
    LineCount = UBound(Lines) - LBound(Lines) + 1
    
    Dim TempArr$()
    If QuoteChar = QuoteCharacterNone Then
        TempArr = Split((Lines(LBound(Lines))), Delimiter)
    Else
        TempArr = _
        SplitByDelimiter((Lines(LBound(Lines))), Delimiter, QChar)
    End If
    
    Dim UnitCount&
    UnitCount = UBound(TempArr) - LBound(TempArr) + 1
    
    ReDim OutArr$(0 To LineCount - 1, 0 To UnitCount - 1)
    
    Dim i&
    Dim j&
    Dim k&
    If QuoteChar = QuoteCharacterNone Then
        For i = LBound(Lines) To UBound(Lines)
            TempArr = Split((Lines(i)), Delimiter)
            k = UBound(TempArr) - LBound(TempArr) + 1
            If k <> UnitCount Then
                Err.Raise 5
            End If
            For j = LBound(TempArr) To UBound(TempArr)
                OutArr(i, j) = TempArr(j)
            Next j
        Next i
    Else
        For i = LBound(Lines) To UBound(Lines)
            TempArr = _
            SplitByDelimiter((Lines(i)), Delimiter, QChar)
            k = UBound(TempArr) - LBound(TempArr) + 1
            If k <> UnitCount Then
                Err.Raise 5
            End If
            For j = LBound(TempArr) To UBound(TempArr)
                OutArr(i, j) = TempArr(j)
            Next j
        Next i
    End If
    
    CSVToStringArray = OutArr

End Function

Public Function CSVFileToJaggedArray(FilePath$, _
Optional Delimiter$ = ",", _
Optional NewLineChar As NewLineCharacter = NewLineCharacterCRLF, _
Optional QuoteChar As QuoteCharacter = QuoteCharacterDouble) As Variant()
Attribute CSVFileToJaggedArray.VB_Description = "Gets a CSV string from a file and converts it to a jagged array."
    
    Dim FileContents$
    FileContents = ReadAllFromTextFile(FilePath$)
    
    CSVFileToJaggedArray = _
    CSVToJaggedArray(FileContents, Delimiter, NewLineChar, QuoteChar)
    
End Function

Public Function CSVFileToStringArray(FilePath$, _
Optional Delimiter$ = ",", _
Optional NewLineChar As NewLineCharacter = NewLineCharacterCRLF, _
Optional QuoteChar As QuoteCharacter = QuoteCharacterDouble) As String()
Attribute CSVFileToStringArray.VB_Description = "Gets a CSV string from a file and converts it to a 2D string array."

    Dim FileContents$
    FileContents = ReadAllFromTextFile(FilePath$)
        
    CSVFileToStringArray = _
    CSVToStringArray(FileContents, Delimiter, NewLineChar, QuoteChar)
    
End Function

Public Function JaggedArrayToCSVString$(CSVData(), _
Optional Delimiter$ = ",", _
Optional NewLineChar As NewLineCharacter = NewLineCharacterCRLF, _
Optional QuoteChar As QuoteCharacter = QuoteCharacterDouble, _
Optional QuoteOpt As QuoteOption = QuoteOptionQuoteEmbedded)
Attribute JaggedArrayToCSVString.VB_Description = "Converts a jagged array to a CSV string."
    
    If Len(Delimiter) <> 1 Then
        Err.Raise 5
    End If
    
    Dim CSVData1()
    CSVData1 = CSVData
    
    Dim Q$
    Q = GetQuoteCharacter(QuoteChar)
    
    Dim i As Long
    Dim j As Long
    Select Case QuoteOpt
        Case QuoteOptionQuoteEmbedded
            For i = LBound(CSVData1) To UBound(CSVData1)
                For j = LBound(CSVData1(i)) To UBound(CSVData1(i))
                    If InStr(CSVData1(i)(j), Delimiter) > 0 Then
                        CSVData1(i)(j) = Q & CSVData1(i)(j) & Q
                    End If
                Next j
            Next i
        Case QuoteOptionQuoteNonNumeric
            For i = LBound(CSVData1) To UBound(CSVData1)
                For j = LBound(CSVData1(i)) To UBound(CSVData1(i))
                    If Not IsNumeric(CSVData1(i)(j)) Then
                        CSVData1(i)(j) = Q & CSVData1(i)(j) & Q
                    End If
                Next j
            Next i
        Case QuoteOptionQuoteAll
            For i = LBound(CSVData1) To UBound(CSVData1)
                For j = LBound(CSVData1(i)) To UBound(CSVData1(i))
                    CSVData1(i)(j) = Q & CSVData1(i)(j) & Q
                Next j
            Next i
    End Select
    
    Dim Arr$()
    ReDim Arr(LBound(CSVData1) To UBound(CSVData1))

    For i = LBound(CSVData1) To UBound(CSVData1)
        Arr(i) = Join(CSVData1(i), Delimiter)
    Next i
    
    JaggedArrayToCSVString = _
    Join(Arr, GetNewLineCharacter(NewLineChar))
    
End Function

Public Function StringArrayToCSVString$(CSVData$(), _
Optional Delimiter$ = ",", _
Optional NewLineChar As NewLineCharacter = NewLineCharacterCRLF, _
Optional QuoteChar As QuoteCharacter = QuoteCharacterDouble, _
Optional QuoteOpt As QuoteOption = QuoteOptionQuoteEmbedded)
Attribute StringArrayToCSVString.VB_Description = "Converts a 2D string array to a CSV string."
    
    If Len(Delimiter) <> 1 Then
        Err.Raise 5
    End If
    
    Dim Q$
    Q = GetQuoteCharacter(QuoteChar)
    
    Dim Arr$()
    ReDim Arr(LBound(CSVData, 1) To UBound(CSVData, 1))
    
    Dim TmpArr$()
    ReDim TmpArr$(LBound(CSVData, 2) To UBound(CSVData, 2))
    
    Dim i As Long
    Dim j As Long
    Dim S$
    Select Case QuoteOpt
        Case QuoteOptionQuoteNone
            For i = LBound(CSVData, 1) To UBound(CSVData, 1)
                For j = LBound(CSVData, 2) To UBound(CSVData, 2)
                    TmpArr(j) = CSVData(i, j)
                Next j
                Arr(i) = Join(TmpArr, Delimiter)
            Next i
        Case QuoteOptionQuoteNonNumeric
            For i = LBound(CSVData, 1) To UBound(CSVData, 1)
                For j = LBound(CSVData, 2) To UBound(CSVData, 2)
                    S = CSVData(i, j)
                    If Not IsNumeric(S) Then
                        TmpArr(j) = Q & S & Q
                    Else
                        TmpArr(j) = S
                    End If
                Next j
                Arr(i) = Join(TmpArr, Delimiter)
            Next i
        Case QuoteOptionQuoteEmbedded
            For i = LBound(CSVData, 1) To UBound(CSVData, 1)
                For j = LBound(CSVData, 2) To UBound(CSVData, 2)
                    S = CSVData(i, j)
                    If InStr(S, Delimiter) > 0 Then
                        TmpArr(j) = Q & S & Q
                    Else
                        TmpArr(j) = S
                    End If
                Next j
                Arr(i) = Join(TmpArr, Delimiter)
            Next i
        Case QuoteOptionQuoteAll
            For i = LBound(CSVData, 1) To UBound(CSVData, 1)
                For j = LBound(CSVData, 2) To UBound(CSVData, 2)
                    TmpArr(j) = Q & CSVData(i, j) & Q
                Next j
                Arr(i) = Join(TmpArr, Delimiter)
            Next i
    End Select
    
    StringArrayToCSVString = _
    Join(Arr, GetNewLineCharacter(NewLineChar))

End Function

Public Sub JaggedArrayToCSVFile(CSVData(), FilePath$, _
Optional Delimiter$ = ",", _
Optional NewLineChar As NewLineCharacter = NewLineCharacterCRLF, _
Optional QuoteChar As QuoteCharacter = QuoteCharacterDouble, _
Optional QuoteOpt As QuoteOption = QuoteOptionQuoteEmbedded)
Attribute JaggedArrayToCSVFile.VB_Description = "Converts a jagged array to a CSV string and writes the CSV string to a file."

    If Len(Delimiter) <> 1 Then
        Err.Raise 5
    End If
        
    Dim Buffer$
    Buffer = _
    JaggedArrayToCSVString(CSVData, Delimiter, NewLineChar, QuoteChar, QuoteOpt)
    
    WriteBufferToFile FilePath, Buffer
    
End Sub

Public Sub StringArrayToCSVFile(CSVData$(), FilePath$, _
Optional Delimiter$ = ",", _
Optional NewLineChar As NewLineCharacter = NewLineCharacterCRLF, _
Optional QuoteChar As QuoteCharacter = QuoteCharacterDouble, _
Optional QuoteOpt As QuoteOption = QuoteOptionQuoteEmbedded)
Attribute StringArrayToCSVFile.VB_Description = "Converts a 2D string array to a CSV string and writes the CSV string to a file."

    If Len(Delimiter) <> 1 Then
        Err.Raise 5
    End If
        
    Dim Buffer$
    Buffer = _
    StringArrayToCSVString(CSVData, Delimiter, NewLineChar, QuoteChar, QuoteOpt)
    
    WriteBufferToFile FilePath, Buffer
    
End Sub


'Unit Tests============================================================
'======================================================================

Private Function TestmodCSV() As Boolean
    
    TestmodCSV = _
        TestGetNewLineCharacter And _
        TestGetQuoteCharacter And _
        TestTrimNewLineCharacters And _
        TestReplaceDelimiters And _
        TestResolveQuoteCharacters And _
        TestSplitByDelimiter And _
        TestReadAllFromTextFileAndWriteBufferToFile And _
        TestCSVToJaggedArray And _
        TestCSVToStringArray And _
        TestCSVFileToJaggedArray And _
        TestCSVFileToStringArray And _
        TestJaggedArrayToCSVString And _
        TestStringArrayToCSVString And _
        TestJaggedArrayToCSVFile And _
        TestStringArrayToCSVFile
    
    Debug.Print "TestmodCSV: " & TestmodCSV
    
End Function

Private Function TestGetNewLineCharacter() As Boolean
        
    TestGetNewLineCharacter = True
    
    'Default (CrLf on Windows, Lf on Mac)
    #If Mac Then
    If GetNewLineCharacter(NewLineCharacterDefault) <> vbLf Then
    #Else
    If GetNewLineCharacter(NewLineCharacterDefault) <> vbCrLf Then
    #End If
        TestGetNewLineCharacter = False
        Debug.Print "Default"
    End If
    
    'Lf
    If GetNewLineCharacter(NewLineCharacterLF) <> vbLf Then
        TestGetNewLineCharacter = False
        Debug.Print "Lf"
    End If
    
    'Cr
    If GetNewLineCharacter(NewLineCharacterCR) <> vbCr Then
        TestGetNewLineCharacter = False
        Debug.Print "Cr"
    End If
    
    'CrLf
    If GetNewLineCharacter(NewLineCharacterCRLF) <> vbCrLf Then
        TestGetNewLineCharacter = False
        Debug.Print "CrLf"
    End If
    
    'Invalid lower
    If GetNewLineCharacter(0) <> "" Then
        TestGetNewLineCharacter = False
        Debug.Print "Invalid lower"
    End If
    
    'Invalid upper
    If GetNewLineCharacter(5) <> "" Then
        TestGetNewLineCharacter = False
        Debug.Print "Invalid upper"
    End If
    
    Debug.Print "TestGetNewLineCharacter: " & TestGetNewLineCharacter
    
End Function

Private Function TestGetQuoteCharacter() As Boolean

    TestGetQuoteCharacter = True

    'None
    If GetQuoteCharacter(QuoteCharacterNone) <> "" Then
        TestGetQuoteCharacter = False
        Debug.Print "None"
    End If
    
    'Double
    If GetQuoteCharacter(QuoteCharacterDouble) <> """" Then
        TestGetQuoteCharacter = False
        Debug.Print "Double"
    End If
    
    'Single
    If GetQuoteCharacter(QuoteCharacterSingle) <> "'" Then
        TestGetQuoteCharacter = False
        Debug.Print "Single"
    End If
      
    'Invalid lower
    If GetQuoteCharacter(0) <> "" Then
        TestGetQuoteCharacter = False
        Debug.Print "Invalid lower"
    End If
    
    'Invalid upper
    If GetQuoteCharacter(4) <> "" Then
        TestGetQuoteCharacter = False
        Debug.Print "Invalid upper"
    End If
    
    Debug.Print "TestGetQuoteCharacter: " & TestGetQuoteCharacter
    
End Function

Private Function TestTrimNewLineCharacters() As Boolean
    
    TestTrimNewLineCharacters = True
    
    '*****vbNewLine*****
    
    'Only
    If TrimNewLineCharacters(vbNewLine, vbNewLine) <> "" Then
        TestTrimNewLineCharacters = False
        Debug.Print "vbNewLine Only"
    End If
    
    'None
    If TrimNewLineCharacters("Hello", vbNewLine) <> "Hello" Then
        TestTrimNewLineCharacters = False
        Debug.Print "vbNewLine None"
    End If
    
    'Leading
    If TrimNewLineCharacters(vbNewLine & "Hello", vbNewLine) <> "Hello" Then
        TestTrimNewLineCharacters = False
        Debug.Print "vbNewLine Leading"
    End If
    
    'Trailing
    If TrimNewLineCharacters("Hello" & vbNewLine, vbNewLine) <> "Hello" Then
        TestTrimNewLineCharacters = False
        Debug.Print "vbNewLine Trailing"
    End If
    
    'Leading and Trailing
    If TrimNewLineCharacters(vbNewLine & "Hello" & vbNewLine, vbNewLine) <> "Hello" Then
        TestTrimNewLineCharacters = False
        Debug.Print "vbNewLine Leading and Trailing"
    End If
    
    'Multiple Leading and Trailing
    If TrimNewLineCharacters(vbNewLine & vbNewLine & "Hello" & vbNewLine & vbNewLine, vbNewLine) <> "Hello" Then
        TestTrimNewLineCharacters = False
        Debug.Print "vbNewLine Multiple Leading and Trailing "
    End If
    
    'Embedded
    If TrimNewLineCharacters("Hello" & vbNewLine & "Hello", vbNewLine) <> "Hello" & vbNewLine & "Hello" Then
        TestTrimNewLineCharacters = False
        Debug.Print "vbNewLine Embedded"
    End If
    
    '*****vbCrLf*****
    
    'Only
    If TrimNewLineCharacters(vbCrLf, vbCrLf) <> "" Then
        TestTrimNewLineCharacters = False
        Debug.Print "vbCrLf Only"
    End If
    
    'None
    If TrimNewLineCharacters("Hello", vbCrLf) <> "Hello" Then
        TestTrimNewLineCharacters = False
        Debug.Print "vbCrLf None"
    End If
    
    'Leading
    If TrimNewLineCharacters(vbCrLf & "Hello", vbCrLf) <> "Hello" Then
        TestTrimNewLineCharacters = False
        Debug.Print "vbCrLf Leading"
    End If
    
    'Trailing
    If TrimNewLineCharacters("Hello" & vbCrLf, vbCrLf) <> "Hello" Then
        TestTrimNewLineCharacters = False
        Debug.Print "vbCrLf Trailing"
    End If
    
    'Leading and Trailing
    If TrimNewLineCharacters(vbCrLf & "Hello" & vbCrLf, vbCrLf) <> "Hello" Then
        TestTrimNewLineCharacters = False
        Debug.Print "vbCrLf Leading and Trailing"
    End If
    
    'Multiple Leading and Trailing
    If TrimNewLineCharacters(vbCrLf & vbCrLf & "Hello" & vbCrLf & vbCrLf, vbCrLf) <> "Hello" Then
        TestTrimNewLineCharacters = False
        Debug.Print "vbCrLf Multiple Leading and Trailing"
    End If
    
    'Embedded
    If TrimNewLineCharacters("Hello" & vbCrLf & "Hello", vbCrLf) <> "Hello" & vbCrLf & "Hello" Then
        TestTrimNewLineCharacters = False
        Debug.Print "vbCrLf Embedded"
    End If
    
    '*****vbLf*****
    
    'Only
    If TrimNewLineCharacters(vbLf, vbLf) <> "" Then
        TestTrimNewLineCharacters = False
        Debug.Print "vbLf Only"
    End If
    
    'None
    If TrimNewLineCharacters("Hello", vbLf) <> "Hello" Then
        TestTrimNewLineCharacters = False
        Debug.Print "vbLf None"
    End If
    
    'Leading
    If TrimNewLineCharacters(vbLf & "Hello", vbLf) <> "Hello" Then
        TestTrimNewLineCharacters = False
        Debug.Print "vbLf Leading"
    End If
    
    'Trailing
    If TrimNewLineCharacters("Hello" & vbLf, vbLf) <> "Hello" Then
        TestTrimNewLineCharacters = False
        Debug.Print "vbLf Trailing"
    End If
    
    'Leading and Trailing
    If TrimNewLineCharacters(vbLf & "Hello" & vbLf, vbLf) <> "Hello" Then
        TestTrimNewLineCharacters = False
        Debug.Print "vbLf Leading and Trailing"
    End If
    
    'Multiple Leading and Trailing
    If TrimNewLineCharacters(vbLf & vbLf & "Hello" & vbLf & vbLf, vbLf) <> "Hello" Then
        TestTrimNewLineCharacters = False
        Debug.Print "vbLf Multiple Leading and Trailing"
    End If
    
    'Embedded
    If TrimNewLineCharacters("Hello" & vbLf & "Hello", vbLf) <> "Hello" & vbLf & "Hello" Then
        TestTrimNewLineCharacters = False
        Debug.Print "vbLf Embedded"
    End If
    
    '*****vbCr*****
    
    'Only
    If TrimNewLineCharacters(vbCr, vbCr) <> "" Then
        TestTrimNewLineCharacters = False
        Debug.Print "vbCr Only"
    End If
    
    'None
    If TrimNewLineCharacters("Hello", vbCr) <> "Hello" Then
        TestTrimNewLineCharacters = False
        Debug.Print "vbCr None"
    End If
    
    'Leading
    If TrimNewLineCharacters(vbCr & "Hello", vbCr) <> "Hello" Then
        TestTrimNewLineCharacters = False
        Debug.Print "vbCr Leading"
    End If
    
    'Trailing
    If TrimNewLineCharacters("Hello" & vbCr, vbCr) <> "Hello" Then
        TestTrimNewLineCharacters = False
        Debug.Print "vbCr Trailing"
    End If
    
    'Leading and Trailing
    If TrimNewLineCharacters(vbCr & "Hello" & vbCr, vbCr) <> "Hello" Then
        TestTrimNewLineCharacters = False
        Debug.Print "vbCr Leading and Trailing"
    End If
    
    'Multiple Leading and Trailing
    If TrimNewLineCharacters(vbCr & vbCr & "Hello" & vbCr & vbCr, vbCr) <> "Hello" Then
        TestTrimNewLineCharacters = False
        Debug.Print "vbCr Multiple Leading and Trailing"
    End If
    
    'Embedded
    If TrimNewLineCharacters("Hello" & vbCr & "Hello", vbCr) <> "Hello" & vbCr & "Hello" Then
        TestTrimNewLineCharacters = False
        Debug.Print "vbCr Embedded"
    End If
    
    Debug.Print "TestTrimNewLineCharacters: " & TestTrimNewLineCharacters
    
End Function

Private Function TestReplaceDelimiters() As Boolean

    TestReplaceDelimiters = True
    
    Dim S$
    
    'Delimiter Only
    S = ","
    ReplaceDelimiters S, ",", """"
    If S <> REPLACEMENT_CHARACTER Then
        TestReplaceDelimiters = False
        Debug.Print "Delimiter Only"
    End If
    
    'Delimiter Only Multiple
    S = ",,,"
    ReplaceDelimiters S, ",", """"
    If S <> String(3, REPLACEMENT_CHARACTER) Then
        TestReplaceDelimiters = False
        Debug.Print "Delimiter Only Multiple"
    End If
    
    'Leading Delimiter
    S = ",Hello"
    ReplaceDelimiters S, ",", """"
    If S <> REPLACEMENT_CHARACTER & "Hello" Then
        TestReplaceDelimiters = False
        Debug.Print "Leading Delimiter"
    End If
    
    'Trailing Delimiter
    S = "Hello,"
    ReplaceDelimiters S, ",", """"
    If S <> "Hello" & REPLACEMENT_CHARACTER Then
        TestReplaceDelimiters = False
        Debug.Print "Trailing Delimiter"
    End If
    
    'Delimited Text
    S = "Hello,Hello"
    ReplaceDelimiters S, ",", """"
    If S <> "Hello" & REPLACEMENT_CHARACTER & "Hello" Then
        TestReplaceDelimiters = False
        Debug.Print "Delimited Text"
    End If
    
    'Multiple Delimited Text
    S = "Hello,Hello,Hello"
    ReplaceDelimiters S, ",", """"
    If S <> "Hello" & REPLACEMENT_CHARACTER & _
    "Hello" & REPLACEMENT_CHARACTER & "Hello" Then
        TestReplaceDelimiters = False
        Debug.Print "Multiple Delimited Text"
    End If
    
    'Embedded Delimiter
    S = """Hello,Hello"""
    ReplaceDelimiters S, ",", """"
    If S <> """Hello,Hello""" Then
        TestReplaceDelimiters = False
        Debug.Print "Embdedded Delimiter"
    End If
    
    'No Delimiter
    S = "Hello"
    ReplaceDelimiters S, ",", """"
    If S <> "Hello" Then
        TestReplaceDelimiters = False
        Debug.Print "No Delimiter"
    End If
    
    'Empty String
    S = ""
    ReplaceDelimiters S, ",", """"
    If S <> "" Then
        TestReplaceDelimiters = False
        Debug.Print "Empty String"
    End If

    Debug.Print "TestReplaceDelimiters: " & TestReplaceDelimiters
    
End Function

Private Function TestResolveQuoteCharacters() As Boolean
    
    TestResolveQuoteCharacters = True
    
    Const DQ$ = """"
     
    Dim Units$(0 To 9)
    Units(0) = "Hello"                                                    ' Hello
    Units(1) = DQ & "Hello,World" & DQ                                    ' "Hello,World"
    Units(2) = String$(3, DQ) & "Hello,World" & String$(3, DQ)            ' """Hello,World"""
    Units(3) = String$(5, DQ) & "Hello,World" & String$(5, DQ)            ' """""Hello,World"""""
    Units(4) = DQ & "Hello," & DQ & DQ & "World" & DQ & DQ & DQ           ' "Hello,""World"""
    Units(5) = DQ & "Hello," & String$(4, DQ) & "World" & String$(5, DQ)  ' "Hello,""""World"""""
    Units(6) = DQ & "Hello" & DQ & DQ & "World" & DQ                      ' "Hello""World"
    Units(7) = ""                                                         '
    Units(8) = DQ & DQ                                                    ' ""
    Units(9) = DQ & DQ & DQ & DQ                                          ' """"
    
    ResolveQuoteCharacters Units, """"
    
    'No quotes
    If Units(0) <> "Hello" Then
        TestResolveQuoteCharacters = False
        Debug.Print "No quotes"
    End If
    
    'Quoted
    If Units(1) <> "Hello,World" Then
        TestResolveQuoteCharacters = False
        Debug.Print "Quoted"
    End If
    
    'Quoted with embedded quotes 1
    If Units(2) <> DQ & "Hello,World" & DQ Then
        TestResolveQuoteCharacters = False
        Debug.Print "Quoted with embedded quotes 1"
    End If
    
    'Quoted with double embedded quotes 1
    If Units(3) <> DQ & DQ & "Hello,World" & DQ & DQ Then
        TestResolveQuoteCharacters = False
        Debug.Print "Quoted with double embedded quotes 1"
    End If
    
    'Quoted with embedded quotes 2
    If Units(4) <> "Hello," & DQ & "World" & DQ Then
        TestResolveQuoteCharacters = False
        Debug.Print "Quoted with embedded quotes 2"
    End If
    
    'Quoted with double embedded quotes 2
    If Units(5) <> "Hello," & DQ & DQ & "World" & DQ & DQ Then
        TestResolveQuoteCharacters = False
        Debug.Print "Quoted with double embedded quotes 2"
    End If
    
    
    'Quoted with single embedded quote
    If Units(6) <> "Hello" & DQ & "World" Then
        TestResolveQuoteCharacters = False
        Debug.Print "Quoted with single embedded quote"
    End If
    
    'Empty String
    If Units(7) <> "" Then
        TestResolveQuoteCharacters = False
        Debug.Print "Empty String"
    End If
    
    'Quoted Empty String
    If Units(8) <> "" Then
        TestResolveQuoteCharacters = False
        Debug.Print "Quoted Empty String"
    End If
    
    'Quoted Double Quote
    If Units(9) <> DQ Then
        TestResolveQuoteCharacters = False
        Debug.Print "Quoted Double Quote"
    End If
    
    Debug.Print "TestResolveQuoteCharacters: " & TestResolveQuoteCharacters
    
End Function

Private Function TestSplitByDelimiter() As Boolean
    
    TestSplitByDelimiter = True
    
    Const DQ$ = """"
    
    Dim S$
    Dim SArr$()
    
    'Normal
    S = "Hello,Hello,Hello"
    SArr = SplitByDelimiter(S, ",", """")
    If SArr(0) <> "Hello" Or _
    SArr(1) <> "Hello" Or _
    SArr(2) <> "Hello" Then
        TestSplitByDelimiter = False
        Debug.Print "Normal"
    End If
    
    'Embedded
    S = "Hello," & DQ & "Hello,World" & DQ & ",Hello"
    SArr = SplitByDelimiter(S, ",", """")
    If SArr(0) <> "Hello" Or _
    SArr(1) <> "Hello,World" Or _
    SArr(2) <> "Hello" Then
        TestSplitByDelimiter = False
        Debug.Print "Embedded"
    End If
    
    'Quoted embedded
    S = "Hello," & DQ & DQ & DQ & "Hello,World" & DQ & DQ & DQ & ",Hello"
    SArr = SplitByDelimiter(S, ",", """")
    If SArr(0) <> "Hello" Or _
    SArr(1) <> """Hello,World""" Or _
    SArr(2) <> "Hello" Then
        TestSplitByDelimiter = False
        Debug.Print "Quoted embedded"
    End If
    
    'Bad embdedded
    S = "Hello," & DQ & DQ & "Hello,World" & DQ & DQ & ",Hello"
    SArr = SplitByDelimiter(S, ",", """")
    If SArr(0) <> "Hello" Or _
    SArr(1) <> """Hello" Or _
    SArr(2) <> "World""" Or _
    SArr(3) <> "Hello" Then
        TestSplitByDelimiter = False
        Debug.Print "Bad embdedded"
    End If
    
    'Single Delimiter only
    S = ","
    SArr = SplitByDelimiter(S, ",", """")
    If SArr(0) <> "" Or _
    SArr(1) <> "" Then
        TestSplitByDelimiter = False
        Debug.Print "Single Delimiter only"
    End If
    
    'Multiple Delimiters only
    S = ",,"
    SArr = SplitByDelimiter(S, ",", """")
    If SArr(0) <> "" Or _
    SArr(1) <> "" Or _
    SArr(2) <> "" Then
        TestSplitByDelimiter = False
        Debug.Print "Multiple Delimiters only"
    End If

    'Empty string
    S = ""
    SArr = SplitByDelimiter(S, ",", """")
    If SArr(0) <> "" Then
        TestSplitByDelimiter = False
        Debug.Print "Empty string"
    End If
    
    Debug.Print "TestSplitByDelimiter: " & TestSplitByDelimiter
    
End Function

Private Function TestGetDelimiterPositions() As Boolean
    
    TestGetDelimiterPositions = True
    
    Dim DPs As Collection
    
    'Empty string
    Set DPs = GetDelimiterPositions("", ",", """")
    If DPs(1) <> 1 Then
        TestGetDelimiterPositions = False
        Debug.Print "Empty String"
    End If
    
    'No delimiter
    Set DPs = GetDelimiterPositions("Hello", ",", """")
    If DPs(1) <> 6 Then
        TestGetDelimiterPositions = False
        Debug.Print "No Delimiter"
    End If
    
    'Delimiter only
    Set DPs = GetDelimiterPositions(",", ",", """")
    If DPs(1) <> 1 Or DPs(2) <> 2 Then
        TestGetDelimiterPositions = False
        Debug.Print "Delimiter only"
    End If
    
    'Delimiter only multiple
    Set DPs = GetDelimiterPositions(",,", ",", """")
    If DPs(1) <> 1 Or DPs(2) <> 2 Or DPs(3) <> 3 Then
        TestGetDelimiterPositions = False
        Debug.Print "Delimiter only multiple"
    End If
    
    'Normal single
    Set DPs = GetDelimiterPositions("1,2,3", ",", """")
    If DPs(1) <> 2 Or DPs(2) <> 4 Or DPs(3) <> 6 Then
        TestGetDelimiterPositions = False
        Debug.Print "Normal single"
    End If
    
    'Normal multiple
    Set DPs = GetDelimiterPositions("Hello,Hello,Hello", ",", """")
    If DPs(1) <> 6 Or DPs(2) <> 12 Or DPs(3) <> 18 Then
        TestGetDelimiterPositions = False
        Debug.Print "Normal multiple"
    End If
    
    'Embdedded
    Set DPs = GetDelimiterPositions("A,""B,C"",D", ",", """")
    If DPs(1) <> 2 Or DPs(2) <> 8 Or DPs(3) <> 10 Then
        TestGetDelimiterPositions = False
        Debug.Print "Embdedded"
    End If
    
    Debug.Print "TestGetDelimiterPositions: " & TestGetDelimiterPositions
            
End Function

Private Function TestSplitByDelimiterPosition() As Boolean

    TestSplitByDelimiterPosition = True

    Dim S$
    Dim SArr$()
    
    Const DQ$ = """"
    
    'Normal
    S = "Hello,Hello,Hello"
    SArr = SplitByDelimiterPosition(S, ",", """")
    If SArr(0) <> "Hello" Or _
    SArr(1) <> "Hello" Or _
    SArr(2) <> "Hello" Then
        TestSplitByDelimiterPosition = False
        Debug.Print "Normal"
    End If
    
    'Embedded
    S = "Hello," & DQ & "Hello,World" & DQ & ",Hello"
    SArr = SplitByDelimiterPosition(S, ",", """")
    If SArr(0) <> "Hello" Or _
    SArr(1) <> "Hello,World" Or _
    SArr(2) <> "Hello" Then
        TestSplitByDelimiterPosition = False
        Debug.Print "Embedded"
    End If
    
    'Quoted embedded
    S = "Hello," & DQ & DQ & DQ & "Hello,World" & DQ & DQ & DQ & ",Hello"
    SArr = SplitByDelimiterPosition(S, ",", """")
    If SArr(0) <> "Hello" Or _
    SArr(1) <> """Hello,World""" Or _
    SArr(2) <> "Hello" Then
        TestSplitByDelimiterPosition = False
        Debug.Print "Quoted embedded"
    End If
    
    'Bad embdedded
    S = "Hello," & DQ & DQ & "Hello,World" & DQ & DQ & ",Hello"
    SArr = SplitByDelimiterPosition(S, ",", """")
    If SArr(0) <> "Hello" Or _
    SArr(1) <> """Hello" Or _
    SArr(2) <> "World""" Or _
    SArr(3) <> "Hello" Then
        TestSplitByDelimiterPosition = False
        Debug.Print "Bad embdedded"
    End If
    
    'Single Delimiter only
    S = ","
    SArr = SplitByDelimiterPosition(S, ",", """")
    If SArr(0) <> "" Or _
    SArr(1) <> "" Then
        TestSplitByDelimiterPosition = False
        Debug.Print "Single Delimiter only"
    End If
    
    'Multiple Delimiters only
    S = ",,"
    SArr = SplitByDelimiterPosition(S, ",", """")
    If SArr(0) <> "" Or _
    SArr(1) <> "" Or _
    SArr(2) <> "" Then
        TestSplitByDelimiterPosition = False
        Debug.Print "Multiple Delimiters only"
    End If

    'Empty string
    S = ""
    SArr = SplitByDelimiterPosition(S, ",", """")
    If SArr(0) <> "" Then
        TestSplitByDelimiterPosition = False
        Debug.Print "Empty string"
    End If
    
    Debug.Print "TestSplitByDelimiterPosition: " & TestSplitByDelimiterPosition
            
End Function

Private Function TestReadAllFromTextFileAndWriteBufferToFile() As Boolean
    
    TestReadAllFromTextFileAndWriteBufferToFile = True
    
    'Write
    Dim TestCSVFilePath$
    TestCSVFilePath = _
    Environ$("USERPROFILE") & "\Desktop\example" & Format$(Now, "mmddyyyyhhmmss") & ".csv"
    Dim S$
    S = "Header1,Header2,Header3" & vbNewLine & _
    "1,2,3" & vbNewLine & _
    "4,5,6" & vbNewLine & _
    "7,8,9"
    WriteBufferToFile TestCSVFilePath, S
    
    'Read
    Dim S1$
    S1 = ReadAllFromTextFile(TestCSVFilePath)
    If S <> S1 Then
        TestReadAllFromTextFileAndWriteBufferToFile = False
        Debug.Print "String changed"
    End If
    
    Kill TestCSVFilePath
    
    'File not found
    Dim FakePath$
    FakePath = "C:\FakeFileThatDoesNotExist" & Format$(Now, "mmddyyyyhhmmss") & ".csv"
    On Error Resume Next
    S = ReadAllFromTextFile(FakePath)
    If Err.Number <> 53 Then
        TestReadAllFromTextFileAndWriteBufferToFile = False
        Debug.Print "File Not Found"
    End If
    On Error GoTo 0
    
    Debug.Print _
        "TestReadAllFromTextFileAndWriteBufferToFile: " & _
         TestReadAllFromTextFileAndWriteBufferToFile

End Function

Private Function TestCSVToJaggedArray() As Boolean
    
    TestCSVToJaggedArray = True
    
    Dim Arr()
    Dim S$
    
    'Different sized arrays
    S = "Header1,Header2,Header3" & vbNewLine & _
        "1,2,3" & vbNewLine & _
        "4,5,6,7"
    Arr = CSVToJaggedArray(S, ",", NewLineCharacterDefault, QuoteCharacterDouble)
    If Join(Arr(0), ",") <> "Header1,Header2,Header3" Or _
    Join(Arr(1), ",") <> "1,2,3" Or _
    Join(Arr(2), ",") <> "4,5,6,7" Then
        TestCSVToJaggedArray = False
        Debug.Print "Different sized arrays"
    End If
    
    'Same sized arrays
    S = "Header1,Header2,Header3" & vbNewLine & _
        "1,2,3" & vbNewLine & _
        "4,5,6" & vbNewLine & _
        "7,8,9"
    Arr = CSVToJaggedArray(S, ",", NewLineCharacterDefault, QuoteCharacterDouble)
    If Join(Arr(0), ",") <> "Header1,Header2,Header3" Or _
    Join(Arr(1), ",") <> "1,2,3" Or _
    Join(Arr(2), ",") <> "4,5,6" Or _
    Join(Arr(3), ",") <> "7,8,9" Then
        TestCSVToJaggedArray = False
        Debug.Print "Same sized arrays"
    End If
    
    'Empty string
    S = ""
    Arr = CSVToJaggedArray(S, ",", NewLineCharacterDefault, QuoteCharacterDouble)
    If Arr(0)(0) <> "" Then
        TestCSVToJaggedArray = False
        Debug.Print "Empty String"
    End If
    
    'No Delimiter
    S = "Header1,Header2,Header3" & vbNewLine & _
        "1,2,3" & vbNewLine & _
        "4,5,6" & vbNewLine & _
        "7,8,9"
    Arr = CSVToJaggedArray(S, vbTab, NewLineCharacterDefault, QuoteCharacterDouble)
    If Join(Arr(0), ",") <> "Header1,Header2,Header3" Or _
    Join(Arr(1), ",") <> "1,2,3" Or _
    Join(Arr(2), ",") <> "4,5,6" Or _
    Join(Arr(3), ",") <> "7,8,9" Then
        TestCSVToJaggedArray = False
        Debug.Print "No Delimiter"
    End If
    
    'Leading and Trailing New Line Characters
    S = vbNewLine & "Header1,Header2,Header3" & vbNewLine & _
        "1,2,3" & vbNewLine & _
        "4,5,6" & vbNewLine & _
        "7,8,9" & vbNewLine
    Arr = CSVToJaggedArray(S, ",", NewLineCharacterDefault, QuoteCharacterDouble)
    If Join(Arr(0), ",") <> "Header1,Header2,Header3" Or _
    Join(Arr(1), ",") <> "1,2,3" Or _
    Join(Arr(2), ",") <> "4,5,6" Or _
    Join(Arr(3), ",") <> "7,8,9" Then
        TestCSVToJaggedArray = False
        Debug.Print "Leading and Trailing New Line Characters"
    End If
    
    '*****New Line Characters*****
    
    'vbNewLine
    S = "Header1,Header2,Header3" & vbNewLine & _
        "1,2,3" & vbNewLine & _
        "4,5,6" & vbNewLine & _
        "7,8,9"
    Arr = CSVToJaggedArray(S, ",", NewLineCharacterDefault, QuoteCharacterDouble)
    If Join(Arr(0), ",") <> "Header1,Header2,Header3" Or _
    Join(Arr(1), ",") <> "1,2,3" Or _
    Join(Arr(2), ",") <> "4,5,6" Or _
    Join(Arr(3), ",") <> "7,8,9" Then
        TestCSVToJaggedArray = False
        Debug.Print "vbNewLine"
    End If
    
    'vbLf
    S = "Header1,Header2,Header3" & vbLf & _
        "1,2,3" & vbLf & _
        "4,5,6" & vbLf & _
        "7,8,9"
    Arr = CSVToJaggedArray(S, ",", NewLineCharacterLF, QuoteCharacterDouble)
    If Join(Arr(0), ",") <> "Header1,Header2,Header3" Or _
    Join(Arr(1), ",") <> "1,2,3" Or _
    Join(Arr(2), ",") <> "4,5,6" Or _
    Join(Arr(3), ",") <> "7,8,9" Then
        TestCSVToJaggedArray = False
        Debug.Print "vbLf"
    End If
    
    'vbCr
    S = "Header1,Header2,Header3" & vbCr & _
        "1,2,3" & vbCr & _
        "4,5,6" & vbCr & _
        "7,8,9"
    Arr = CSVToJaggedArray(S, ",", NewLineCharacterCR, QuoteCharacterDouble)
    If Join(Arr(0), ",") <> "Header1,Header2,Header3" Or _
    Join(Arr(1), ",") <> "1,2,3" Or _
    Join(Arr(2), ",") <> "4,5,6" Or _
    Join(Arr(3), ",") <> "7,8,9" Then
        TestCSVToJaggedArray = False
        Debug.Print "vbCr"
    End If
    
    'vbCrLf
    S = "Header1,Header2,Header3" & vbCrLf & _
        "1,2,3" & vbCrLf & _
        "4,5,6" & vbCrLf & _
        "7,8,9"
    Arr = CSVToJaggedArray(S, ",", NewLineCharacterCRLF, QuoteCharacterDouble)
    If Join(Arr(0), ",") <> "Header1,Header2,Header3" Or _
    Join(Arr(1), ",") <> "1,2,3" Or _
    Join(Arr(2), ",") <> "4,5,6" Or _
    Join(Arr(3), ",") <> "7,8,9" Then
        TestCSVToJaggedArray = False
        Debug.Print "vbCrLf"
    End If
    
    '*****Quote Characters*****
    
    'Double Quote
    S = "Header1,Header2,Header3" & vbNewLine & _
        "1,2,3" & vbNewLine & _
        "4,5,6" & vbNewLine & _
        "7,8,9"
    Arr = CSVToJaggedArray(S, ",", NewLineCharacterDefault, QuoteCharacterDouble)
    If Join(Arr(0), ",") <> "Header1,Header2,Header3" Or _
    Join(Arr(1), ",") <> "1,2,3" Or _
    Join(Arr(2), ",") <> "4,5,6" Or _
    Join(Arr(3), ",") <> "7,8,9" Then
        TestCSVToJaggedArray = False
        Debug.Print "Double Quote"
    End If
    
    'Single Quote
    S = "Header1,Header2,Header3" & vbNewLine & _
        "1,2,3" & vbNewLine & _
        "4,5,6" & vbNewLine & _
        "7,8,9"
    Arr = CSVToJaggedArray(S, ",", NewLineCharacterDefault, QuoteCharacterSingle)
    If Join(Arr(0), ",") <> "Header1,Header2,Header3" Or _
    Join(Arr(1), ",") <> "1,2,3" Or _
    Join(Arr(2), ",") <> "4,5,6" Or _
    Join(Arr(3), ",") <> "7,8,9" Then
        TestCSVToJaggedArray = False
        Debug.Print "Single Quote"
    End If
    
    'No Quote
    S = "Header1,Header2,Header3" & vbNewLine & _
        "1,2,3" & vbNewLine & _
        "4,5,6" & vbNewLine & _
        "7,8,9"
    Arr = CSVToJaggedArray(S, ",", NewLineCharacterDefault, QuoteCharacterNone)
    If Join(Arr(0), ",") <> "Header1,Header2,Header3" Or _
    Join(Arr(1), ",") <> "1,2,3" Or _
    Join(Arr(2), ",") <> "4,5,6" Or _
    Join(Arr(3), ",") <> "7,8,9" Then
        TestCSVToJaggedArray = False
        Debug.Print "No Quote"
    End If
    
    Debug.Print "TestCSVToJaggedArray: " & TestCSVToJaggedArray
    
End Function

Private Function TestCSVToStringArray() As Boolean
    
    TestCSVToStringArray = True
    
    Dim Arr$()
    Dim S$
    
    'Different sized arrays
    S = "Header1,Header2,Header3" & vbNewLine & _
        "1,2,3" & vbNewLine & _
        "4,5,6,7"
    On Error Resume Next
    Arr = CSVToStringArray(S, ",", NewLineCharacterDefault, QuoteCharacterDouble)
    If Err.Number <> 5 Then
        TestCSVToStringArray = False
        Debug.Print "Different Sized Arrays"
    End If
    On Error GoTo 0
    
    'Same sized arrays
    S = "Header1,Header2,Header3" & vbNewLine & _
        "1,2,3" & vbNewLine & _
        "4,5,6" & vbNewLine & _
        "7,8,9"
    Arr = CSVToStringArray(S, ",", NewLineCharacterDefault, QuoteCharacterDouble)
    If Arr(0, 0) & Arr(0, 1) & Arr(0, 2) <> "Header1Header2Header3" Or _
    Arr(1, 0) & Arr(1, 1) & Arr(1, 2) <> "123" Or _
    Arr(2, 0) & Arr(2, 1) & Arr(2, 2) <> "456" Or _
    Arr(3, 0) & Arr(3, 1) & Arr(3, 2) <> "789" Then
        TestCSVToStringArray = False
        Debug.Print "Same Sized Arrays"
    End If
    
    'Empty string
    S = ""
    Arr = CSVToStringArray(S, ",", NewLineCharacterDefault, QuoteCharacterDouble)
    If Arr(0, 0) <> "" Then
        TestCSVToStringArray = False
        Debug.Print "Empty String"
    End If
    
    'No Delimiter
    S = "Header1,Header2,Header3" & vbNewLine & _
        "1,2,3" & vbNewLine & _
        "4,5,6" & vbNewLine & _
        "7,8,9"
    Arr = CSVToStringArray(S, vbTab, NewLineCharacterDefault, QuoteCharacterDouble)
    If Arr(0, 0) <> "Header1,Header2,Header3" Or _
    Arr(1, 0) <> "1,2,3" Or _
    Arr(2, 0) <> "4,5,6" Or _
    Arr(3, 0) <> "7,8,9" Then
        TestCSVToStringArray = False
        Debug.Print "No Delimiter"
    End If
    
    'Leading and Trailing New Line Characters
    S = vbNewLine & "Header1,Header2,Header3" & vbNewLine & _
        "1,2,3" & vbNewLine & _
        "4,5,6" & vbNewLine & _
        "7,8,9" & vbNewLine
    Arr = CSVToStringArray(S, ",", NewLineCharacterDefault, QuoteCharacterDouble)
    If Arr(0, 0) & Arr(0, 1) & Arr(0, 2) <> "Header1Header2Header3" Or _
    Arr(1, 0) & Arr(1, 1) & Arr(1, 2) <> "123" Or _
    Arr(2, 0) & Arr(2, 1) & Arr(2, 2) <> "456" Or _
    Arr(3, 0) & Arr(3, 1) & Arr(3, 2) <> "789" Then
        TestCSVToStringArray = False
        Debug.Print "Leading and Trailing New Line Characters"
    End If
    
    '*****New Line Characters*****
    
    'vbNewLine
    S = "Header1,Header2,Header3" & vbNewLine & _
        "1,2,3" & vbNewLine & _
        "4,5,6" & vbNewLine & _
        "7,8,9"
    Arr = CSVToStringArray(S, ",", NewLineCharacterDefault, QuoteCharacterDouble)
    If Arr(0, 0) & Arr(0, 1) & Arr(0, 2) <> "Header1Header2Header3" Or _
    Arr(1, 0) & Arr(1, 1) & Arr(1, 2) <> "123" Or _
    Arr(2, 0) & Arr(2, 1) & Arr(2, 2) <> "456" Or _
    Arr(3, 0) & Arr(3, 1) & Arr(3, 2) <> "789" Then
        TestCSVToStringArray = False
        Debug.Print "vbNewLine"
    End If
    
    'vbLf
    S = "Header1,Header2,Header3" & vbLf & _
        "1,2,3" & vbLf & _
        "4,5,6" & vbLf & _
        "7,8,9"
    Arr = CSVToStringArray(S, ",", NewLineCharacterLF, QuoteCharacterDouble)
    If Arr(0, 0) & Arr(0, 1) & Arr(0, 2) <> "Header1Header2Header3" Or _
    Arr(1, 0) & Arr(1, 1) & Arr(1, 2) <> "123" Or _
    Arr(2, 0) & Arr(2, 1) & Arr(2, 2) <> "456" Or _
    Arr(3, 0) & Arr(3, 1) & Arr(3, 2) <> "789" Then
        TestCSVToStringArray = False
        Debug.Print "vbLf"
    End If
    
    'vbCr
    S = "Header1,Header2,Header3" & vbCr & _
        "1,2,3" & vbCr & _
        "4,5,6" & vbCr & _
        "7,8,9"
    Arr = CSVToStringArray(S, ",", NewLineCharacterCR, QuoteCharacterDouble)
    If Arr(0, 0) & Arr(0, 1) & Arr(0, 2) <> "Header1Header2Header3" Or _
    Arr(1, 0) & Arr(1, 1) & Arr(1, 2) <> "123" Or _
    Arr(2, 0) & Arr(2, 1) & Arr(2, 2) <> "456" Or _
    Arr(3, 0) & Arr(3, 1) & Arr(3, 2) <> "789" Then
        TestCSVToStringArray = False
        Debug.Print "vbCr"
    End If
    
    'vbCrLf
    S = "Header1,Header2,Header3" & vbCrLf & _
        "1,2,3" & vbCrLf & _
        "4,5,6" & vbCrLf & _
        "7,8,9"
    Arr = CSVToStringArray(S, ",", NewLineCharacterCRLF, QuoteCharacterDouble)
    If Arr(0, 0) & Arr(0, 1) & Arr(0, 2) <> "Header1Header2Header3" Or _
    Arr(1, 0) & Arr(1, 1) & Arr(1, 2) <> "123" Or _
    Arr(2, 0) & Arr(2, 1) & Arr(2, 2) <> "456" Or _
    Arr(3, 0) & Arr(3, 1) & Arr(3, 2) <> "789" Then
        TestCSVToStringArray = False
        Debug.Print "vbCrLf"
    End If
    
    '*****Quote Characters*****
    
    'Double Quote
    S = "Header1,Header2,Header3" & vbNewLine & _
        "1,2,3" & vbNewLine & _
        "4,5,6" & vbNewLine & _
        "7,8,9"
    Arr = CSVToStringArray(S, ",", NewLineCharacterDefault, QuoteCharacterDouble)
    If Arr(0, 0) & Arr(0, 1) & Arr(0, 2) <> "Header1Header2Header3" Or _
    Arr(1, 0) & Arr(1, 1) & Arr(1, 2) <> "123" Or _
    Arr(2, 0) & Arr(2, 1) & Arr(2, 2) <> "456" Or _
    Arr(3, 0) & Arr(3, 1) & Arr(3, 2) <> "789" Then
        TestCSVToStringArray = False
        Debug.Print "Double Quote"
    End If
    
    'Single Quote
    S = "Header1,Header2,Header3" & vbNewLine & _
        "1,2,3" & vbNewLine & _
        "4,5,6" & vbNewLine & _
        "7,8,9"
    Arr = CSVToStringArray(S, ",", NewLineCharacterDefault, QuoteCharacterSingle)
    If Arr(0, 0) & Arr(0, 1) & Arr(0, 2) <> "Header1Header2Header3" Or _
    Arr(1, 0) & Arr(1, 1) & Arr(1, 2) <> "123" Or _
    Arr(2, 0) & Arr(2, 1) & Arr(2, 2) <> "456" Or _
    Arr(3, 0) & Arr(3, 1) & Arr(3, 2) <> "789" Then
        TestCSVToStringArray = False
        Debug.Print "Single Quote"
    End If
    
    'No Quote
    S = "Header1,Header2,Header3" & vbNewLine & _
        "1,2,3" & vbNewLine & _
        "4,5,6" & vbNewLine & _
        "7,8,9"
    Arr = CSVToStringArray(S, ",", NewLineCharacterDefault, QuoteCharacterNone)
    If Arr(0, 0) & Arr(0, 1) & Arr(0, 2) <> "Header1Header2Header3" Or _
    Arr(1, 0) & Arr(1, 1) & Arr(1, 2) <> "123" Or _
    Arr(2, 0) & Arr(2, 1) & Arr(2, 2) <> "456" Or _
    Arr(3, 0) & Arr(3, 1) & Arr(3, 2) <> "789" Then
        TestCSVToStringArray = False
        Debug.Print "No Quote"
    End If
    
    Debug.Print "TestCSVToStringArray: " & TestCSVToStringArray
    
End Function

Private Function TestCSVFileToJaggedArray() As Boolean

    TestCSVFileToJaggedArray = True

    'Create Test CSV File
    Dim TestCSVFilePath$
    TestCSVFilePath = Environ$("USERPROFILE") & "\Desktop\example" & Format(Now, "mmddyyyyhhmmss") & ".csv"
    Dim TestCSVFileContents$
    TestCSVFileContents = _
        "Header1,Header2,Header3" & vbNewLine & _
        "1,2,3" & vbNewLine & _
        "4,5,6" & vbNewLine & _
        "7,8,9"
    WriteBufferToFile TestCSVFilePath, TestCSVFileContents

    'TestCSVFileToJaggedArray
    Dim Arr()
    Arr = CSVFileToJaggedArray(TestCSVFilePath, ",", NewLineCharacterDefault, QuoteCharacterDouble)
    If Join(Arr(0), ",") <> "Header1,Header2,Header3" Or _
    Join(Arr(1), ",") <> "1,2,3" Or _
    Join(Arr(2), ",") <> "4,5,6" Or _
    Join(Arr(3), ",") <> "7,8,9" Then
        TestCSVFileToJaggedArray = False
        Debug.Print "TestCSVFileToJaggedArray"
    End If

    Kill TestCSVFilePath

    Debug.Print "TestCSVFileToJaggedArray: " & TestCSVFileToJaggedArray

End Function

Private Function TestCSVFileToStringArray() As Boolean
    
    TestCSVFileToStringArray = True
    
    'Create Test CSV File
    Dim TestCSVFilePath$
    TestCSVFilePath = Environ$("USERPROFILE") & "\Desktop\example" & Format(Now, "mmddyyyyhhmmss") & ".csv"
    Dim TestCSVFileContents$
    TestCSVFileContents = _
        "Header1,Header2,Header3" & vbNewLine & _
        "1,2,3" & vbNewLine & _
        "4,5,6" & vbNewLine & _
        "7,8,9"
    WriteBufferToFile TestCSVFilePath, TestCSVFileContents
    
    'TestCSVFileToStringArray
    Dim Arr$()
    Arr = CSVFileToStringArray(TestCSVFilePath, ",", NewLineCharacterDefault, QuoteCharacterDouble)
    If Arr(0, 0) & Arr(0, 1) & Arr(0, 2) <> "Header1Header2Header3" Or _
    Arr(1, 0) & Arr(1, 1) & Arr(1, 2) <> "123" Or _
    Arr(2, 0) & Arr(2, 1) & Arr(2, 2) <> "456" Or _
    Arr(3, 0) & Arr(3, 1) & Arr(3, 2) <> "789" Then
        TestCSVFileToStringArray = False
        Debug.Print "TestCSVFileToStringArray"
    End If
    
    Kill TestCSVFilePath
    
    Debug.Print "TestCSVFileToStringArray: " & TestCSVFileToStringArray
    
End Function

Private Function TestJaggedArrayToCSVString() As Boolean

    TestJaggedArrayToCSVString = True

    Dim S$

    Dim Arr()
    ReDim Arr(0 To 3)
    Arr(0) = Array("Header1", "Header2", "Header3")
    Arr(1) = Array(1, 2, 3)
    Arr(2) = Array(4, 5, 6)
    Arr(3) = Array(7, 8, 9)
    
    'Normal
    S = JaggedArrayToCSVString(Arr, ",", NewLineCharacterCRLF, QuoteCharacterDouble, QuoteOptionQuoteEmbedded)
    If S <> "Header1,Header2,Header3" & vbCrLf & _
    "1,2,3" & vbCrLf & _
    "4,5,6" & vbCrLf & _
    "7,8,9" Then
        TestJaggedArrayToCSVString = False
        Debug.Print "Normal"
    End If
    
    'Empty Array
    ReDim Arr(0 To 0)
    Arr(0) = Array("")
    S = JaggedArrayToCSVString(Arr, ",", NewLineCharacterCRLF, QuoteCharacterDouble, QuoteOptionQuoteEmbedded)
    If S <> "" Then
        TestJaggedArrayToCSVString = False
        Debug.Print "Empty Array"
    End If
    
    ReDim Arr(0 To 3)
    Arr(0) = Array("Header1,Test", "Header2", "Header3")
    Arr(1) = Array(1, 2, 3)
    Arr(2) = Array(4, 5, 6)
    Arr(3) = Array(7, 8, 9)
    
    '*****New Line Character*****
    
    'vbCrLf
    S = JaggedArrayToCSVString(Arr, ",", NewLineCharacterCRLF, QuoteCharacterDouble, QuoteOptionQuoteEmbedded)
    If S <> """Header1,Test"",Header2,Header3" & vbCrLf & _
    "1,2,3" & vbCrLf & _
    "4,5,6" & vbCrLf & _
    "7,8,9" Then
        TestJaggedArrayToCSVString = False
        Debug.Print "vbCrLf"
    End If
    
    'vbCr
    S = JaggedArrayToCSVString(Arr, ",", NewLineCharacterCR, QuoteCharacterDouble, QuoteOptionQuoteEmbedded)
    If S <> """Header1,Test"",Header2,Header3" & vbCr & _
    "1,2,3" & vbCr & _
    "4,5,6" & vbCr & _
    "7,8,9" Then
        TestJaggedArrayToCSVString = False
        Debug.Print "vbCr"
    End If
    
    'vbLf
    S = JaggedArrayToCSVString(Arr, ",", NewLineCharacterLF, QuoteCharacterDouble, QuoteOptionQuoteEmbedded)
    If S <> """Header1,Test"",Header2,Header3" & vbLf & _
    "1,2,3" & vbLf & _
    "4,5,6" & vbLf & _
    "7,8,9" Then
        TestJaggedArrayToCSVString = False
        Debug.Print "vbLf"
    End If
    
    'Default
    S = JaggedArrayToCSVString(Arr, ",", NewLineCharacterDefault, QuoteCharacterDouble, QuoteOptionQuoteEmbedded)
    If S <> """Header1,Test"",Header2,Header3" & vbNewLine & _
    "1,2,3" & vbNewLine & _
    "4,5,6" & vbNewLine & _
    "7,8,9" Then
        TestJaggedArrayToCSVString = False
        Debug.Print "vbNewLine"
    End If
    
    '*****Quote option*****
    
    'Embedded
    S = JaggedArrayToCSVString(Arr, ",", NewLineCharacterCRLF, QuoteCharacterDouble, QuoteOptionQuoteEmbedded)
    If S <> """Header1,Test"",Header2,Header3" & vbCrLf & _
    "1,2,3" & vbCrLf & _
    "4,5,6" & vbCrLf & _
    "7,8,9" Then
        TestJaggedArrayToCSVString = False
        Debug.Print "Embedded"
    End If
    
    'All
    S = JaggedArrayToCSVString(Arr, ",", NewLineCharacterCRLF, QuoteCharacterDouble, QuoteOptionQuoteAll)
    If S <> """Header1,Test"",""Header2"",""Header3""" & vbCrLf & _
    """1"",""2"",""3""" & vbCrLf & _
    """4"",""5"",""6""" & vbCrLf & _
    """7"",""8"",""9""" Then
        TestJaggedArrayToCSVString = False
        Debug.Print "All"
    End If
    
    'None
    S = JaggedArrayToCSVString(Arr, ",", NewLineCharacterCRLF, QuoteCharacterDouble, QuoteOptionQuoteNone)
    If S <> "Header1,Test,Header2,Header3" & vbCrLf & _
    "1,2,3" & vbCrLf & _
    "4,5,6" & vbCrLf & _
    "7,8,9" Then
        TestJaggedArrayToCSVString = False
        Debug.Print "None"
    End If
    
    'NonNumeric
    S = JaggedArrayToCSVString(Arr, ",", NewLineCharacterCRLF, QuoteCharacterDouble, QuoteOptionQuoteNonNumeric)
    If S <> """Header1,Test"",""Header2"",""Header3""" & vbCrLf & _
    "1,2,3" & vbCrLf & _
    "4,5,6" & vbCrLf & _
    "7,8,9" Then
        TestJaggedArrayToCSVString = False
        Debug.Print "NonNumeric"
    End If
    
    '*****Quote Character*****
    
    'Double Quote All
    S = JaggedArrayToCSVString(Arr, ",", NewLineCharacterCRLF, QuoteCharacterDouble, QuoteOptionQuoteAll)
    If S <> """Header1,Test"",""Header2"",""Header3""" & vbCrLf & _
    """1"",""2"",""3""" & vbCrLf & _
    """4"",""5"",""6""" & vbCrLf & _
    """7"",""8"",""9""" Then
        TestJaggedArrayToCSVString = False
        Debug.Print "Double quote all"
    End If
    
    'Single Quote All
    S = JaggedArrayToCSVString(Arr, ",", NewLineCharacterCRLF, QuoteCharacterSingle, QuoteOptionQuoteAll)
    If S <> "'Header1,Test','Header2','Header3'" & vbCrLf & _
    "'1','2','3'" & vbCrLf & _
    "'4','5','6'" & vbCrLf & _
    "'7','8','9'" Then
        TestJaggedArrayToCSVString = False
        Debug.Print "Single quote all"
    End If
    
    'No Quote All
    S = JaggedArrayToCSVString(Arr, ",", NewLineCharacterCRLF, QuoteCharacterNone, QuoteOptionQuoteAll)
    If S <> "Header1,Test,Header2,Header3" & vbCrLf & _
    "1,2,3" & vbCrLf & _
    "4,5,6" & vbCrLf & _
    "7,8,9" Then
        TestJaggedArrayToCSVString = False
        Debug.Print "No quote all"
    End If
    
    Debug.Print "TestJaggedArrayToCSVString: " & TestJaggedArrayToCSVString

End Function

Private Function TestStringArrayToCSVString() As Boolean
    
    TestStringArrayToCSVString = True
    
    Dim S$
    Dim Arr$()
    
    ReDim Arr(0 To 3, 0 To 2)
    Arr(0, 0) = "Header1"
    Arr(0, 1) = "Header2"
    Arr(0, 2) = "Header3"
    Arr(1, 0) = "1"
    Arr(1, 1) = "2"
    Arr(1, 2) = "3"
    Arr(2, 0) = "4"
    Arr(2, 1) = "5"
    Arr(2, 2) = "6"
    Arr(3, 0) = "7"
    Arr(3, 1) = "8"
    Arr(3, 2) = "9"
    
    'Normal
    S = StringArrayToCSVString(Arr, ",", NewLineCharacterCRLF, QuoteCharacterDouble, QuoteOptionQuoteEmbedded)
    If S <> "Header1,Header2,Header3" & vbCrLf & _
    "1,2,3" & vbCrLf & _
    "4,5,6" & vbCrLf & _
    "7,8,9" Then
        TestStringArrayToCSVString = False
        Debug.Print "Normal"
    End If
    
    'Empty Array
    ReDim Arr$(0 To 0, 0 To 0)
    Arr(0, 0) = ""
    S = StringArrayToCSVString(Arr, ",", NewLineCharacterCRLF, QuoteCharacterDouble, QuoteOptionQuoteEmbedded)
    If S <> "" Then
        TestStringArrayToCSVString = False
        Debug.Print "Empty Array"
    End If
    
    ReDim Arr(0 To 3, 0 To 2)
    Arr(0, 0) = "Header1,Test"
    Arr(0, 1) = "Header2"
    Arr(0, 2) = "Header3"
    Arr(1, 0) = "1"
    Arr(1, 1) = "2"
    Arr(1, 2) = "3"
    Arr(2, 0) = "4"
    Arr(2, 1) = "5"
    Arr(2, 2) = "6"
    Arr(3, 0) = "7"
    Arr(3, 1) = "8"
    Arr(3, 2) = "9"
    
    '*****New Line Character*****
    
    'vbCrLf
    S = StringArrayToCSVString(Arr, ",", NewLineCharacterCRLF, QuoteCharacterDouble, QuoteOptionQuoteEmbedded)
    If S <> """Header1,Test"",Header2,Header3" & vbCrLf & _
    "1,2,3" & vbCrLf & _
    "4,5,6" & vbCrLf & _
    "7,8,9" Then
        TestStringArrayToCSVString = False
        Debug.Print "vbCrLf"
    End If
    
    'vbCr
    S = StringArrayToCSVString(Arr, ",", NewLineCharacterCR, QuoteCharacterDouble, QuoteOptionQuoteEmbedded)
    If S <> """Header1,Test"",Header2,Header3" & vbCr & _
    "1,2,3" & vbCr & _
    "4,5,6" & vbCr & _
    "7,8,9" Then
        TestStringArrayToCSVString = False
        Debug.Print "vbCr"
    End If
    
    'vbLf
    S = StringArrayToCSVString(Arr, ",", NewLineCharacterLF, QuoteCharacterDouble, QuoteOptionQuoteEmbedded)
    If S <> """Header1,Test"",Header2,Header3" & vbLf & _
    "1,2,3" & vbLf & _
    "4,5,6" & vbLf & _
    "7,8,9" Then
        TestStringArrayToCSVString = False
        Debug.Print "vbLf"
    End If
    
    'Default
    S = StringArrayToCSVString(Arr, ",", NewLineCharacterDefault, QuoteCharacterDouble, QuoteOptionQuoteEmbedded)
    If S <> """Header1,Test"",Header2,Header3" & vbNewLine & _
    "1,2,3" & vbNewLine & _
    "4,5,6" & vbNewLine & _
    "7,8,9" Then
        TestStringArrayToCSVString = False
        Debug.Print "vbNewLine"
    End If
    
    '*****Quote option*****
    
    'Embedded
    S = StringArrayToCSVString(Arr, ",", NewLineCharacterCRLF, QuoteCharacterDouble, QuoteOptionQuoteEmbedded)
    If S <> """Header1,Test"",Header2,Header3" & vbCrLf & _
    "1,2,3" & vbCrLf & _
    "4,5,6" & vbCrLf & _
    "7,8,9" Then
        TestStringArrayToCSVString = False
        Debug.Print "Embedded"
    End If
    
    'All
    S = StringArrayToCSVString(Arr, ",", NewLineCharacterCRLF, QuoteCharacterDouble, QuoteOptionQuoteAll)
    If S <> """Header1,Test"",""Header2"",""Header3""" & vbCrLf & _
    """1"",""2"",""3""" & vbCrLf & _
    """4"",""5"",""6""" & vbCrLf & _
    """7"",""8"",""9""" Then
        TestStringArrayToCSVString = False
        Debug.Print "All"
    End If
    
    'None
    S = StringArrayToCSVString(Arr, ",", NewLineCharacterCRLF, QuoteCharacterDouble, QuoteOptionQuoteNone)
    If S <> "Header1,Test,Header2,Header3" & vbCrLf & _
    "1,2,3" & vbCrLf & _
    "4,5,6" & vbCrLf & _
    "7,8,9" Then
        TestStringArrayToCSVString = False
        Debug.Print "None"
    End If
    
    'NonNumeric
    S = StringArrayToCSVString(Arr, ",", NewLineCharacterCRLF, QuoteCharacterDouble, QuoteOptionQuoteNonNumeric)
    If S <> """Header1,Test"",""Header2"",""Header3""" & vbCrLf & _
    "1,2,3" & vbCrLf & _
    "4,5,6" & vbCrLf & _
    "7,8,9" Then
        TestStringArrayToCSVString = False
        Debug.Print "Non numeric"
    End If
    
    '*****Quote Character*****
    
    'Double Quote All
    S = StringArrayToCSVString(Arr, ",", NewLineCharacterCRLF, QuoteCharacterDouble, QuoteOptionQuoteAll)
    If S <> """Header1,Test"",""Header2"",""Header3""" & vbCrLf & _
    """1"",""2"",""3""" & vbCrLf & _
    """4"",""5"",""6""" & vbCrLf & _
    """7"",""8"",""9""" Then
        TestStringArrayToCSVString = False
        Debug.Print "Double quote all"
    End If
    
    'Single Quote All
    S = StringArrayToCSVString(Arr, ",", NewLineCharacterCRLF, QuoteCharacterSingle, QuoteOptionQuoteAll)
    If S <> "'Header1,Test','Header2','Header3'" & vbCrLf & _
    "'1','2','3'" & vbCrLf & _
    "'4','5','6'" & vbCrLf & _
    "'7','8','9'" Then
        TestStringArrayToCSVString = False
        Debug.Print "Single quote all"
    End If
    
    'No Quote All
    S = StringArrayToCSVString(Arr, ",", NewLineCharacterCRLF, QuoteCharacterNone, QuoteOptionQuoteAll)
    If S <> "Header1,Test,Header2,Header3" & vbCrLf & _
    "1,2,3" & vbCrLf & _
    "4,5,6" & vbCrLf & _
    "7,8,9" Then
        TestStringArrayToCSVString = False
        Debug.Print "No quote all"
    End If
    
    Debug.Print "TestStringArrayToCSVString: " & TestStringArrayToCSVString
    
End Function

Private Function TestJaggedArrayToCSVFile() As Boolean

    TestJaggedArrayToCSVFile = True

    Dim S$

    Dim Arr()
    ReDim Arr(0 To 3)

    Arr(0) = Array("Header1", "Header2", "Header3")
    Arr(1) = Array(1, 2, 3)
    Arr(2) = Array(4, 5, 6)
    Arr(3) = Array(7, 8, 9)

    Dim FilePath$
    FilePath = _
    Environ$("USERPROFILE") & "\Desktop\example" & Format$(Now, "mmddyyyhhmmss") & ".csv"
    
    JaggedArrayToCSVFile Arr, FilePath, ",", NewLineCharacterCRLF, QuoteCharacterDouble, QuoteOptionQuoteEmbedded
    
    S = ReadAllFromTextFile(FilePath)
    
    If S <> "Header1,Header2,Header3" & vbCrLf & _
    "1,2,3" & vbCrLf & _
    "4,5,6" & vbCrLf & _
    "7,8,9" Then
        TestJaggedArrayToCSVFile = False
        Debug.Print "TestJaggedArrayToCSVFile"
    End If
    
    Kill FilePath
    
    Debug.Print "TestJaggedArrayToCSVFile: " & TestJaggedArrayToCSVFile
    
End Function

Private Function TestStringArrayToCSVFile() As Boolean
    
    TestStringArrayToCSVFile = True
    
    Dim S$
    
    Dim Arr$()
    ReDim Arr(0 To 3, 0 To 2)
    
    Arr(0, 0) = "Header1"
    Arr(0, 1) = "Header2"
    Arr(0, 2) = "Header3"
    Arr(1, 0) = "1"
    Arr(1, 1) = "2"
    Arr(1, 2) = "3"
    Arr(2, 0) = "4"
    Arr(2, 1) = "5"
    Arr(2, 2) = "6"
    Arr(3, 0) = "7"
    Arr(3, 1) = "8"
    Arr(3, 2) = "9"
    
    Dim FilePath$
    FilePath = _
    Environ$("USERPROFILE") & "\Desktop\example" & Format$(Now, "mmddyyyhhmmss") & ".csv"
    
    StringArrayToCSVFile Arr, FilePath, ",", NewLineCharacterCRLF, QuoteCharacterDouble, QuoteOptionQuoteEmbedded
    
    S = ReadAllFromTextFile(FilePath)
    
    If S <> "Header1,Header2,Header3" & vbCrLf & _
    "1,2,3" & vbCrLf & _
    "4,5,6" & vbCrLf & _
    "7,8,9" Then
        TestStringArrayToCSVFile = False
        Debug.Print "TestJaggedArrayToCSVFile"
    End If
    
    Kill FilePath
    
    Debug.Print "TestStringArrayToCSVFile: " & TestStringArrayToCSVFile
    
End Function
