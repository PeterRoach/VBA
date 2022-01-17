Attribute VB_Name = "testArrayListUDT"
Option Explicit

Public Type TExample
    Message As String
    Flag    As Boolean
End Type

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
'  Module Name: testArrayListUDT
'  Module Description: Unit tests for clsArrayListUDT class.
'  Module Version: 1.0
'  Module License: MIT
'  Module Author: Peter Roach; PeterRoach.Code@gmail.com
'  Module Contents:
'        TestArrayListUDT
'        TestCapacity
'        TestSize
'        TestItem
'        TestGrowCapacity
'        TestShrinkCapacity
'        TestEnsureCapacity
'        TestTrimToSize
'        TestReinitialize
'        TestAppend
'        TestAppendArray
'        TestAppendArrayList
'        TestInsert
'        TestInsertArray
'        TestInsertArrayList
'        TestRemove
'        TestRemoveRange
'        TestClear
'        TestReverse
'        TestToArray


Private Const DEFAULT_CAPACITY& = 10


'Example Usage=========================================================
'======================================================================

Public Sub Example()

    Dim AL As clsArrayListUDT
    Set AL = New clsArrayListUDT
    
    Dim E As TExample
    E.Message = "A"
    E.Flag = True
    
    AL.Append E
    
    With AL.Item(0)
        Debug.Print .Message, .Flag
    End With
    
    Debug.Print AL.Size
    
    AL.Remove 0
    
    Debug.Print AL.Size
    
    AL.Reinitialize
    
    Dim Arr() As TExample
    ReDim Arr(0 To 2)
    With Arr(0)
        .Message = "A"
        .Flag = True
    End With
    With Arr(1)
        .Message = "B"
        .Flag = True
    End With
    With Arr(2)
        .Message = "C"
        .Flag = True
    End With
    
    AL.AppendArray Arr
    
    Dim i As Long
    For i = 0 To AL.Size - 1
        Debug.Print AL.Item(i).Message
    Next i
    
End Sub


'Unit Tests============================================================
'======================================================================

Public Function TestArrayListUDT() As Boolean
    
    TestArrayListUDT = _
        TestCapacity And _
        TestSize And _
        TestItem And _
        TestGrowCapacity And _
        TestShrinkCapacity And _
        TestEnsureCapacity And _
        TestTrimToSize And _
        TestReinitialize And _
        TestAppend And _
        TestAppendArray And _
        TestAppendArrayList And _
        TestInsert And _
        TestInsertArray And _
        TestInsertArrayList And _
        TestRemove And _
        TestRemoveRange And _
        TestClear And _
        TestReverse And _
        TestToArray

    Debug.Print "TestArrayListUDT: " & TestArrayListUDT
    
End Function

Private Function TestCapacity() As Boolean

    TestCapacity = True
    
    Dim AL As clsArrayListUDT
    Set AL = New clsArrayListUDT
    
    Dim E As TExample
    
    ''Causes compiler error - Test Passed
    'AL.Capacity = DEFAULT_CAPACITY
    
    'GrowCapacity
    AL.GrowCapacity AL.Capacity + 1
    If AL.Capacity <> DEFAULT_CAPACITY + 1 Then
        TestCapacity = False
        Debug.Print "GrowCapacity"
    End If
    
    'ShrinkCapacity
    AL.ShrinkCapacity AL.Capacity - 1
    If AL.Capacity <> DEFAULT_CAPACITY Then
        TestCapacity = False
        Debug.Print "ShrinkCapacity"
    End If
    
    'EnsureCapacity
    AL.EnsureCapacity AL.Capacity + 1
    If AL.Capacity <> DEFAULT_CAPACITY * 2 + 2 Then
        TestCapacity = False
        Debug.Print "EnsureCapacity"
    End If
    
    'TrimToSize
    AL.Append E
    AL.TrimToSize
    If AL.Capacity <> 1 Then
        TestCapacity = False
        Debug.Print "TrimToSize"
    End If
    
    'Reinitialize
    AL.Reinitialize
    If AL.Capacity <> DEFAULT_CAPACITY Then
        TestCapacity = False
        Debug.Print "Reinitialize"
    End If
    
    'Append
    Dim i&
    For i = 1 To 11
        AL.Append E
    Next i
    If AL.Capacity <> DEFAULT_CAPACITY * 2 + 2 Then
        TestCapacity = False
        Debug.Print "Append"
    End If
    
    'AppendArray
    AL.Reinitialize
    Dim Arr(0 To 10) As TExample
    AL.AppendArray Arr
    If AL.Capacity <> DEFAULT_CAPACITY * 2 + 2 Then
        TestCapacity = False
        Debug.Print "AppendArray"
    End If
    
    'AppendArrayList
    AL.Reinitialize
    Dim AL1 As clsArrayListUDT
    Set AL1 = New clsArrayListUDT
    AL1.AppendArray Arr
    AL.AppendArrayList AL1
    If AL.Capacity <> DEFAULT_CAPACITY * 2 + 2 Then
        TestCapacity = False
        Debug.Print "AppendArrayList"
    End If
    
    'Insert
    AL.Reinitialize
    For i = 1 To 11
        AL.Insert 0, E
    Next i
    If AL.Capacity <> DEFAULT_CAPACITY * 2 + 2 Then
        TestCapacity = False
        Debug.Print "Insert"
    End If
    
    'InsertArray
    AL.Reinitialize
    AL.InsertArray 0, Arr
    If AL.Capacity <> DEFAULT_CAPACITY * 2 + 2 Then
        TestCapacity = False
        Debug.Print "InsertArray"
    End If
    
    'InsertArrayList
    AL.Reinitialize
    AL.InsertArrayList 0, AL1
    If AL.Capacity <> DEFAULT_CAPACITY * 2 + 2 Then
        TestCapacity = False
        Debug.Print "InsertArrayList"
    End If
    
    'Remove
    AL.Reinitialize
    AL.Append E
    AL.Remove 0
    If AL.Capacity <> DEFAULT_CAPACITY Then
        TestCapacity = False
        Debug.Print "Remove"
    End If
    
    'RemoveRange
    AL.Reinitialize
    AL.Append E
    AL.Append E
    AL.Append E
    AL.RemoveRange 0, 1
    If AL.Capacity <> DEFAULT_CAPACITY Then
        TestCapacity = False
        Debug.Print "RemoveRange"
    End If
    
    'Clear
    AL.Reinitialize
    AL.AppendArray Arr
    AL.Clear
    If AL.Capacity <> DEFAULT_CAPACITY * 2 + 2 Then
        TestCapacity = False
        Debug.Print "Clear"
    End If
    
    'Reverse
    AL.Reinitialize
    AL.Append E
    AL.Append E
    AL.Append E
    AL.Reverse
    If AL.Capacity <> DEFAULT_CAPACITY Then
        TestCapacity = False
        Debug.Print "Reverse"
    End If
    
    'ToArray
    AL.Reinitialize
    AL.Append E
    AL.Append E
    AL.Append E
    If AL.Capacity <> DEFAULT_CAPACITY Then
        TestCapacity = False
        Debug.Print "ToStringArray"
    End If
    
    'JoinString
    AL.Reinitialize
    AL.Append E
    AL.Append E
    AL.Append E
    If AL.Capacity <> DEFAULT_CAPACITY Then
        TestCapacity = False
        Debug.Print "JoinString"
    End If

    Debug.Print "TestCapacity: " & TestCapacity

End Function

Private Function TestSize() As Boolean

    TestSize = True
    
    Dim AL As clsArrayListUDT
    Set AL = New clsArrayListUDT
    
    Dim E As TExample
    
    'Causes compiler error - Test Passed
    'AL.Size = 1
    
    'GrowCapacity
    AL.GrowCapacity AL.Capacity + 1
    If AL.Size <> 0 Then
        TestSize = False
        Debug.Print "GrowCapacity"
    End If
    
    'ShrinkCapacity
    AL.ShrinkCapacity AL.Capacity - 1
    If AL.Size <> 0 Then
        TestSize = False
        Debug.Print "ShrinkCapacity"
    End If
    
    'EnsureCapacity
    AL.EnsureCapacity AL.Capacity + 1
    If AL.Size <> 0 Then
        TestSize = False
        Debug.Print "EnsureCapacity"
    End If
    
    'TrimToSize
    AL.TrimToSize
    If AL.Size <> 0 Then
        TestSize = False
        Debug.Print "TrimToSize"
    End If
    AL.Append E
    AL.TrimToSize
    If AL.Size <> 1 Then
        TestSize = False
        Debug.Print "TrimToSize"
    End If
    
    'Reinitialize
    AL.Reinitialize
    If AL.Size <> 0 Then
        TestSize = False
        Debug.Print "Reinitialize"
    End If
    
    'Append
    AL.Append E
    If AL.Size <> 1 Then
        TestSize = False
        Debug.Print "Append"
    End If
    
    'AppendArray
    Dim Arr(0 To 2) As TExample
    AL.AppendArray Arr
    If AL.Size <> 4 Then
        TestSize = False
        Debug.Print "AppendArray"
    End If
    
    'AppendArrayList
    Dim AL1 As clsArrayListUDT
    Set AL1 = New clsArrayListUDT
    AL1.AppendArray Arr
    AL.AppendArrayList AL1
    If AL.Size <> 7 Then
        TestSize = False
        Debug.Print "AppendArrayList"
    End If
    
    'Insert
    AL.Insert 0, E
    If AL.Size <> 8 Then
        TestSize = False
        Debug.Print "Insert"
    End If
    
    'InsertArray
    AL.InsertArray 0, Arr
    If AL.Size <> 11 Then
        TestSize = False
        Debug.Print "InsertArray"
    End If
    
    'InsertArrayList
    AL.AppendArrayList AL1
    If AL.Size <> 14 Then
        TestSize = False
        Debug.Print "InsertArrayList"
    End If
    
    'Remove
    AL.Remove 0
    If AL.Size <> 13 Then
        TestSize = False
        Debug.Print "Remove"
    End If
    
    'RemoveRange
    AL.RemoveRange 0, 1
    If AL.Size <> 11 Then
        TestSize = False
        Debug.Print "RemoveRange"
    End If
    
    'Clear
    AL.Clear
    If AL.Size <> 0 Then
        TestSize = False
        Debug.Print "Clear"
    End If
    
    'Reverse
    AL.Reinitialize
    AL.Append E
    AL.Append E
    AL.Append E
    AL.Reverse
    If AL.Size <> 3 Then
        TestSize = False
        Debug.Print "Reverse"
    End If
    
    'ToArray
    AL.Reinitialize
    AL.Append E
    AL.Append E
    AL.Append E
    If AL.Size <> 3 Then
        TestSize = False
        Debug.Print "ToStringArray"
    End If
    
    Debug.Print "TestSize: " & TestSize

End Function

Private Function TestItem() As Boolean

    TestItem = True

    Dim AL As clsArrayListUDT
    Set AL = New clsArrayListUDT
    
    Dim E As TExample
    E.Message = "A"
    
    'Empty
        'Get
        On Error Resume Next
        Debug.Print AL.Item(0).Message
        If Err.NUMBER <> 9 Then
            TestItem = False
            Debug.Print "Empty Get"
        End If
        On Error GoTo 0
        'Let
        On Error Resume Next
        AL.Item(0) = E
        If Err.NUMBER <> 9 Then
            TestItem = False
            Debug.Print "Empty Let"
        End If
        On Error GoTo 0

    Dim E1 As TExample
    E1.Message = "B"
    
    'Non-Empty
    AL.Append E
        'Get
        If AL.Item(0).Message <> "A" Then
            TestItem = False
            Debug.Print "Non-Empty Get"
        End If
        'Let
        AL.Item(0) = E1
        If AL.Item(0).Message <> "B" Then
            TestItem = False
            Debug.Print "Non-Empty Let"
        End If

    'Invalid lower bound
    On Error Resume Next
    Debug.Print AL.Item(-1).Message
    If Err.NUMBER <> 9 Then
        TestItem = False
        Debug.Print "Invalid lower bound"
    End If
    On Error GoTo 0

    'Invalid upper bound
    On Error Resume Next
    Debug.Print AL.Item(1).Message
    If Err.NUMBER <> 9 Then
        TestItem = False
        Debug.Print "Invalid upper bound"
    End If
    On Error GoTo 0

    Debug.Print "TestItem: " & TestItem

End Function

Private Function TestGrowCapacity() As Boolean

    TestGrowCapacity = True
    
    Dim AL As clsArrayListUDT
    Set AL = New clsArrayListUDT
    
    'Less
    AL.GrowCapacity AL.Capacity - 1
    If AL.Capacity <> DEFAULT_CAPACITY Then
        TestGrowCapacity = False
        Debug.Print "Less"
    End If
    
    'Same
    AL.GrowCapacity AL.Capacity
    If AL.Capacity <> DEFAULT_CAPACITY Then
        TestGrowCapacity = False
        Debug.Print "Same"
    End If
    
    'More
    AL.GrowCapacity AL.Capacity + 1
    If AL.Capacity <> DEFAULT_CAPACITY + 1 Then
        TestGrowCapacity = False
        Debug.Print "More"
    End If
    
    'Invalid 0 - implicitly not possible
    AL.Reinitialize
    AL.GrowCapacity 0
    If AL.Capacity <> DEFAULT_CAPACITY Then
        TestGrowCapacity = False
        Debug.Print "Invalid 0"
    End If
    
    'Invalid negative - implicitly not possible
    AL.Reinitialize
    AL.GrowCapacity -1
    If AL.Capacity <> DEFAULT_CAPACITY Then
        TestGrowCapacity = False
        Debug.Print "Invalid negative"
    End If
    
    Debug.Print "TestGrowCapacity: " & TestGrowCapacity

End Function

Private Function TestShrinkCapacity() As Boolean

    TestShrinkCapacity = True
    
    Dim AL As clsArrayListUDT
    Set AL = New clsArrayListUDT
    
    'More
    AL.ShrinkCapacity AL.Capacity + 1
    If AL.Capacity <> DEFAULT_CAPACITY Then
        TestShrinkCapacity = False
        Debug.Print "More"
    End If
    
    'Same
    AL.ShrinkCapacity AL.Capacity
    If AL.Capacity <> DEFAULT_CAPACITY Then
        TestShrinkCapacity = False
        Debug.Print "Same"
    End If
    
    'Less
    AL.ShrinkCapacity AL.Capacity - 1
    If AL.Capacity <> DEFAULT_CAPACITY - 1 Then
        TestShrinkCapacity = False
        Debug.Print "Less"
    End If
    
    Dim E As TExample
    
    'Size
    AL.Reinitialize
    AL.Append E
    AL.Append E
    AL.Append E
    AL.ShrinkCapacity 3
    If AL.Size <> 3 Then
        TestShrinkCapacity = False
        Debug.Print "Size"
    End If
    
    'Less than size
    AL.Reinitialize
    AL.Append E
    AL.Append E
    AL.Append E
    AL.ShrinkCapacity 2
    If AL.Size <> 3 Then
        TestShrinkCapacity = False
        Debug.Print "Less than size"
    End If
    
    'Invalid 0
    AL.Reinitialize
    On Error Resume Next
    AL.ShrinkCapacity 0
    If Err.NUMBER <> 5 Then
        TestShrinkCapacity = False
        Debug.Print "Invalid 0"
    End If
    On Error GoTo 0
    
    'Invalid negative
    AL.Reinitialize
    On Error Resume Next
    AL.ShrinkCapacity -1
    If Err.NUMBER <> 5 Then
        TestShrinkCapacity = False
        Debug.Print "Invalid negative"
    End If
    On Error GoTo 0
    
    Debug.Print "TestShrinkCapacity: " & TestShrinkCapacity

End Function

Private Function TestEnsureCapacity() As Boolean

    TestEnsureCapacity = True
    
    Dim AL As clsArrayListUDT
    Set AL = New clsArrayListUDT
    
    'Less
    AL.EnsureCapacity AL.Capacity - 1
    If AL.Capacity <> DEFAULT_CAPACITY Then
        TestEnsureCapacity = False
        Debug.Print "Less"
    End If
    
    'Same
    AL.EnsureCapacity AL.Capacity
    If AL.Capacity <> DEFAULT_CAPACITY Then
        TestEnsureCapacity = False
        Debug.Print "Same"
    End If
    
    'More
    AL.EnsureCapacity AL.Capacity + 1
    If AL.Capacity <> DEFAULT_CAPACITY * 2 + 2 Then
        TestEnsureCapacity = False
        Debug.Print "More"
    End If
    
    'More than double
    AL.Reinitialize
    AL.EnsureCapacity DEFAULT_CAPACITY * 2 + 3
    If AL.Capacity <> DEFAULT_CAPACITY * 2 + 3 Then
        TestEnsureCapacity = False
        Debug.Print "More than double"
    End If
    
    'Invalid 0 - implicitly not possible
     AL.Reinitialize
     AL.EnsureCapacity 0
    If AL.Capacity <> DEFAULT_CAPACITY Then
        TestEnsureCapacity = False
        Debug.Print "Invalid 0"
    End If
    
    'Invalid negative - implicitly not possible
     AL.Reinitialize
     AL.EnsureCapacity -1
    If AL.Capacity <> DEFAULT_CAPACITY Then
        TestEnsureCapacity = False
        Debug.Print "Invalid negative"
    End If
    
    Debug.Print "TestEnsureCapacity: " & TestEnsureCapacity

End Function

Private Function TestTrimToSize() As Boolean

    TestTrimToSize = True

    Dim AL As clsArrayListUDT
    Set AL = New clsArrayListUDT

    Dim E As TExample
    
    'Size 0
    AL.TrimToSize
    If AL.Capacity <> DEFAULT_CAPACITY Or AL.Size <> 0 Then
        TestTrimToSize = False
        Debug.Print "Size 0"
    End If

    'Size 1
    AL.Append E
    AL.TrimToSize
    If AL.Capacity <> 1 Or AL.Size <> 1 Then
        TestTrimToSize = False
        Debug.Print "Size 1"
    End If

    'Size > 1
    AL.Reinitialize
    AL.Append E
    AL.Append E
    AL.Append E
    AL.TrimToSize
    If AL.Capacity <> 3 Or AL.Size <> 3 Then
        TestTrimToSize = False
        Debug.Print "Size > 1"
    End If

    'Reinitialize
    AL.Reinitialize
    AL.GrowCapacity DEFAULT_CAPACITY * 2 + 2
    AL.Append E
    AL.Append E
    AL.Append E
    AL.Clear
    AL.TrimToSize
    If AL.Capacity <> DEFAULT_CAPACITY Or AL.Size <> 0 Then
        TestTrimToSize = False
        Debug.Print "Reinitialize"
    End If

    Debug.Print "TestTrimToSize: " & TestTrimToSize

End Function

Private Function TestReinitialize() As Boolean

    TestReinitialize = True
    
    Dim AL As clsArrayListUDT
    Set AL = New clsArrayListUDT
    
    Dim E As TExample
    
    'Default
    AL.GrowCapacity DEFAULT_CAPACITY * 2 + 2
    AL.Append E
    AL.Append E
    AL.Append E
    AL.Reinitialize
    If AL.Capacity <> DEFAULT_CAPACITY Or AL.Size <> 0 Then
        TestReinitialize = False
        Debug.Print "Default"
    End If
    
    'Less than default
    AL.Reinitialize
    AL.GrowCapacity DEFAULT_CAPACITY * 2 + 2
    AL.Append E
    AL.Append E
    AL.Append E
    AL.Reinitialize DEFAULT_CAPACITY - 1
    If AL.Capacity <> DEFAULT_CAPACITY - 1 Or AL.Size <> 0 Then
        TestReinitialize = False
        Debug.Print "Less than default"
    End If
    
    'More than default
    AL.Reinitialize
    AL.GrowCapacity DEFAULT_CAPACITY * 2 + 2
    AL.Append E
    AL.Append E
    AL.Append E
    AL.Reinitialize DEFAULT_CAPACITY + 1
    If AL.Capacity <> DEFAULT_CAPACITY + 1 Or AL.Size <> 0 Then
        TestReinitialize = False
        Debug.Print "More than default"
    End If
    
    'Invalid 0
    AL.Reinitialize
    AL.GrowCapacity DEFAULT_CAPACITY * 2 + 2
    AL.Append E
    AL.Append E
    AL.Append E
    On Error Resume Next
    AL.Reinitialize 0
    If Err.NUMBER <> 5 Then
        TestReinitialize = False
        Debug.Print "Invalid 0"
    End If
    On Error GoTo 0
    
    'Invalid negative
    AL.Reinitialize
    AL.GrowCapacity DEFAULT_CAPACITY * 2 + 2
    AL.Append E
    AL.Append E
    AL.Append E
    On Error Resume Next
    AL.Reinitialize 0
    If Err.NUMBER <> 5 Then
        TestReinitialize = False
        Debug.Print "Invalid negative"
    End If
    On Error GoTo 0
    
    Debug.Print "TestReinitialize: " & TestReinitialize

End Function

Private Function TestAppend() As Boolean

    TestAppend = True

    Dim AL As clsArrayListUDT
    Set AL = New clsArrayListUDT
    
    Dim E As TExample
    E.Message = "A"
    
    'Empty
    AL.Append E
    If AL.Item(0).Message <> "A" Or AL.Size <> 1 Then
        TestAppend = False
        Debug.Print "Empty"
    End If

    E.Message = "B"
    
    'Non empty
    AL.Append E
    If AL.Item(0).Message & AL.Item(1).Message <> "AB" Or AL.Size <> 2 Then
        TestAppend = False
        Debug.Print "Non Empty"
    End If

    'Until capacity
    Dim i&
    AL.Reinitialize
    For i = 1 To 11
        AL.Append E
    Next i
    If AL.Item(0).Message & _
    AL.Item(1).Message & _
    AL.Item(2).Message & _
    AL.Item(3).Message & _
    AL.Item(4).Message & _
    AL.Item(5).Message & _
    AL.Item(6).Message & _
    AL.Item(7).Message & _
    AL.Item(8).Message & _
    AL.Item(9).Message & _
    AL.Item(10).Message <> String$(11, "B") Or _
    AL.Capacity <> DEFAULT_CAPACITY * 2 + 2 Or _
    AL.Size <> 11 Then
        TestAppend = False
        Debug.Print "Until capacity"
    End If

    Debug.Print "TestAppend: " & TestAppend

End Function

Private Function TestAppendArray() As Boolean

    TestAppendArray = True

    Dim AL As clsArrayListUDT
    Set AL = New clsArrayListUDT
    
    Dim E As TExample
    E.Message = "A"
    
    Dim Arr(0 To 2) As TExample

    'Empty
    Arr(0) = E
    Arr(1) = E
    Arr(2) = E
    AL.AppendArray Arr
    If AL.Item(0).Message & AL.Item(1).Message & AL.Item(2).Message _
    <> "AAA" Or AL.Size <> 3 Then
        TestAppendArray = False
        Debug.Print "Empty"
    End If

    'Non empty
    Arr(0) = E
    Arr(1) = E
    Arr(2) = E
    AL.AppendArray Arr
    If AL.Item(0).Message & AL.Item(1).Message & AL.Item(2).Message & _
    AL.Item(3).Message & AL.Item(4).Message & AL.Item(5).Message _
    <> "AAAAAA" Or AL.Size <> 6 Then
        TestAppendArray = False
        Debug.Print "Non empty"
    End If

    'Until capacity
    Arr(0) = E
    Arr(1) = E
    Arr(2) = E
    AL.AppendArray Arr
    AL.AppendArray Arr
    If AL.Item(0).Message & _
    AL.Item(1).Message & _
    AL.Item(2).Message & _
    AL.Item(3).Message & _
    AL.Item(4).Message & _
    AL.Item(5).Message & _
    AL.Item(6).Message & _
    AL.Item(7).Message & _
    AL.Item(8).Message & _
    AL.Item(9).Message & _
    AL.Item(10).Message & _
    AL.Item(11).Message _
    <> String$(12, "A") Or _
    AL.Size <> 12 Or _
    AL.Capacity <> DEFAULT_CAPACITY * 2 + 2 Then
        TestAppendArray = False
        Debug.Print "Until capacity"
    End If

    'Empty Array
    Dim Arr1() As TExample
    On Error Resume Next
    AL.AppendArray Arr1
    If Err.NUMBER <> 9 Then
        TestAppendArray = False
        Debug.Print "Empty Array"
    End If
    On Error GoTo 0

    Debug.Print "TestAppendArray: " & TestAppendArray

End Function

Private Function TestAppendArrayList() As Boolean

    TestAppendArrayList = True
    
    Dim AL As clsArrayListUDT
    Set AL = New clsArrayListUDT
    
    Dim AL1 As clsArrayListUDT
    Set AL1 = New clsArrayListUDT
    
    Dim E As TExample
    E.Message = "A"
    
    'Empty
    AL1.Append E
    AL1.Append E
    AL1.Append E
    AL.AppendArrayList AL1
    If AL.Item(0).Message & AL.Item(1).Message & AL.Item(2).Message _
    <> "AAA" Or AL.Size <> 3 Then
        TestAppendArrayList = False
        Debug.Print "Empty"
    End If
    
    'Non empty
    AL1.Reinitialize
    AL1.Append E
    AL1.Append E
    AL1.Append E
    AL.AppendArrayList AL1
    If AL.Item(0).Message & AL.Item(1).Message & AL.Item(2).Message & _
    AL.Item(3).Message & AL.Item(4).Message & AL.Item(5).Message _
    <> "AAAAAA" Or AL.Size <> 6 Then
        TestAppendArrayList = False
        Debug.Print "Non empty"
    End If
    
    E.Message = "A"
    
    'Until capacity
    AL1.Reinitialize
    AL1.Append E
    AL1.Append E
    AL1.Append E
    AL.AppendArrayList AL1
    AL1.Reinitialize
    AL1.Append E
    AL1.Append E
    AL1.Append E
    AL.AppendArrayList AL1
    If AL.Item(0).Message & _
    AL.Item(1).Message & _
    AL.Item(2).Message & _
    AL.Item(3).Message & _
    AL.Item(4).Message & _
    AL.Item(5).Message & _
    AL.Item(6).Message & _
    AL.Item(7).Message & _
    AL.Item(8).Message & _
    AL.Item(9).Message & _
    AL.Item(10).Message & _
    AL.Item(11).Message _
    <> String$(12, "A") Or _
    AL.Size <> 12 Or _
    AL.Capacity <> DEFAULT_CAPACITY * 2 + 2 Then
        TestAppendArrayList = False
        Debug.Print "Until capacity"
    End If
    
    'Empty ArrayList
    AL.Reinitialize
    AL1.Reinitialize
    AL.AppendArrayList AL1
    If AL.Capacity <> DEFAULT_CAPACITY Or _
    AL.Size <> 0 Then
        TestAppendArrayList = False
        Debug.Print "Empty ArrayList"
    End If
    
    Debug.Print "TestAppendArrayList: " & TestAppendArrayList

End Function

Private Function TestInsert() As Boolean

    TestInsert = True
    
    Dim AL As clsArrayListUDT
    Set AL = New clsArrayListUDT
    
    Dim E As TExample
    E.Message = "A"
    
    'Empty
    AL.Insert 0, E
    If AL.Item(0).Message <> "A" Or AL.Size <> 1 Then
        TestInsert = False
        Debug.Print "Empty"
    End If
    
    'Non empty
        'Start
        E.Message = "B"
        AL.Insert 0, E
        If AL.Item(0).Message <> "B" Or AL.Size <> 2 Then
            TestInsert = False
            Debug.Print "Start"
        End If
        'Middle
        E.Message = "C"
        AL.Insert 1, E
        If AL.Item(1).Message <> "C" Or AL.Size <> 3 Then
            TestInsert = False
            Debug.Print "Middle"
        End If
        'End
        E.Message = "D"
        AL.Insert 3, E
        If AL.Item(3).Message <> "D" Or AL.Size <> 4 Then
            TestInsert = False
            Debug.Print "End"
        End If
        
    'Invalid lower bound
    On Error Resume Next
    AL.Insert -1, E
    If Err.NUMBER <> 9 Then
        TestInsert = False
        Debug.Print "Invalid lower bound"
    End If
    On Error GoTo 0
    
    'Invalid upper bound
    On Error Resume Next
    AL.Insert 5, E
    If Err.NUMBER <> 9 Then
        TestInsert = False
        Debug.Print "Invalid upper bound"
    End If
    On Error GoTo 0
    
    Debug.Print "TestInsert: " & TestInsert

End Function

Private Function TestInsertArray() As Boolean

    TestInsertArray = True

    Dim AL As clsArrayListUDT
    Set AL = New clsArrayListUDT

    Dim E As TExample
    
    Dim Arr(0 To 2) As TExample
    
    'Empty
    E.Message = "A"
    Arr(0) = E
    E.Message = "B"
    Arr(1) = E
    E.Message = "C"
    Arr(2) = E
    AL.InsertArray 0, Arr
    If AL.Item(0).Message & AL.Item(1).Message & AL.Item(2).Message <> "ABC" _
    Or AL.Size <> 3 Or AL.Capacity <> DEFAULT_CAPACITY Then
        TestInsertArray = False
        Debug.Print "Empty"
    End If

    'Non empty
        'Start
        E.Message = "D"
        Arr(0) = E
        E.Message = "E"
        Arr(1) = E
        E.Message = "F"
        Arr(2) = E
        AL.InsertArray 0, Arr
        If AL.Item(0).Message & AL.Item(1).Message & AL.Item(2).Message & _
        AL.Item(3).Message & AL.Item(4).Message & AL.Item(5).Message _
        <> "DEFABC" Or AL.Size <> 6 Or AL.Capacity <> DEFAULT_CAPACITY Then
            TestInsertArray = False
            Debug.Print "Start"
        End If
        'Middle
        E.Message = "G"
        Arr(0) = E
        E.Message = "H"
        Arr(1) = E
        E.Message = "I"
        Arr(2) = E
        AL.InsertArray 3, Arr
        If AL.Item(0).Message & AL.Item(1).Message & AL.Item(2).Message & _
        AL.Item(3).Message & AL.Item(4).Message & AL.Item(5).Message & _
        AL.Item(6).Message & AL.Item(7).Message & AL.Item(8).Message _
        <> "DEFGHIABC" Or AL.Size <> 9 Or AL.Capacity <> DEFAULT_CAPACITY Then
            TestInsertArray = False
            Debug.Print "Middle"
        End If
        'End
        E.Message = "J"
        Arr(0) = E
        E.Message = "K"
        Arr(1) = E
        E.Message = "L"
        Arr(2) = E
        AL.InsertArray 9, Arr
        If AL.Item(0).Message & AL.Item(1).Message & AL.Item(2).Message & _
        AL.Item(3).Message & AL.Item(4).Message & AL.Item(5).Message & _
        AL.Item(6).Message & AL.Item(7).Message & AL.Item(8).Message & _
        AL.Item(9).Message & AL.Item(10).Message & AL.Item(11).Message _
        <> "DEFGHIABCJKL" Or AL.Size <> 12 Or AL.Capacity <> DEFAULT_CAPACITY * 2 + 2 Then
            TestInsertArray = False
            Debug.Print "End"
        End If

    'Invalid lower bound
    On Error Resume Next
    AL.InsertArray -1, Arr
    If Err.NUMBER <> 9 Then
        TestInsertArray = False
        Debug.Print "Invalid lower bound"
    End If
    On Error GoTo 0

    'Invalid upper bound
    On Error Resume Next
    AL.InsertArray 13, Arr
    If Err.NUMBER <> 9 Then
        TestInsertArray = False
        Debug.Print "Invalid upper bound"
    End If
    On Error GoTo 0

    'Empty array
    AL.Reinitialize
    Dim Arr1() As TExample
    On Error Resume Next
    AL.InsertArray 0, Arr1
    If Err.NUMBER <> 9 Then
        TestInsertArray = False
        Debug.Print "Empty array"
    End If
    On Error GoTo 0

    Debug.Print "TestInsertArray: " & TestInsertArray

End Function

Private Function TestInsertArrayList() As Boolean

    TestInsertArrayList = True

    Dim AL As clsArrayListUDT
    Set AL = New clsArrayListUDT
    
    Dim E As TExample
    
    Dim AL1 As clsArrayListUDT
    Set AL1 = New clsArrayListUDT
    
    'Empty
    AL1.Reinitialize
    E.Message = "A"
    AL1.Append E
    E.Message = "B"
    AL1.Append E
    E.Message = "C"
    AL1.Append E
    AL.InsertArrayList 0, AL1
    If AL.Item(0).Message & AL.Item(1).Message & AL.Item(2).Message <> "ABC" Or _
    AL.Size <> 3 Or AL.Capacity <> DEFAULT_CAPACITY Then
        TestInsertArrayList = False
        Debug.Print "Empty"
    End If

    'Non empty
        'Start
        AL1.Reinitialize
        E.Message = "D"
        AL1.Append E
        E.Message = "E"
        AL1.Append E
        E.Message = "F"
        AL1.Append E
        AL.InsertArrayList 0, AL1
        If AL.Item(0).Message & AL.Item(1).Message & AL.Item(2).Message & _
        AL.Item(3).Message & AL.Item(4).Message & AL.Item(5).Message _
        <> "DEFABC" Or AL.Size <> 6 Or AL.Capacity <> DEFAULT_CAPACITY Then
            TestInsertArrayList = False
            Debug.Print "Start"
        End If
        'Middle
        AL1.Reinitialize
        E.Message = "G"
        AL1.Append E
        E.Message = "H"
        AL1.Append E
        E.Message = "I"
        AL1.Append E
        AL.InsertArrayList 3, AL1
        If AL.Item(0).Message & AL.Item(1).Message & AL.Item(2).Message & _
        AL.Item(3).Message & AL.Item(4).Message & AL.Item(5).Message & _
        AL.Item(6).Message & AL.Item(7).Message & AL.Item(8).Message _
        <> "DEFGHIABC" Or AL.Size <> 9 Or AL.Capacity <> DEFAULT_CAPACITY Then
            TestInsertArrayList = False
            Debug.Print "Middle"
        End If
        'End
        AL1.Reinitialize
        E.Message = "J"
        AL1.Append E
        E.Message = "K"
        AL1.Append E
        E.Message = "L"
        AL1.Append E
        AL.InsertArrayList 9, AL1
        If AL.Item(0).Message & AL.Item(1).Message & AL.Item(2).Message & _
        AL.Item(3).Message & AL.Item(4).Message & AL.Item(5).Message & _
        AL.Item(6).Message & AL.Item(7).Message & AL.Item(8).Message & _
        AL.Item(9).Message & AL.Item(10).Message & AL.Item(11).Message _
        <> "DEFGHIABCJKL" Or AL.Size <> 12 Or AL.Capacity <> DEFAULT_CAPACITY * 2 + 2 Then
            TestInsertArrayList = False
            Debug.Print "End"
        End If

    'Invalid lower bound
    On Error Resume Next
    AL.InsertArrayList -1, AL1
    If Err.NUMBER <> 9 Then
        TestInsertArrayList = False
        Debug.Print "Invalid lower bound"
    End If
    On Error GoTo 0

    'Invalid upper bound
    On Error Resume Next
    AL.InsertArrayList 13, AL1
    If Err.NUMBER <> 9 Then
        TestInsertArrayList = False
        Debug.Print "Invalid upper bound"
    End If
    On Error GoTo 0

    'Empty ArrayList
    AL.Reinitialize
    AL1.Reinitialize
    AL.InsertArrayList 0, AL1
    If AL.Size <> 0 Or AL.Capacity <> DEFAULT_CAPACITY Then
        TestInsertArrayList = False
        Debug.Print "Empty ArrayList"
    End If

    Debug.Print "TestInsertArrayList: " & TestInsertArrayList

End Function

Private Function TestRemove() As Boolean

    TestRemove = True

    Dim AL As clsArrayListUDT
    Set AL = New clsArrayListUDT
    
    Dim E As TExample
    E.Message = "A"
    
    'Empty
    On Error Resume Next
    AL.Remove 0
    If Err.NUMBER <> 9 Then
        TestRemove = False
        Debug.Print "Empty"
    End If
    On Error GoTo 0

    'Non empty
    AL.Append E
    AL.Remove 0
    If AL.Size <> 0 Then
        TestRemove = False
        Debug.Print "Non empty"
    End If

    'Start
    AL.Reinitialize
    AL.Append E
    AL.Append E
    AL.Append E
    AL.Remove 0
    If AL.Size <> 2 Then
        TestRemove = False
        Debug.Print "Start"
    End If

    'Middle
    AL.Reinitialize
    AL.Append E
    AL.Append E
    AL.Append E
    AL.Remove 1
    If AL.Size <> 2 Then
        TestRemove = False
        Debug.Print "Middle"
    End If

    'End
    AL.Reinitialize
    AL.Append E
    AL.Append E
    AL.Append E
    AL.Remove AL.Size - 1
    If AL.Size <> 2 Then
        TestRemove = False
        Debug.Print "End"
    End If

    'Invalid lower bound
    AL.Reinitialize
    AL.Append E
    AL.Append E
    AL.Append E
    On Error Resume Next
    AL.Remove -1
    If Err.NUMBER <> 9 Then
        TestRemove = False
        Debug.Print "Invalid lower bound"
    End If

    'Invalid upper bound
    AL.Reinitialize
    AL.Append E
    AL.Append E
    AL.Append E
    On Error Resume Next
    AL.Remove 3
    If Err.NUMBER <> 9 Then
        TestRemove = False
        Debug.Print "Invalid upper bound"
    End If

    Debug.Print "TestRemove: " & TestRemove

End Function

Private Function TestRemoveRange() As Boolean

    TestRemoveRange = True

    Dim AL As clsArrayListUDT
    Set AL = New clsArrayListUDT

    Dim E As TExample
    E.Message = "A"
    
    'Empty
    On Error Resume Next
    AL.RemoveRange 0, 1
    If Err.NUMBER <> 9 Then
        TestRemoveRange = False
        Debug.Print "Empty"
    End If
    On Error GoTo 0

    'Non empty
        'Start
        AL.Reinitialize
        AL.Append E
        AL.Append E
        AL.Append E
        AL.RemoveRange 0, 1
        If AL.Size <> 1 Then
            TestRemoveRange = False
            Debug.Print "Start"
        End If
        'Middle
        AL.Reinitialize
        AL.Append E
        AL.Append E
        AL.Append E
        AL.Append E
        AL.RemoveRange 1, 2
        If AL.Size <> 2 Then
            TestRemoveRange = False
            Debug.Print "Middle"
        End If
        'End
        AL.Reinitialize
        AL.Append E
        AL.Append E
        AL.Append E
        AL.RemoveRange 1, 2
        If AL.Size <> 1 Then
            TestRemoveRange = False
            Debug.Print "End"
        End If
        
    'Invalid lower > upper
    AL.Reinitialize
    AL.Append E
    AL.Append E
    AL.Append E
    On Error Resume Next
    AL.RemoveRange 1, 0
    If Err.NUMBER <> 5 Then
        TestRemoveRange = False
        Debug.Print "Invalid lower > upper"
    End If
    On Error GoTo 0

    'Invalid lower bound
    AL.Reinitialize
    AL.Append E
    AL.Append E
    AL.Append E
    On Error Resume Next
    AL.RemoveRange -1, 0
    If Err.NUMBER <> 9 Then
        TestRemoveRange = False
        Debug.Print "Invalid lower bound"
    End If
    On Error GoTo 0

    'Invalid upper bound
    AL.Reinitialize
    AL.Append E
    AL.Append E
    AL.Append E
    On Error Resume Next
    AL.RemoveRange 2, 3
    If Err.NUMBER <> 9 Then
        TestRemoveRange = False
        Debug.Print "Invalid upper bound"
    End If
    On Error GoTo 0

    Debug.Print "TestRemoveRange: " & TestRemoveRange

End Function

Private Function TestClear() As Boolean

    TestClear = True
    
    Dim AL As clsArrayListUDT
    Set AL = New clsArrayListUDT
    
    Dim E As TExample
    E.Message = "A"
    
    'Empty
    AL.Clear
    If AL.Size <> 0 Then
        TestClear = False
        Debug.Print "Empty"
    End If
    
    'One
    AL.Append E
    AL.Clear
    If AL.Size <> 0 Then
        TestClear = False
        Debug.Print "One"
    End If

    'Multiple
    AL.Append E
    AL.Append E
    AL.Append E
    AL.Clear
    If AL.Size <> 0 Then
        TestClear = False
        Debug.Print "Multiple"
    End If
    
    'Capacity persists
    Dim i&
    For i = 1 To 11
        AL.Append E
    Next i
    AL.Clear
    If AL.Capacity <> DEFAULT_CAPACITY * 2 + 2 Then
        TestClear = False
        Debug.Print "Capacity persists"
    End If
    
    Debug.Print "TestClear: " & TestClear

End Function

Private Function TestReverse() As Boolean

    TestReverse = True
    
    Dim AL As clsArrayListUDT
    Set AL = New clsArrayListUDT
    
    Dim E As TExample
    E.Message = "A"
    
    'Empty
    On Error Resume Next
    AL.Reverse
    If Err.NUMBER <> 0 Then
        TestReverse = False
        Debug.Print "Empty"
    End If
    On Error GoTo 0
    
    'One
    AL.Append E
    AL.Reverse
    If AL.Item(0).Message <> "A" Then
        TestReverse = False
        Debug.Print "One"
    End If
    
    'Multiple
    E.Message = "B"
    AL.Append E
    E.Message = "C"
    AL.Append E
    AL.Reverse
    If AL.Item(0).Message & AL.Item(1).Message & AL.Item(2).Message <> "CBA" Then
        TestReverse = False
        Debug.Print "Multiple"
    End If
    
    'Even
    AL.Reinitialize
    E.Message = "A"
    AL.Append E
    E.Message = "B"
    AL.Append E
    E.Message = "C"
    AL.Append E
    E.Message = "D"
    AL.Append E
    AL.Reverse
    If AL.Item(0).Message & AL.Item(1).Message & _
    AL.Item(2).Message & AL.Item(3).Message <> "DCBA" Then
        TestReverse = False
        Debug.Print "Even"
    End If
    
    'Odd
    AL.Reinitialize
    E.Message = "A"
    AL.Append E
    E.Message = "B"
    AL.Append E
    E.Message = "C"
    AL.Append E
    AL.Reverse
    If AL.Item(0).Message & AL.Item(1).Message & _
    AL.Item(2).Message <> "CBA" Then
        TestReverse = False
        Debug.Print "Odd"
    End If
    
    Debug.Print "TestReverse: " & TestReverse

End Function

Private Function TestToArray() As Boolean

    TestToArray = True
    
    Dim AL As clsArrayListUDT
    Set AL = New clsArrayListUDT
    
    Dim E As TExample
    
    'Empty
    Dim Arr() As TExample
    Arr = AL.ToArray
    On Error Resume Next
    Debug.Print LBound(Arr)
    If Err.NUMBER <> 9 Then
        TestToArray = False
        Debug.Print "Empty"
    End If
    On Error GoTo 0
    
    'One
    AL.Append E
    Arr = AL.ToArray
    If UBound(Arr) - LBound(Arr) + 1 <> AL.Size Then
        TestToArray = False
        Debug.Print "One"
    End If
    If Arr(0).Message <> AL.Item(0).Message Or _
    Arr(0).Flag <> AL.Item(0).Flag Then
        TestToArray = False
        Debug.Print "One"
    End If
    
    'Multiple
    E.Message = "A"
    E.Flag = True
    AL.Append E
    E.Message = "B"
    E.Flag = False
    AL.Append E
    Arr = AL.ToArray
    If UBound(Arr) - LBound(Arr) + 1 <> AL.Size Then
        TestToArray = False
        Debug.Print "Multiple"
    End If
    If Arr(0).Message <> AL.Item(0).Message Or _
    Arr(0).Flag <> AL.Item(0).Flag Or _
    Arr(1).Message <> AL.Item(1).Message Or _
    Arr(1).Flag <> AL.Item(1).Flag Then
        TestToArray = False
        Debug.Print "Multiple"
    End If
    
    Debug.Print "TestToArray: " & TestToArray

End Function
