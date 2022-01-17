Attribute VB_Name = "testArrayListString"
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
'  Module Name: testArrayListString
'  Module Description: Unit tests for clsArrayListString class.
'  Module Version: 1.0
'  Module License: MIT
'  Module Author: Peter Roach; PeterRoach.Code@gmail.com
'  Module Contents:
'        TestArrayListString
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
'        TestRemoveFirst
'        TestRemoveLast
'        TestRemoveAll
'        TestReplaceAll
'        TestIndexOf
'        TestLastIndexOf
'        TestCount
'        TestContains
'        TestClear
'        TestReverse
'        TestSort
'        TestToStringArray
'        TestJoinString


Private Const DEFAULT_CAPACITY& = 10


'Example Usage=========================================================
'======================================================================

Public Sub Example()

    Dim AL As clsArrayListString
    Set AL = New clsArrayListString
    
    AL.Append "A"
    AL.Append "B"
    AL.Append "C"

    Dim Arr$(0 To 2)
    Arr(0) = "A"
    Arr(0) = "B"
    Arr(0) = "C"
    AL.AppendArray Arr
    
    Dim AL1 As clsArrayListString
    Set AL1 = New clsArrayListString
    AL1.Append "A"
    AL1.Append "B"
    AL1.Append "C"
    AL.AppendArrayList AL1
    
    AL.Insert 0, "A"
    AL.InsertArray 0, Arr
    AL.InsertArrayList 0, AL1
    
    AL.Remove 0
    AL.RemoveRange 0, 1
    AL.RemoveFirst "B"
    AL.RemoveLast "B"
    AL.RemoveAll "C"
    AL.ReplaceAll "D", "Z"
    
    Debug.Print AL.Contains("A")
    Debug.Print AL.Count("A")
    Debug.Print AL.IndexOf("A")
    Debug.Print AL.LastIndexOf("A")
    
    AL.Reverse
    AL.Sort
    
    AL.Clear
    AL.Reinitialize
    
    AL.Append "A"
    AL.Append "B"
    AL.Append "C"
    
    Debug.Print AL.JoinString()
    Debug.Print AL.JoinString(",")
    
    Dim Arr1$()
    Arr1 = AL.ToStringArray()
    
End Sub


'Unit Tests============================================================
'======================================================================

Public Function TestArrayListString() As Boolean
    
    TestArrayListString = _
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
        TestRemoveFirst And _
        TestRemoveLast And _
        TestRemoveAll And _
        TestReplaceAll And _
        TestIndexOf And _
        TestLastIndexOf And _
        TestCount And _
        TestContains
    TestArrayListString = TestArrayListString And _
        TestClear And _
        TestReverse And _
        TestToStringArray And _
        TestJoinString And _
        TestSort

    Debug.Print "TestArrayListString: " & TestArrayListString
    
End Function

Private Function TestCapacity() As Boolean

    TestCapacity = True
    
    Dim AL As clsArrayListString
    Set AL = New clsArrayListString
    
    'Causes compiler error - Test Passed
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
    AL.Append "A"
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
        AL.Append "A"
    Next i
    If AL.Capacity <> DEFAULT_CAPACITY * 2 + 2 Then
        TestCapacity = False
        Debug.Print "Append"
    End If
    
    'AppendArray
    AL.Reinitialize
    Dim Arr$(0 To 10)
    AL.AppendArray Arr
    If AL.Capacity <> DEFAULT_CAPACITY * 2 + 2 Then
        TestCapacity = False
        Debug.Print "AppendArray"
    End If
    
    'AppendArrayList
    AL.Reinitialize
    Dim AL1 As clsArrayListString
    Set AL1 = New clsArrayListString
    AL1.AppendArray Arr
    AL.AppendArrayList AL1
    If AL.Capacity <> DEFAULT_CAPACITY * 2 + 2 Then
        TestCapacity = False
        Debug.Print "AppendArrayList"
    End If
    
    'Insert
    AL.Reinitialize
    For i = 1 To 11
        AL.Insert 0, "A"
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
    AL.Append "A"
    AL.Remove 0
    If AL.Capacity <> DEFAULT_CAPACITY Then
        TestCapacity = False
        Debug.Print "Remove"
    End If
    
    'RemoveRange
    AL.Reinitialize
    AL.Append "A"
    AL.Append "B"
    AL.Append "C"
    AL.RemoveRange 0, 1
    If AL.Capacity <> DEFAULT_CAPACITY Then
        TestCapacity = False
        Debug.Print "RemoveRange"
    End If
    
    'RemoveFirst
    AL.Reinitialize
    AL.Append "A"
    AL.Append "B"
    AL.Append "A"
    AL.RemoveFirst "A"
    If AL.Capacity <> DEFAULT_CAPACITY Then
        TestCapacity = False
        Debug.Print "RemoveFirst"
    End If
    
    'RemoveLast
    AL.Reinitialize
    AL.Append "A"
    AL.Append "B"
    AL.Append "A"
    AL.RemoveLast "A"
    If AL.Capacity <> DEFAULT_CAPACITY Then
        TestCapacity = False
        Debug.Print "RemoveLast"
    End If
    
    'RemoveAll
    AL.Reinitialize
    AL.Append "A"
    AL.Append "B"
    AL.Append "A"
    AL.RemoveLast "A"
    If AL.Capacity <> DEFAULT_CAPACITY Then
        TestCapacity = False
        Debug.Print "RemoveAll"
    End If
    
    'ReplaceAll
    AL.Reinitialize
    AL.Append "A"
    AL.Append "B"
    AL.Append "A"
    AL.ReplaceAll "A", "C"
    If AL.Capacity <> DEFAULT_CAPACITY Then
        TestCapacity = False
        Debug.Print "ReplaceAll"
    End If
    
    'IndexOf
    AL.Reinitialize
    AL.Append "A"
    AL.Append "B"
    AL.Append "A"
    If AL.IndexOf("A") <> 0 Or AL.Capacity <> DEFAULT_CAPACITY Then
        TestCapacity = False
        Debug.Print "IndexOf"
    End If
    
    'LastIndexOf
    AL.Reinitialize
    AL.Append "A"
    AL.Append "B"
    AL.Append "A"
    If AL.LastIndexOf("A") <> 2 Or AL.Capacity <> DEFAULT_CAPACITY Then
        TestCapacity = False
        Debug.Print "LastIndexOf"
    End If
    
    'Count
    AL.Reinitialize
    AL.Append "A"
    AL.Append "B"
    AL.Append "A"
    If AL.Count("A") <> 2 Or AL.Capacity <> DEFAULT_CAPACITY Then
        TestCapacity = False
        Debug.Print "Count"
    End If
    
    'Contains
    AL.Reinitialize
    AL.Append "A"
    AL.Append "B"
    AL.Append "A"
    If AL.Contains("A") <> True Or AL.Capacity <> DEFAULT_CAPACITY Then
        TestCapacity = False
        Debug.Print "Contains"
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
    AL.Append "A"
    AL.Append "B"
    AL.Append "C"
    AL.Reverse
    If AL.JoinString <> "CBA" Or AL.Capacity <> DEFAULT_CAPACITY Then
        TestCapacity = False
        Debug.Print "Reverse"
    End If
    
    'ToStringArray
    AL.Reinitialize
    AL.Append "A"
    AL.Append "B"
    AL.Append "C"
    If AL.JoinString <> Join(AL.ToStringArray, "") Or AL.Capacity <> DEFAULT_CAPACITY Then
        TestCapacity = False
        Debug.Print "ToStringArray"
    End If
    
    'JoinString
    AL.Reinitialize
    AL.Append "A"
    AL.Append "B"
    AL.Append "C"
    If AL.JoinString <> "ABC" Or AL.Capacity <> DEFAULT_CAPACITY Then
        TestCapacity = False
        Debug.Print "JoinString"
    End If
    
    'Sort
    AL.Reinitialize
    AL.Append "C"
    AL.Append "B"
    AL.Append "A"
    AL.Sort
    If AL.JoinString <> "ABC" Or AL.Capacity <> DEFAULT_CAPACITY Then
        TestCapacity = False
        Debug.Print "Sort"
    End If

    Debug.Print "TestCapacity: " & TestCapacity

End Function

Private Function TestSize() As Boolean

    TestSize = True
    
    Dim AL As clsArrayListString
    Set AL = New clsArrayListString
    
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
    AL.Append "A"
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
    AL.Append "A"
    If AL.Size <> 1 Then
        TestSize = False
        Debug.Print "Append"
    End If
    
    'AppendArray
    Dim Arr$(0 To 2)
    AL.AppendArray Arr
    If AL.Size <> 4 Then
        TestSize = False
        Debug.Print "AppendArray"
    End If
    
    'AppendArrayList
    Dim AL1 As clsArrayListString
    Set AL1 = New clsArrayListString
    AL1.AppendArray Arr
    AL.AppendArrayList AL1
    If AL.Size <> 7 Then
        TestSize = False
        Debug.Print "AppendArrayList"
    End If
    
    'Insert
    AL.Insert 0, "A"
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
    
    'RemoveFirst
    AL.Reinitialize
    AL.Append "A"
    AL.Append "B"
    AL.Append "A"
    AL.RemoveFirst "A"
    If AL.Size <> 2 Then
        TestSize = False
        Debug.Print "RemoveFirst"
    End If
    
    'RemoveLast
    AL.Reinitialize
    AL.Append "A"
    AL.Append "B"
    AL.Append "A"
    AL.RemoveLast "A"
    If AL.Size <> 2 Then
        TestSize = False
        Debug.Print "RemoveLast"
    End If
    
    'RemoveAll
    AL.Reinitialize
    AL.Append "A"
    AL.Append "B"
    AL.Append "A"
    AL.RemoveAll "A"
    If AL.Size <> 1 Then
        TestSize = False
        Debug.Print "RemoveAll"
    End If
    
    'ReplaceAll
    AL.Reinitialize
    AL.Append "A"
    AL.Append "B"
    AL.Append "A"
    AL.ReplaceAll "A", "B"
    If AL.Size <> 3 Then
        TestSize = False
        Debug.Print "ReplaceAll"
    End If
    
    'IndexOf
    AL.Reinitialize
    AL.Append "A"
    AL.Append "B"
    AL.Append "A"
    If AL.IndexOf("A") <> 0 Or AL.Size <> 3 Then
        TestSize = False
        Debug.Print "IndexOf"
    End If
    
    'LastIndexOf
    AL.Reinitialize
    AL.Append "A"
    AL.Append "B"
    AL.Append "A"
    If AL.LastIndexOf("A") <> 2 Or AL.Size <> 3 Then
        TestSize = False
        Debug.Print "LastIndexOf"
    End If
    
    'Count
    AL.Reinitialize
    AL.Append "A"
    AL.Append "B"
    AL.Append "A"
    If AL.Count("A") <> 2 Or AL.Size <> 3 Then
        TestSize = False
        Debug.Print "Count"
    End If
    
    'Contains
    AL.Reinitialize
    AL.Append "A"
    AL.Append "B"
    AL.Append "A"
    If AL.Contains("A") <> True Or AL.Size <> 3 Then
        TestSize = False
        Debug.Print "Contains"
    End If
    
    'Clear
    AL.Clear
    If AL.Size <> 0 Then
        TestSize = False
        Debug.Print "Clear"
    End If
    
    'Reverse
    AL.Reinitialize
    AL.Append "A"
    AL.Append "B"
    AL.Append "C"
    AL.Reverse
    If AL.JoinString <> "CBA" Or AL.Size <> 3 Then
        TestSize = False
        Debug.Print "Reverse"
    End If
    
    'ToStringArray
    AL.Reinitialize
    AL.Append "A"
    AL.Append "B"
    AL.Append "C"
    If AL.JoinString <> Join(AL.ToStringArray, "") Or AL.Size <> 3 Then
        TestSize = False
        Debug.Print "ToStringArray"
    End If
    
    'JoinString
    AL.Reinitialize
    AL.Append "A"
    AL.Append "B"
    AL.Append "C"
    If AL.JoinString <> "ABC" Or AL.Size <> 3 Then
        TestSize = False
        Debug.Print "JoinString"
    End If
    
    'Sort
    AL.Reinitialize
    AL.Append "C"
    AL.Append "B"
    AL.Append "A"
    AL.Sort
    If AL.JoinString <> "ABC" Or AL.Size <> 3 Then
        TestSize = False
        Debug.Print "Sort"
    End If
    
    Debug.Print "TestSize: " & TestSize

End Function

Private Function TestItem() As Boolean

    TestItem = True

    Dim AL As clsArrayListString
    Set AL = New clsArrayListString

    'Empty
        'Get
        On Error Resume Next
        Debug.Print AL.Item(0)
        If Err.Number <> 9 Then
            TestItem = False
            Debug.Print "Empty Get"
        End If
        On Error GoTo 0
        'Let
        On Error Resume Next
        AL.Item(0) = "A"
        If Err.Number <> 9 Then
            TestItem = False
            Debug.Print "Empty Let"
        End If
        On Error GoTo 0

    'Non-Empty
    AL.Append "A"
        'Get
        If AL.Item(0) <> "A" Then
            TestItem = False
            Debug.Print "Non-Empty Get"
        End If
        'Let
        AL.Item(0) = "B"
        If AL.Item(0) <> "B" Then
            TestItem = False
            Debug.Print "Non-Empty Let"
        End If

    'Invalid lower bound
    On Error Resume Next
    Debug.Print AL.Item(-1)
    If Err.Number <> 9 Then
        TestItem = False
        Debug.Print "Invalid lower bound"
    End If
    On Error GoTo 0

    'Invalid upper bound
    On Error Resume Next
    Debug.Print AL.Item(1)
    If Err.Number <> 9 Then
        TestItem = False
        Debug.Print "Invalid upper bound"
    End If
    On Error GoTo 0

    Debug.Print "TestItem: " & TestItem

End Function

Private Function TestGrowCapacity() As Boolean

    TestGrowCapacity = True
    
    Dim AL As clsArrayListString
    Set AL = New clsArrayListString
    
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
    
    Dim AL As clsArrayListString
    Set AL = New clsArrayListString
    
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
    
    'Size
    AL.Reinitialize
    AL.Append "A"
    AL.Append "B"
    AL.Append "C"
    AL.ShrinkCapacity 3
    If AL.Size <> 3 Then
        TestShrinkCapacity = False
        Debug.Print "Size"
    End If
    
    'Less than size
    AL.Reinitialize
    AL.Append "A"
    AL.Append "B"
    AL.Append "C"
    AL.ShrinkCapacity 2
    If AL.Size <> 3 Then
        TestShrinkCapacity = False
        Debug.Print "Less than size"
    End If
    
    'Invalid 0
    AL.Reinitialize
    On Error Resume Next
    AL.ShrinkCapacity 0
    If Err.Number <> 5 Then
        TestShrinkCapacity = False
        Debug.Print "Invalid 0"
    End If
    On Error GoTo 0
    
    'Invalid negative
    AL.Reinitialize
    On Error Resume Next
    AL.ShrinkCapacity -1
    If Err.Number <> 5 Then
        TestShrinkCapacity = False
        Debug.Print "Invalid negative"
    End If
    On Error GoTo 0
    
    Debug.Print "TestShrinkCapacity: " & TestShrinkCapacity

End Function

Private Function TestEnsureCapacity() As Boolean

    TestEnsureCapacity = True
    
    Dim AL As clsArrayListString
    Set AL = New clsArrayListString
    
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

    Dim AL As clsArrayListString
    Set AL = New clsArrayListString

    'Size 0
    AL.TrimToSize
    If AL.Capacity <> DEFAULT_CAPACITY Or AL.Size <> 0 Then
        TestTrimToSize = False
        Debug.Print "Size 0"
    End If

    'Size 1
    AL.Append "A"
    AL.TrimToSize
    If AL.Capacity <> 1 Or AL.Size <> 1 Then
        TestTrimToSize = False
        Debug.Print "Size 1"
    End If

    'Size > 1
    AL.Reinitialize
    AL.Append "A"
    AL.Append "B"
    AL.Append "C"
    AL.TrimToSize
    If AL.Capacity <> 3 Or AL.Size <> 3 Then
        TestTrimToSize = False
        Debug.Print "Size > 1"
    End If

    'Reinitialize
    AL.Reinitialize
    AL.GrowCapacity DEFAULT_CAPACITY * 2 + 2
    AL.Append "A"
    AL.Append "B"
    AL.Append "C"
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
    
    Dim AL As clsArrayListString
    Set AL = New clsArrayListString
    
    'Default
    AL.GrowCapacity DEFAULT_CAPACITY * 2 + 2
    AL.Append "A"
    AL.Append "B"
    AL.Append "C"
    AL.Reinitialize
    If AL.Capacity <> DEFAULT_CAPACITY Or AL.Size <> 0 Then
        TestReinitialize = False
        Debug.Print "Default"
    End If
    
    'Less than default
    AL.Reinitialize
    AL.GrowCapacity DEFAULT_CAPACITY * 2 + 2
    AL.Append "A"
    AL.Append "B"
    AL.Append "C"
    AL.Reinitialize DEFAULT_CAPACITY - 1
    If AL.Capacity <> DEFAULT_CAPACITY - 1 Or AL.Size <> 0 Then
        TestReinitialize = False
        Debug.Print "Less than default"
    End If
    
    'More than default
    AL.Reinitialize
    AL.GrowCapacity DEFAULT_CAPACITY * 2 + 2
    AL.Append "A"
    AL.Append "B"
    AL.Append "C"
    AL.Reinitialize DEFAULT_CAPACITY + 1
    If AL.Capacity <> DEFAULT_CAPACITY + 1 Or AL.Size <> 0 Then
        TestReinitialize = False
        Debug.Print "More than default"
    End If
    
    'Invalid 0
    AL.Reinitialize
    AL.GrowCapacity DEFAULT_CAPACITY * 2 + 2
    AL.Append "A"
    AL.Append "B"
    AL.Append "C"
    On Error Resume Next
    AL.Reinitialize 0
    If Err.Number <> 5 Then
        TestReinitialize = False
        Debug.Print "Invalid 0"
    End If
    On Error GoTo 0
    
    'Invalid negative
    AL.Reinitialize
    AL.GrowCapacity DEFAULT_CAPACITY * 2 + 2
    AL.Append "A"
    AL.Append "B"
    AL.Append "C"
    On Error Resume Next
    AL.Reinitialize 0
    If Err.Number <> 5 Then
        TestReinitialize = False
        Debug.Print "Invalid negative"
    End If
    On Error GoTo 0
    
    Debug.Print "TestReinitialize: " & TestReinitialize

End Function

Private Function TestAppend() As Boolean

    TestAppend = True

    Dim AL As clsArrayListString
    Set AL = New clsArrayListString

    'Empty
    AL.Append "A"
    If AL.JoinString <> "A" Or AL.Size <> 1 Then
        TestAppend = False
        Debug.Print "Empty"
    End If

    'Non empty
    AL.Append "B"
    If AL.JoinString <> "AB" Or AL.Size <> 2 Then
        TestAppend = False
        Debug.Print "Non Empty"
    End If

    'Until capacity
    Dim i&
    AL.Reinitialize
    For i = 1 To 11
        AL.Append "A"
    Next i
    If AL.JoinString <> String$(11, "A") Or _
    AL.Capacity <> DEFAULT_CAPACITY * 2 + 2 Or _
    AL.Size <> 11 Then
        TestAppend = False
        Debug.Print "Until capacity"
    End If

    Debug.Print "TestAppend: " & TestAppend

End Function

Private Function TestAppendArray() As Boolean

    TestAppendArray = True

    Dim AL As clsArrayListString
    Set AL = New clsArrayListString

    Dim Arr$(0 To 2)

    'Empty
    Arr(0) = "A"
    Arr(1) = "B"
    Arr(2) = "C"
    AL.AppendArray Arr
    If AL.JoinString <> "ABC" Or AL.Size <> 3 Then
        TestAppendArray = False
        Debug.Print "Empty"
    End If
    
    'Non empty
    Arr(0) = "X"
    Arr(1) = "Y"
    Arr(2) = "Z"
    AL.AppendArray Arr
    If AL.JoinString <> "ABCXYZ" Or AL.Size <> 6 Then
        TestAppendArray = False
        Debug.Print "Non empty"
    End If
    
    'Until capacity
    Arr(0) = "L"
    Arr(1) = "M"
    Arr(2) = "N"
    AL.AppendArray Arr
    Arr(0) = "T"
    Arr(1) = "U"
    Arr(2) = "V"
    AL.AppendArray Arr
    If AL.JoinString <> "ABCXYZLMNTUV" Or _
    AL.Size <> 12 Or _
    AL.Capacity <> DEFAULT_CAPACITY * 2 + 2 Then
        TestAppendArray = False
        Debug.Print "Until capacity"
    End If
    
    'Empty Array
    Dim Arr1$()
    On Error Resume Next
    AL.AppendArray Arr1
    If Err.Number <> 9 Then
        TestAppendArray = False
        Debug.Print "Empty Array"
    End If
    On Error GoTo 0
    
    Debug.Print "TestAppendArray: " & TestAppendArray

End Function

Private Function TestAppendArrayList() As Boolean

    TestAppendArrayList = True
    
    Dim AL As clsArrayListString
    Set AL = New clsArrayListString
    
    Dim AL1 As clsArrayListString
    Set AL1 = New clsArrayListString

    'Empty
    AL1.Append "A"
    AL1.Append "B"
    AL1.Append "C"
    AL.AppendArrayList AL1
    If AL.JoinString <> "ABC" Or AL.Size <> 3 Then
        TestAppendArrayList = False
        Debug.Print "Empty"
    End If
    
    'Non empty
    AL1.Reinitialize
    AL1.Append "X"
    AL1.Append "Y"
    AL1.Append "Z"
    AL.AppendArrayList AL1
    If AL.JoinString <> "ABCXYZ" Or AL.Size <> 6 Then
        TestAppendArrayList = False
        Debug.Print "Non empty"
    End If
    
    'Until capacity
    AL1.Reinitialize
    AL1.Append "L"
    AL1.Append "M"
    AL1.Append "N"
    AL.AppendArrayList AL1
    AL1.Reinitialize
    AL1.Append "T"
    AL1.Append "U"
    AL1.Append "V"
    AL.AppendArrayList AL1
    If AL.JoinString <> "ABCXYZLMNTUV" Or _
    AL.Size <> 12 Or _
    AL.Capacity <> DEFAULT_CAPACITY * 2 + 2 Then
        TestAppendArrayList = False
        Debug.Print "Until capacity"
    End If
    
    'Empty ArrayList
    AL.Reinitialize
    AL1.Reinitialize
    AL.AppendArrayList AL1
    If AL.JoinString <> "" Or _
    AL.Capacity <> DEFAULT_CAPACITY Or _
    AL.Size <> 0 Then
        TestAppendArrayList = False
        Debug.Print "Empty ArrayList"
    End If
    
    Debug.Print "TestAppendArrayList: " & TestAppendArrayList

End Function

Private Function TestInsert() As Boolean

    TestInsert = True
    
    Dim AL As clsArrayListString
    Set AL = New clsArrayListString
    
    'Empty
    AL.Insert 0, "A"
    If AL.JoinString <> "A" Or AL.Size <> 1 Then
        TestInsert = False
        Debug.Print "Empty"
    End If
    
    'Non empty
        'Start
        AL.Insert 0, "B"
        If AL.JoinString <> "BA" Or AL.Size <> 2 Then
            TestInsert = False
            Debug.Print "Start"
        End If
        'Middle
        AL.Insert 1, "C"
        If AL.JoinString <> "BCA" Or AL.Size <> 3 Then
            TestInsert = False
            Debug.Print "Middle"
        End If
        'End
        AL.Insert 3, "D"
        If AL.JoinString <> "BCAD" Or AL.Size <> 4 Then
            TestInsert = False
            Debug.Print "End"
        End If
        
    'Invalid lower bound
    On Error Resume Next
    AL.Insert -1, "Z"
    If Err.Number <> 9 Then
        TestInsert = False
        Debug.Print "Invalid lower bound"
    End If
    On Error GoTo 0
    
    'Invalid upper bound
    On Error Resume Next
    AL.Insert 5, "Z"
    If Err.Number <> 9 Then
        TestInsert = False
        Debug.Print "Invalid upper bound"
    End If
    On Error GoTo 0
    
    Debug.Print "TestInsert: " & TestInsert

End Function

Private Function TestInsertArray() As Boolean

    TestInsertArray = True

    Dim AL As clsArrayListString
    Set AL = New clsArrayListString

    Dim Arr$(0 To 2)

    'Empty
    Arr(0) = "A"
    Arr(1) = "B"
    Arr(2) = "C"
    AL.InsertArray 0, Arr
    If AL.JoinString <> "ABC" Or AL.Size <> 3 Or AL.Capacity <> DEFAULT_CAPACITY Then
        TestInsertArray = False
        Debug.Print "Empty"
    End If

    'Non empty
        'Start
        Arr(0) = "D"
        Arr(1) = "E"
        Arr(2) = "F"
        AL.InsertArray 0, Arr
        If AL.JoinString <> "DEFABC" Or AL.Size <> 6 Or AL.Capacity <> DEFAULT_CAPACITY Then
            TestInsertArray = False
            Debug.Print "Start"
        End If
        'Middle
        Arr(0) = "G"
        Arr(1) = "H"
        Arr(2) = "I"
        AL.InsertArray 3, Arr
        If AL.JoinString <> "DEFGHIABC" Or AL.Size <> 9 Or AL.Capacity <> DEFAULT_CAPACITY Then
            TestInsertArray = False
            Debug.Print "Middle"
        End If
        'End
        Arr(0) = "J"
        Arr(1) = "K"
        Arr(2) = "L"
        AL.InsertArray 9, Arr
        If AL.JoinString <> "DEFGHIABCJKL" Or AL.Size <> 12 Or AL.Capacity <> DEFAULT_CAPACITY * 2 + 2 Then
            TestInsertArray = False
            Debug.Print "End"
        End If

    'Invalid lower bound
    On Error Resume Next
    AL.InsertArray -1, Arr
    If Err.Number <> 9 Then
        TestInsertArray = False
        Debug.Print "Invalid lower bound"
    End If
    On Error GoTo 0

    'Invalid upper bound
    On Error Resume Next
    AL.InsertArray 13, Arr
    If Err.Number <> 9 Then
        TestInsertArray = False
        Debug.Print "Invalid upper bound"
    End If
    On Error GoTo 0

    'Empty array
    AL.Reinitialize
    Dim Arr1$()
    On Error Resume Next
    AL.InsertArray 0, Arr1
    If Err.Number <> 9 Then
        TestInsertArray = False
        Debug.Print "Empty array"
    End If
    On Error GoTo 0

    Debug.Print "TestInsertArray: " & TestInsertArray

End Function

Private Function TestInsertArrayList() As Boolean

    TestInsertArrayList = True

    Dim AL As clsArrayListString
    Set AL = New clsArrayListString
    
    Dim AL1 As clsArrayListString
    Set AL1 = New clsArrayListString
    
    'Empty
    AL1.Reinitialize
    AL1.Append "A"
    AL1.Append "B"
    AL1.Append "C"
    AL.InsertArrayList 0, AL1
    If AL.JoinString <> "ABC" Or AL.Size <> 3 Or AL.Capacity <> DEFAULT_CAPACITY Then
        TestInsertArrayList = False
        Debug.Print "Empty"
    End If

    'Non empty
        'Start
        AL1.Reinitialize
        AL1.Append "D"
        AL1.Append "E"
        AL1.Append "F"
        AL.InsertArrayList 0, AL1
        If AL.JoinString <> "DEFABC" Or AL.Size <> 6 Or AL.Capacity <> DEFAULT_CAPACITY Then
            TestInsertArrayList = False
            Debug.Print "Start"
        End If
        'Middle
        AL1.Reinitialize
        AL1.Append "G"
        AL1.Append "H"
        AL1.Append "I"
        AL.InsertArrayList 3, AL1
        If AL.JoinString <> "DEFGHIABC" Or AL.Size <> 9 Or AL.Capacity <> DEFAULT_CAPACITY Then
            TestInsertArrayList = False
            Debug.Print "Middle"
        End If
        'End
        AL1.Reinitialize
        AL1.Append "J"
        AL1.Append "K"
        AL1.Append "L"
        AL.InsertArrayList 9, AL1
        If AL.JoinString <> "DEFGHIABCJKL" Or AL.Size <> 12 Or AL.Capacity <> DEFAULT_CAPACITY * 2 + 2 Then
            TestInsertArrayList = False
            Debug.Print "End"
        End If

    'Invalid lower bound
    On Error Resume Next
    AL.InsertArrayList -1, AL1
    If Err.Number <> 9 Then
        TestInsertArrayList = False
        Debug.Print "Invalid lower bound"
    End If
    On Error GoTo 0

    'Invalid upper bound
    On Error Resume Next
    AL.InsertArrayList 13, AL1
    If Err.Number <> 9 Then
        TestInsertArrayList = False
        Debug.Print "Invalid upper bound"
    End If
    On Error GoTo 0

    'Empty ArrayList
    AL.Reinitialize
    AL1.Reinitialize
    AL.InsertArrayList 0, AL1
    If AL.JoinString <> "" Or AL.Size <> 0 Or AL.Capacity <> DEFAULT_CAPACITY Then
        TestInsertArrayList = False
        Debug.Print "Empty ArrayList"
    End If

    Debug.Print "TestInsertArrayList: " & TestInsertArrayList

End Function

Private Function TestRemove() As Boolean

    TestRemove = True

    Dim AL As clsArrayListString
    Set AL = New clsArrayListString

    'Empty
    On Error Resume Next
    AL.Remove 0
    If Err.Number <> 9 Then
        TestRemove = False
        Debug.Print "Empty"
    End If
    On Error GoTo 0
    
    'Non empty
    AL.Append "A"
    AL.Remove 0
    If AL.JoinString <> "" Or AL.Size <> 0 Then
        TestRemove = False
        Debug.Print "Non empty"
    End If

    'Start
    AL.Reinitialize
    AL.Append "A"
    AL.Append "B"
    AL.Append "C"
    AL.Remove 0
    If AL.JoinString <> "BC" Or AL.Size <> 2 Then
        TestRemove = False
        Debug.Print "Start"
    End If

    'Middle
    AL.Reinitialize
    AL.Append "A"
    AL.Append "B"
    AL.Append "C"
    AL.Remove 1
    If AL.JoinString <> "AC" Or AL.Size <> 2 Then
        TestRemove = False
        Debug.Print "Middle"
    End If

    'End
    AL.Reinitialize
    AL.Append "A"
    AL.Append "B"
    AL.Append "C"
    AL.Remove AL.Size - 1
    If AL.JoinString <> "AB" Or AL.Size <> 2 Then
        TestRemove = False
        Debug.Print "End"
    End If

    'Invalid lower bound
    AL.Reinitialize
    AL.Append "A"
    AL.Append "B"
    AL.Append "C"
    On Error Resume Next
    AL.Remove -1
    If Err.Number <> 9 Then
        TestRemove = False
        Debug.Print "Invalid lower bound"
    End If

    'Invalid upper bound
    AL.Reinitialize
    AL.Append "A"
    AL.Append "B"
    AL.Append "C"
    On Error Resume Next
    AL.Remove 3
    If Err.Number <> 9 Then
        TestRemove = False
        Debug.Print "Invalid upper bound"
    End If

    Debug.Print "TestRemove: " & TestRemove

End Function

Private Function TestRemoveRange() As Boolean

    TestRemoveRange = True

    Dim AL As clsArrayListString
    Set AL = New clsArrayListString

    'Empty
    On Error Resume Next
    AL.RemoveRange 0, 1
    If Err.Number <> 9 Then
        TestRemoveRange = False
        Debug.Print "Empty"
    End If
    On Error GoTo 0

    'Non empty
        'Start
        AL.Reinitialize
        AL.Append "A"
        AL.Append "B"
        AL.Append "C"
        AL.RemoveRange 0, 1
        If AL.JoinString <> "C" Or AL.Size <> 1 Then
            TestRemoveRange = False
            Debug.Print "Start"
        End If
        'Middle
        AL.Reinitialize
        AL.Append "A"
        AL.Append "B"
        AL.Append "C"
        AL.Append "D"
        AL.RemoveRange 1, 2
        If AL.JoinString <> "AD" Or AL.Size <> 2 Then
            TestRemoveRange = False
            Debug.Print "Middle"
        End If
        'End
        AL.Reinitialize
        AL.Append "A"
        AL.Append "B"
        AL.Append "C"
        AL.RemoveRange 1, 2
        If AL.JoinString <> "A" Or AL.Size <> 1 Then
            TestRemoveRange = False
            Debug.Print "End"
        End If
        

    'Invalid lower > upper
    AL.Reinitialize
    AL.Append "A"
    AL.Append "B"
    AL.Append "C"
    On Error Resume Next
    AL.RemoveRange 1, 0
    If Err.Number <> 5 Then
        TestRemoveRange = False
        Debug.Print "Invalid lower > upper"
    End If
    On Error GoTo 0

    'Invalid lower bound
    AL.Reinitialize
    AL.Append "A"
    AL.Append "B"
    AL.Append "C"
    On Error Resume Next
    AL.RemoveRange -1, 0
    If Err.Number <> 9 Then
        TestRemoveRange = False
        Debug.Print "Invalid lower bound"
    End If
    On Error GoTo 0

    'Invalid upper bound
    AL.Reinitialize
    AL.Append "A"
    AL.Append "B"
    AL.Append "C"
    On Error Resume Next
    AL.RemoveRange 2, 3
    If Err.Number <> 9 Then
        TestRemoveRange = False
        Debug.Print "Invalid upper bound"
    End If
    On Error GoTo 0

    Debug.Print "TestRemoveRange: " & TestRemoveRange

End Function

Private Function TestRemoveFirst() As Boolean

    TestRemoveFirst = True

    Dim AL As clsArrayListString
    Set AL = New clsArrayListString

    'Empty
    On Error Resume Next
    AL.RemoveFirst ""
    If Err.Number <> 0 Then
        TestRemoveFirst = False
        Debug.Print "Empty"
    End If
    On Error GoTo 0
    If AL.Size <> 0 Then
        TestRemoveFirst = False
        Debug.Print "Empty"
    End If

    'One
    AL.Append "A"
    AL.RemoveFirst "A"
    If AL.JoinString <> "" Or AL.Size <> 0 Then
        TestRemoveFirst = False
        Debug.Print "One"
    End If
    AL.Append "A"
    AL.RemoveFirst ""
    If AL.JoinString <> "A" Or AL.Size <> 1 Then
        TestRemoveFirst = False
        Debug.Print "One"
    End If
    
    'Multiple
    AL.Reinitialize
    AL.Append "A"
    AL.Append "B"
    AL.Append "A"
    AL.RemoveFirst "A"
    If AL.JoinString <> "BA" Or AL.Size <> 2 Then
       TestRemoveFirst = False
        Debug.Print "Multiple"
    End If
    
    'First
    AL.Reinitialize
    AL.Append "A"
    AL.Append "B"
    AL.Append "A"
    AL.RemoveFirst "A"
    If AL.JoinString <> "BA" Or AL.Size <> 2 Then
       TestRemoveFirst = False
        Debug.Print "First"
    End If

    'Middle
    AL.Reinitialize
    AL.Append "B"
    AL.Append "A"
    AL.Append "B"
    AL.Append "A"
    AL.RemoveFirst "A"
    If AL.JoinString <> "BBA" Or AL.Size <> 3 Then
       TestRemoveFirst = False
        Debug.Print "Middle"
    End If

    'End
    AL.Reinitialize
    AL.Append "B"
    AL.Append "B"
    AL.Append "A"
    AL.RemoveFirst "A"
    If AL.JoinString <> "BB" Or AL.Size <> 2 Then
       TestRemoveFirst = False
        Debug.Print "End"
    End If

    'Not there
    AL.Reinitialize
    AL.Append "A"
    AL.Append "B"
    AL.Append "C"
    AL.RemoveFirst "Z"
    If AL.JoinString <> "ABC" Or AL.Size <> 3 Then
       TestRemoveFirst = False
        Debug.Print "Not there"
    End If

    'CompareMethod
    AL.Reinitialize
    AL.Append "A"
    AL.Append "B"
    AL.Append "C"
    AL.RemoveFirst "a"
    If AL.JoinString <> "ABC" Or AL.Size <> 3 Then
       TestRemoveFirst = False
        Debug.Print "Binary"
    End If
    AL.RemoveFirst "a", vbTextCompare
    If AL.JoinString <> "BC" Or AL.Size <> 2 Then
       TestRemoveFirst = False
        Debug.Print "Text"
    End If
    
    Debug.Print "TestRemoveFirst: " & TestRemoveFirst

End Function

Private Function TestRemoveLast() As Boolean

    TestRemoveLast = True
    
    Dim AL As clsArrayListString
    Set AL = New clsArrayListString
    
    'Empty
    On Error Resume Next
    AL.RemoveLast ""
    If Err.Number <> 0 Then
        TestRemoveLast = False
        Debug.Print "Empty"
    End If
    On Error GoTo 0
    If AL.Size <> 0 Then
        TestRemoveLast = False
        Debug.Print "Empty"
    End If

    'One
    AL.Append "A"
    AL.RemoveLast "A"
    If AL.JoinString <> "" Or AL.Size <> 0 Then
        TestRemoveLast = False
        Debug.Print "One"
    End If
    AL.Append "A"
    AL.RemoveLast ""
    If AL.JoinString <> "A" Or AL.Size <> 1 Then
        TestRemoveLast = False
        Debug.Print "One"
    End If
    
    'Multiple
    AL.Reinitialize
    AL.Append "A"
    AL.Append "B"
    AL.Append "A"
    AL.RemoveLast "A"
    If AL.JoinString <> "AB" Or AL.Size <> 2 Then
       TestRemoveLast = False
        Debug.Print "Multiple"
    End If
    
    'First
    AL.Reinitialize
    AL.Append "A"
    AL.Append "B"
    AL.Append "C"
    AL.RemoveLast "A"
    If AL.JoinString <> "BC" Or AL.Size <> 2 Then
       TestRemoveLast = False
        Debug.Print "First"
    End If

    'Middle
    AL.Reinitialize
    AL.Append "A"
    AL.Append "B"
    AL.Append "A"
    AL.Append "B"
    AL.RemoveLast "A"
    If AL.JoinString <> "ABB" Or AL.Size <> 3 Then
       TestRemoveLast = False
        Debug.Print "Middle"
    End If

    'End
    AL.Reinitialize
    AL.Append "A"
    AL.Append "B"
    AL.Append "B"
    AL.Append "A"
    AL.RemoveLast "A"
    If AL.JoinString <> "ABB" Or AL.Size <> 3 Then
       TestRemoveLast = False
        Debug.Print "End"
    End If

    'Not there
    AL.Reinitialize
    AL.Append "A"
    AL.Append "B"
    AL.Append "C"
    AL.RemoveLast "Z"
    If AL.JoinString <> "ABC" Or AL.Size <> 3 Then
       TestRemoveLast = False
        Debug.Print "Not there"
    End If

    'CompareMethod
    AL.Reinitialize
    AL.Append "A"
    AL.Append "B"
    AL.Append "C"
    AL.RemoveLast "a"
    If AL.JoinString <> "ABC" Or AL.Size <> 3 Then
       TestRemoveLast = False
        Debug.Print "Binary"
    End If
    AL.RemoveLast "a", vbTextCompare
    If AL.JoinString <> "BC" Or AL.Size <> 2 Then
       TestRemoveLast = False
        Debug.Print "Text"
    End If
    
    Debug.Print "TestRemoveLast: " & TestRemoveLast

End Function

Private Function TestRemoveAll() As Boolean

    TestRemoveAll = True
    
    Dim AL As clsArrayListString
    Set AL = New clsArrayListString
    
    'Empty
    On Error Resume Next
    AL.RemoveAll ""
    If Err.Number <> 0 Then
        TestRemoveAll = False
        Debug.Print "Empty"
    End If
    On Error GoTo 0
    If AL.Size <> 0 Then
        TestRemoveAll = False
        Debug.Print "Empty"
    End If

    'One
    AL.Append "A"
    AL.RemoveAll "A"
    If AL.JoinString <> "" Or AL.Size <> 0 Then
        TestRemoveAll = False
        Debug.Print "One"
    End If
    AL.Append "A"
    AL.RemoveAll ""
    If AL.JoinString <> "A" Or AL.Size <> 1 Then
        TestRemoveAll = False
        Debug.Print "One"
    End If
    
    'Multiple
    AL.Reinitialize
    AL.Append "A"
    AL.Append "B"
    AL.Append "A"
    AL.RemoveAll "A"
    If AL.JoinString <> "B" Or AL.Size <> 1 Then
       TestRemoveAll = False
       Debug.Print "Multiple"
    End If
    
    'First
    AL.Reinitialize
    AL.Append "A"
    AL.Append "B"
    AL.Append "C"
    AL.RemoveAll "A"
    If AL.JoinString <> "BC" Or AL.Size <> 2 Then
       TestRemoveAll = False
       Debug.Print "First"
    End If

    'Middle
    AL.Reinitialize
    AL.Append "A"
    AL.Append "B"
    AL.Append "A"
    AL.Append "B"
    AL.RemoveAll "A"
    If AL.JoinString <> "BB" Or AL.Size <> 2 Then
       TestRemoveAll = False
       Debug.Print "Middle"
    End If

    'End
    AL.Reinitialize
    AL.Append "A"
    AL.Append "B"
    AL.Append "B"
    AL.Append "A"
    AL.RemoveAll "A"
    If AL.JoinString <> "BB" Or AL.Size <> 2 Then
       TestRemoveAll = False
       Debug.Print "End"
    End If

    'Not there
    AL.Reinitialize
    AL.Append "A"
    AL.Append "B"
    AL.Append "C"
    AL.RemoveAll "Z"
    If AL.JoinString <> "ABC" Or AL.Size <> 3 Then
       TestRemoveAll = False
       Debug.Print "Not there"
    End If

    'CompareMethod
    AL.Reinitialize
    AL.Append "A"
    AL.Append "B"
    AL.Append "C"
    AL.RemoveAll "a"
    If AL.JoinString <> "ABC" Or AL.Size <> 3 Then
       TestRemoveAll = False
       Debug.Print "Binary"
    End If
    AL.RemoveAll "a", vbTextCompare
    If AL.JoinString <> "BC" Or AL.Size <> 2 Then
       TestRemoveAll = False
       Debug.Print "Text"
    End If
    
    Debug.Print "TestRemoveAll: " & TestRemoveAll

End Function

Private Function TestReplaceAll() As Boolean

    TestReplaceAll = True
    
    Dim AL As clsArrayListString
    Set AL = New clsArrayListString
    
    'Empty
    On Error Resume Next
    AL.ReplaceAll "", "A"
    If Err.Number <> 0 Then
        TestReplaceAll = False
        Debug.Print "Empty"
    End If
    On Error GoTo 0
    If AL.Size <> 0 Then
        TestReplaceAll = False
        Debug.Print "Empty"
    End If

    'One
    AL.Append "A"
    AL.ReplaceAll "A", "B"
    If AL.JoinString <> "B" Or AL.Size <> 1 Then
        TestReplaceAll = False
        Debug.Print "One"
    End If
    AL.ReplaceAll "", "A"
    If AL.JoinString <> "B" Or AL.Size <> 1 Then
        TestReplaceAll = False
        Debug.Print "One"
    End If
    
    'Multiple
    AL.Reinitialize
    AL.Append "A"
    AL.Append "B"
    AL.Append "A"
    AL.ReplaceAll "A", "Z"
    If AL.JoinString <> "ZBZ" Or AL.Size <> 3 Then
       TestReplaceAll = False
       Debug.Print "Multiple"
    End If
    
    'First
    AL.Reinitialize
    AL.Append "A"
    AL.Append "B"
    AL.Append "C"
    AL.ReplaceAll "A", "Z"
    If AL.JoinString <> "ZBC" Or AL.Size <> 3 Then
       TestReplaceAll = False
       Debug.Print "First"
    End If

    'Middle
    AL.Reinitialize
    AL.Append "A"
    AL.Append "B"
    AL.Append "A"
    AL.Append "B"
    AL.ReplaceAll "A", "Z"
    If AL.JoinString <> "ZBZB" Or AL.Size <> 4 Then
       TestReplaceAll = False
       Debug.Print "Middle"
    End If

    'End
    AL.Reinitialize
    AL.Append "A"
    AL.Append "B"
    AL.Append "B"
    AL.Append "A"
    AL.ReplaceAll "A", "Z"
    If AL.JoinString <> "ZBBZ" Or AL.Size <> 4 Then
       TestReplaceAll = False
       Debug.Print "End"
    End If

    'Not there
    AL.Reinitialize
    AL.Append "A"
    AL.Append "B"
    AL.Append "C"
    AL.ReplaceAll "Z", ""
    If AL.JoinString <> "ABC" Or AL.Size <> 3 Then
       TestReplaceAll = False
       Debug.Print "Not there"
    End If

    'CompareMethod
    AL.Reinitialize
    AL.Append "A"
    AL.Append "B"
    AL.Append "C"
    AL.ReplaceAll "a", "Z"
    If AL.JoinString <> "ABC" Or AL.Size <> 3 Then
       TestReplaceAll = False
       Debug.Print "Binary"
    End If
    AL.ReplaceAll "a", "Z", vbTextCompare
    If AL.JoinString <> "ZBC" Or AL.Size <> 3 Then
       TestReplaceAll = False
       Debug.Print "Text"
    End If
    
    Debug.Print "TestReplaceAll: " & TestReplaceAll

End Function

Private Function TestIndexOf() As Boolean

    TestIndexOf = True
    
    Dim AL As clsArrayListString
    Set AL = New clsArrayListString
    
    'Empty
    If AL.IndexOf("") <> -1 Then
        TestIndexOf = False
        Debug.Print "Empty"
    End If

    'One
    AL.Append "A"
    If AL.IndexOf("A") <> 0 Then
        TestIndexOf = False
        Debug.Print "One"
    End If
    If AL.IndexOf("") <> -1 Then
        TestIndexOf = False
        Debug.Print "One"
    End If
    
    'Multiple
    AL.Reinitialize
    AL.Append "A"
    AL.Append "B"
    AL.Append "A"
    If AL.IndexOf("A") <> 0 Then
       TestIndexOf = False
       Debug.Print "Multiple"
    End If
    
    'First
    AL.Reinitialize
    AL.Append "A"
    AL.Append "B"
    AL.Append "C"
    If AL.IndexOf("A") <> 0 Then
       TestIndexOf = False
       Debug.Print "First"
    End If

    'Middle
    AL.Reinitialize
    AL.Append "B"
    AL.Append "A"
    AL.Append "B"
    AL.Append "A"
    If AL.IndexOf("A") <> 1 Then
       TestIndexOf = False
       Debug.Print "Middle"
    End If

    'End
    AL.Reinitialize
    AL.Append "B"
    AL.Append "B"
    AL.Append "A"
    If AL.IndexOf("A") <> 2 Then
       TestIndexOf = False
       Debug.Print "End"
    End If

    'Not there
    AL.Reinitialize
    AL.Append "A"
    AL.Append "B"
    AL.Append "C"
    If AL.IndexOf("Z") <> -1 Then
       TestIndexOf = False
       Debug.Print "Not there"
    End If

    'CompareMethod
    AL.Reinitialize
    AL.Append "A"
    AL.Append "B"
    AL.Append "C"
    If AL.IndexOf("a") <> -1 Then
       TestIndexOf = False
       Debug.Print "Binary"
    End If
    If AL.IndexOf("a", 0, vbTextCompare) <> 0 Then
       TestIndexOf = False
       Debug.Print "Text"
    End If
    
    'From
    AL.Reinitialize
    AL.Append "A"
    AL.Append "A"
    AL.Append "B"
    AL.Append "C"
    AL.Append "A"
    If AL.IndexOf("A", 0) <> 0 Then
       TestIndexOf = False
       Debug.Print "From"
    End If
    If AL.IndexOf("A", 1) <> 1 Then
       TestIndexOf = False
       Debug.Print "From"
    End If
    If AL.IndexOf("A", 2) <> 4 Then
       TestIndexOf = False
       Debug.Print "From"
    End If
    If AL.IndexOf("A", 3) <> 4 Then
       TestIndexOf = False
       Debug.Print "From"
    End If
    If AL.IndexOf("A", 4) <> 4 Then
       TestIndexOf = False
       Debug.Print "From"
    End If
    On Error Resume Next
    AL.IndexOf "A", 5
    If Err.Number <> 9 Then
       TestIndexOf = False
       Debug.Print "From"
    End If
    On Error GoTo 0
    
    'From negative
    AL.Reinitialize
    AL.Append "A"
    AL.Append "A"
    AL.Append "B"
    AL.Append "C"
    AL.Append "A"
    If AL.IndexOf("A", -1) <> 4 Then
       TestIndexOf = False
       Debug.Print "From negative"
    End If
    If AL.IndexOf("A", -2) <> 4 Then
       TestIndexOf = False
       Debug.Print "From negative"
    End If
    If AL.IndexOf("A", -3) <> 4 Then
       TestIndexOf = False
       Debug.Print "From negative"
    End If
    If AL.IndexOf("A", -4) <> 1 Then
       TestIndexOf = False
       Debug.Print "From negative"
    End If
    If AL.IndexOf("A", -5) <> 0 Then
       TestIndexOf = False
       Debug.Print "From negative"
    End If
    On Error Resume Next
    AL.IndexOf "A", -6
    If Err.Number <> 9 Then
       TestIndexOf = False
       Debug.Print "From"
    End If
    On Error GoTo 0
    
    Debug.Print "TestIndexOf: " & TestIndexOf

End Function

Private Function TestLastIndexOf() As Boolean

    TestLastIndexOf = True

    Dim AL As clsArrayListString
    Set AL = New clsArrayListString

    'Empty
    If AL.LastIndexOf("") <> -1 Then
        TestLastIndexOf = False
        Debug.Print "Empty"
    End If

    'One
    AL.Append "A"
    If AL.LastIndexOf("A") <> 0 Then
        TestLastIndexOf = False
        Debug.Print "One"
    End If
    If AL.LastIndexOf("") <> -1 Then
        TestLastIndexOf = False
        Debug.Print "One"
    End If

    'Multiple
    AL.Reinitialize
    AL.Append "A"
    AL.Append "B"
    AL.Append "A"
    If AL.LastIndexOf("A") <> 2 Then
       TestLastIndexOf = False
       Debug.Print "Multiple"
    End If

    'First
    AL.Reinitialize
    AL.Append "A"
    AL.Append "B"
    AL.Append "C"
    If AL.LastIndexOf("A") <> 0 Then
       TestLastIndexOf = False
       Debug.Print "First"
    End If

    'Middle
    AL.Reinitialize
    AL.Append "A"
    AL.Append "B"
    AL.Append "A"
    AL.Append "B"
    If AL.LastIndexOf("A") <> 2 Then
       TestLastIndexOf = False
       Debug.Print "Middle"
    End If

    'End
    AL.Reinitialize
    AL.Append "B"
    AL.Append "B"
    AL.Append "A"
    If AL.LastIndexOf("A") <> 2 Then
       TestLastIndexOf = False
       Debug.Print "End"
    End If

    'Not there
    AL.Reinitialize
    AL.Append "A"
    AL.Append "B"
    AL.Append "C"
    If AL.LastIndexOf("Z") <> -1 Then
       TestLastIndexOf = False
       Debug.Print "Not there"
    End If

    'CompareMethod
    AL.Reinitialize
    AL.Append "A"
    AL.Append "B"
    AL.Append "C"
    If AL.LastIndexOf("a") <> -1 Then
       TestLastIndexOf = False
       Debug.Print "Binary"
    End If
    If AL.LastIndexOf("a", 0, vbTextCompare) <> 0 Then
       TestLastIndexOf = False
       Debug.Print "Text"
    End If

    'From
    AL.Reinitialize
    AL.Append "A"
    AL.Append "A"
    AL.Append "B"
    AL.Append "A"
    AL.Append "C"
    If AL.LastIndexOf("A", 4) <> 3 Then
       TestLastIndexOf = False
       Debug.Print "From"
    End If
    If AL.LastIndexOf("A", 3) <> 3 Then
       TestLastIndexOf = False
       Debug.Print "From"
    End If
    If AL.LastIndexOf("A", 2) <> 1 Then
       TestLastIndexOf = False
       Debug.Print "From"
    End If
    If AL.LastIndexOf("A", 1) <> 1 Then
       TestLastIndexOf = False
       Debug.Print "From"
    End If
    If AL.LastIndexOf("A", 0) <> 0 Then
       TestLastIndexOf = False
       Debug.Print "From"
    End If
    On Error Resume Next
    AL.LastIndexOf "A", 5
    If Err.Number <> 9 Then
       TestLastIndexOf = False
       Debug.Print "From"
    End If
    On Error GoTo 0

    'From negative
    AL.Reinitialize
    AL.Append "A"
    AL.Append "A"
    AL.Append "B"
    AL.Append "C"
    AL.Append "A"
    If AL.LastIndexOf("A", -1) <> 4 Then
       TestLastIndexOf = False
       Debug.Print "From negative"
    End If
    If AL.LastIndexOf("A", -2) <> 1 Then
       TestLastIndexOf = False
       Debug.Print "From negative"
    End If
    If AL.LastIndexOf("A", -3) <> 1 Then
       TestLastIndexOf = False
       Debug.Print "From negative"
    End If
    If AL.LastIndexOf("A", -4) <> 1 Then
       TestLastIndexOf = False
       Debug.Print "From negative"
    End If
    If AL.LastIndexOf("A", -5) <> 0 Then
       TestLastIndexOf = False
       Debug.Print "From negative"
    End If
    On Error Resume Next
    AL.LastIndexOf "A", -6
    If Err.Number <> 9 Then
       TestLastIndexOf = False
       Debug.Print "From"
    End If
    On Error GoTo 0
    
    Debug.Print "TestLastIndexOf: " & TestLastIndexOf

End Function

Private Function TestCount() As Boolean

    TestCount = True
    
    Dim AL As clsArrayListString
    Set AL = New clsArrayListString
    
    'Empty
    If AL.Count("A") <> 0 Then
        TestCount = False
        Debug.Print "Empty"
    End If
    If AL.Count("") <> 0 Then
        TestCount = False
        Debug.Print "Empty"
    End If
    
    'One
    AL.Append "A"
    If AL.Count("A") <> 1 Then
        TestCount = False
        Debug.Print "One"
    End If
    If AL.Count("") <> 0 Then
        TestCount = False
        Debug.Print "One"
    End If
    
    'Multiple
    AL.Append "A"
    If AL.Count("A") <> 2 Then
        TestCount = False
        Debug.Print "Multiple"
    End If
    If AL.Count("") <> 0 Then
        TestCount = False
        Debug.Print "Multiple"
    End If
    
    'First
    AL.Reinitialize
    AL.Append "A"
    AL.Append "B"
    AL.Append "C"
    If AL.Count("A") <> 1 Then
        TestCount = False
        Debug.Print "First"
    End If
    
    'Middle
    AL.Reinitialize
    AL.Append "B"
    AL.Append "A"
    AL.Append "C"
    If AL.Count("A") <> 1 Then
        TestCount = False
        Debug.Print "Middle"
    End If
    
    'End
    AL.Reinitialize
    AL.Append "C"
    AL.Append "B"
    AL.Append "A"
    If AL.Count("A") <> 1 Then
        TestCount = False
        Debug.Print "End"
    End If
    
    'Not there
    If AL.Count("Z") <> 0 Then
        TestCount = False
        Debug.Print "Not there"
    End If
    
    'CompareMethod
    AL.Reinitialize
    AL.Append "A"
    AL.Append "B"
    AL.Append "C"
    If AL.Count("a") <> 0 Then
        TestCount = False
        Debug.Print "CompareMethod Binary"
    End If
    If AL.Count("a", vbTextCompare) <> 1 Then
        TestCount = False
        Debug.Print "CompareMethod TExt"
    End If
    
    Debug.Print "TestCount: " & TestCount

End Function

Private Function TestContains() As Boolean

    TestContains = True

    Dim AL As clsArrayListString
    Set AL = New clsArrayListString
    
    'Empty
    If AL.Contains("A") <> False Then
        TestContains = False
        Debug.Print "Empty"
    End If
    If AL.Contains("") <> False Then
        TestContains = False
        Debug.Print "Empty"
    End If
    
    'One
    AL.Append "A"
    If AL.Contains("A") <> True Then
        TestContains = False
        Debug.Print "One"
    End If
    If AL.Contains("") <> False Then
        TestContains = False
        Debug.Print "One"
    End If
    
    'Multiple
    AL.Append "A"
    If AL.Contains("A") <> True Then
        TestContains = False
        Debug.Print "Multiple"
    End If
    If AL.Contains("") <> False Then
        TestContains = False
        Debug.Print "Multiple"
    End If
    
    'First
    AL.Reinitialize
    AL.Append "A"
    AL.Append "B"
    AL.Append "C"
    If AL.Contains("A") <> True Then
        TestContains = False
        Debug.Print "First"
    End If
    
    'Middle
    AL.Reinitialize
    AL.Append "B"
    AL.Append "A"
    AL.Append "C"
    If AL.Contains("A") <> True Then
        TestContains = False
        Debug.Print "Middle"
    End If
    
    'End
    AL.Reinitialize
    AL.Append "C"
    AL.Append "B"
    AL.Append "A"
    If AL.Contains("A") <> True Then
        TestContains = False
        Debug.Print "End"
    End If
    
    'Not there
    If AL.Contains("Z") <> False Then
        TestContains = False
        Debug.Print "Not there"
    End If
    
    'CompareMethod
    AL.Reinitialize
    AL.Append "A"
    AL.Append "B"
    AL.Append "C"
    If AL.Contains("a") <> False Then
        TestContains = False
        Debug.Print "CompareMethod Binary"
    End If
    If AL.Contains("a", vbTextCompare) <> True Then
        TestContains = False
        Debug.Print "CompareMethod TExt"
    End If
    
    Debug.Print "TestContains: " & TestContains

End Function

Private Function TestClear() As Boolean

    TestClear = True
    
    Dim AL As clsArrayListString
    Set AL = New clsArrayListString
    
    'Empty
    AL.Clear
    If AL.JoinString <> "" Or AL.Size <> 0 Then
        TestClear = False
        Debug.Print "Empty"
    End If
    
    'One
    AL.Append "A"
    AL.Clear
    If AL.JoinString <> "" Or AL.Size <> 0 Then
        TestClear = False
        Debug.Print "One"
    End If

    'Multiple
    AL.Append "A"
    AL.Append "B"
    AL.Append "C"
    AL.Clear
    If AL.JoinString <> "" Or AL.Size <> 0 Then
        TestClear = False
        Debug.Print "Multiple"
    End If
    
    'Capacity persists
    Dim i&
    For i = 1 To 11
        AL.Append "A"
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
    
    Dim AL As clsArrayListString
    Set AL = New clsArrayListString
    
    'Empty
    On Error Resume Next
    AL.Reverse
    If Err.Number <> 0 Then
        TestReverse = False
        Debug.Print "Empty"
    End If
    On Error GoTo 0
    
    'One
    AL.Append "A"
    AL.Reverse
    If AL.JoinString <> "A" Then
        TestReverse = False
        Debug.Print "One"
    End If
    
    'Multiple
    AL.Append "B"
    AL.Append "C"
    AL.Reverse
    If AL.JoinString <> "CBA" Then
        TestReverse = False
        Debug.Print "Multiple"
    End If
    
    'Even
    AL.Reinitialize
    AL.Append "A"
    AL.Append "B"
    AL.Append "C"
    AL.Append "D"
    AL.Reverse
    If AL.JoinString <> "DCBA" Then
        TestReverse = False
        Debug.Print "Even"
    End If
    
    'Odd
    AL.Reinitialize
    AL.Append "A"
    AL.Append "B"
    AL.Append "C"
    AL.Reverse
    If AL.JoinString <> "CBA" Then
        TestReverse = False
        Debug.Print "Odd"
    End If
    
    Debug.Print "TestReverse: " & TestReverse

End Function

Private Function TestToStringArray() As Boolean

    TestToStringArray = True
    
    Dim AL As clsArrayListString
    Set AL = New clsArrayListString
    
    'Empty
    Dim Arr$()
    Arr = AL.ToStringArray
    On Error Resume Next
    Debug.Print LBound(Arr)
    If Err.Number <> 9 Then
        TestToStringArray = False
        Debug.Print "Empty"
    End If
    On Error GoTo 0
    
    'One
    AL.Append "A"
    Arr = AL.ToStringArray
    If UBound(Arr) - LBound(Arr) + 1 <> AL.Size Then
        TestToStringArray = False
        Debug.Print "One"
    End If
    If Join(Arr, "") <> AL.JoinString Then
        TestToStringArray = False
        Debug.Print "One"
    End If
    
    'Multiple
    AL.Append "B"
    AL.Append "C"
    Arr = AL.ToStringArray
    If UBound(Arr) - LBound(Arr) + 1 <> AL.Size Then
        TestToStringArray = False
        Debug.Print "Multiple"
    End If
    If Join(Arr, "") <> AL.JoinString Then
        TestToStringArray = False
        Debug.Print "Multiple"
    End If
    
    Debug.Print "TestToStringArray: " & TestToStringArray

End Function

Private Function TestJoinString() As Boolean

    TestJoinString = True
    
    Dim AL As clsArrayListString
    Set AL = New clsArrayListString
    
    'No delimiter
        'Empty
        If AL.JoinString <> "" Then
            TestJoinString = False
            Debug.Print "No delimiter Empty"
        End If
        'One
        AL.Append "A"
        If AL.JoinString <> "A" Then
            TestJoinString = False
            Debug.Print "No delimiter One"
        End If
        'Multiple
        AL.Append "B"
        AL.Append "C"
        If AL.JoinString <> "ABC" Then
            TestJoinString = False
            Debug.Print "No delimiter Multiple"
        End If
    
    'Delimiter
        AL.Reinitialize
        'Empty
        If AL.JoinString(",") <> "" Then
            TestJoinString = False
            Debug.Print "Delimiter Empty"
        End If
        'One
        AL.Append "A"
        If AL.JoinString(",") <> "A" Then
            TestJoinString = False
            Debug.Print "Delimiter One"
        End If
        'Multiple
        AL.Append "B"
        AL.Append "C"
        If AL.JoinString(",") <> "A,B,C" Then
            TestJoinString = False
            Debug.Print "Delimiter Multiple"
        End If
        
    Debug.Print "TestJoinString: " & TestJoinString

End Function

Private Function TestSort() As Boolean

    TestSort = True

    Dim AL As clsArrayListString
    Set AL = New clsArrayListString

    AL.Reinitialize
    '(Insertion Sort)
        'Empty
        AL.Sort
        If AL.JoinString <> "" Then
            TestSort = False
            Debug.Print "Insertion Sort Empty"
        End If
        'One
        AL.Append "A"
        AL.Sort
        If AL.JoinString <> "A" Then
            TestSort = False
            Debug.Print "Insertion Sort One"
        End If
        'Many
        AL.Reinitialize
        AL.Append "A"
        AL.Append "M"
        AL.Append "F"
        AL.Append "I"
        AL.Append "J"
        AL.Append "U"
        AL.Append "K"
        AL.Append "W"
        AL.Append "R"
        AL.Append "Z"
        AL.Append "B"
        AL.Append "X"
        AL.Append "S"
        AL.Append "E"
        AL.Append "Y"
        AL.Append "G"
        AL.Append "H"
        AL.Append "N"
        AL.Append "O"
        AL.Append "P"
        AL.Append "C"
        AL.Append "D"
        AL.Append "Q"
        AL.Append "L"
        AL.Append "T"
        AL.Append "V"
        AL.Sort
        If AL.JoinString <> "ABCDEFGHIJKLMNOPQRSTUVWXYZ" Then
            TestSort = False
            Debug.Print "Insertion Sort Many"
        End If
        'CompareMethod Binary
        AL.Reinitialize
        AL.Append "l"
        AL.Append "M"
        AL.Append "N"
        AL.Append "B"
        AL.Append "f"
        AL.Append "G"
        AL.Append "F"
        AL.Append "H"
        AL.Append "D"
        AL.Append "e"
        AL.Append "O"
        AL.Append "c"
        AL.Append "d"
        AL.Append "p"
        AL.Append "J"
        AL.Append "K"
        AL.Append "A"
        AL.Append "n"
        AL.Append "E"
        AL.Append "P"
        AL.Append "a"
        AL.Append "L"
        AL.Append "b"
        AL.Append "C"
        AL.Append "m"
        AL.Append "o"
        AL.Append "g"
        AL.Append "h"
        AL.Append "i"
        AL.Append "j"
        AL.Append "k"
        AL.Append "I"
        AL.Sort
        If AL.JoinString <> "ABCDEFGHIJKLMNOPabcdefghijklmnop" Then
            TestSort = False
            Debug.Print "Insertion Sort CompareMethod Binary"
        End If
        'CompareMethod Text
        AL.Reinitialize
        AL.Append "e"
        AL.Append "g"
        AL.Append "f"
        AL.Append "K"
        AL.Append "A"
        AL.Append "F"
        AL.Append "h"
        AL.Append "i"
        AL.Append "n"
        AL.Append "o"
        AL.Append "p"
        AL.Append "I"
        AL.Append "L"
        AL.Append "b"
        AL.Append "c"
        AL.Append "d"
        AL.Append "l"
        AL.Append "m"
        AL.Append "O"
        AL.Append "P"
        AL.Append "M"
        AL.Append "C"
        AL.Append "D"
        AL.Append "E"
        AL.Append "N"
        AL.Append "G"
        AL.Append "H"
        AL.Append "j"
        AL.Append "k"
        AL.Append "B"
        AL.Append "a"
        AL.Append "J"
        AL.Sort vbTextCompare
        If StrComp(AL.JoinString, "AaBbCcDdEeFfGgHhIiJjKkLlMmNnOoPp", vbTextCompare) <> 0 Then
            TestSort = False
            Debug.Print "Insertion Sort CompareMethod Text"
        End If

    AL.Reinitialize
    '(Merge Sort)
        'Empty
        AL.Sort
        If AL.JoinString <> "" Then
            TestSort = False
            Debug.Print "Merge Sort Empty"
        End If
        'One
        AL.Append "A"
        AL.Sort
        If AL.JoinString <> "A" Then
            TestSort = False
            Debug.Print "Merge Sort One"
        End If
        'Many
        AL.Reinitialize
        AL.Append "A"
        AL.Append "M"
        AL.Append "F"
        AL.Append "I"
        AL.Append "J"
        AL.Append "U"
        AL.Append "K"
        AL.Append "W"
        AL.Append "R"
        AL.Append "Z"
        AL.Append "B"
        AL.Append "X"
        AL.Append "S"
        AL.Append "E"
        AL.Append "Y"
        AL.Append "G"
        AL.Append "H"
        AL.Append "N"
        AL.Append "O"
        AL.Append "P"
        AL.Append "C"
        AL.Append "D"
        AL.Append "Q"
        AL.Append "L"
        AL.Append "T"
        AL.Append "V"
        AL.Append "A"
        AL.Append "M"
        AL.Append "F"
        AL.Append "I"
        AL.Append "J"
        AL.Append "U"
        AL.Append "K"
        AL.Append "W"
        AL.Append "R"
        AL.Append "Z"
        AL.Append "B"
        AL.Append "X"
        AL.Append "S"
        AL.Append "E"
        AL.Append "Y"
        AL.Append "G"
        AL.Append "H"
        AL.Append "N"
        AL.Append "O"
        AL.Append "P"
        AL.Append "C"
        AL.Append "D"
        AL.Append "Q"
        AL.Append "L"
        AL.Append "T"
        AL.Append "V"
        AL.Sort
        If AL.JoinString <> "AABBCCDDEEFFGGHHIIJJKKLLMMNNOOPPQQRRSSTTUUVVWWXXYYZZ" Then
            TestSort = False
            Debug.Print "Merge Sort Many"
        End If
        'CompareMethod Binary
        AL.Reinitialize
        AL.Append "w"
        AL.Append "n"
        AL.Append "z"
        AL.Append "Z"
        AL.Append "a"
        AL.Append "y"
        AL.Append "N"
        AL.Append "J"
        AL.Append "h"
        AL.Append "i"
        AL.Append "j"
        AL.Append "b"
        AL.Append "c"
        AL.Append "d"
        AL.Append "x"
        AL.Append "X"
        AL.Append "V"
        AL.Append "s"
        AL.Append "o"
        AL.Append "A"
        AL.Append "m"
        AL.Append "t"
        AL.Append "u"
        AL.Append "E"
        AL.Append "F"
        AL.Append "G"
        AL.Append "L"
        AL.Append "k"
        AL.Append "K"
        AL.Append "v"
        AL.Append "Y"
        AL.Append "g"
        AL.Append "W"
        AL.Append "H"
        AL.Append "I"
        AL.Append "B"
        AL.Append "O"
        AL.Append "P"
        AL.Append "l"
        AL.Append "p"
        AL.Append "q"
        AL.Append "r"
        AL.Append "U"
        AL.Append "C"
        AL.Append "D"
        AL.Append "e"
        AL.Append "f"
        AL.Append "M"
        AL.Append "Q"
        AL.Append "R"
        AL.Append "S"
        AL.Append "T"
        AL.Sort
        If AL.JoinString <> "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz" Then
            TestSort = False
            Debug.Print "MergeSort CompareMethod Binary"
        End If
        'CompareMethod Text
        AL.Reinitialize
        AL.Append "w"
        AL.Append "n"
        AL.Append "z"
        AL.Append "Z"
        AL.Append "a"
        AL.Append "y"
        AL.Append "N"
        AL.Append "J"
        AL.Append "h"
        AL.Append "i"
        AL.Append "j"
        AL.Append "b"
        AL.Append "c"
        AL.Append "d"
        AL.Append "x"
        AL.Append "X"
        AL.Append "V"
        AL.Append "s"
        AL.Append "o"
        AL.Append "A"
        AL.Append "m"
        AL.Append "t"
        AL.Append "u"
        AL.Append "E"
        AL.Append "F"
        AL.Append "G"
        AL.Append "L"
        AL.Append "k"
        AL.Append "K"
        AL.Append "v"
        AL.Append "Y"
        AL.Append "g"
        AL.Append "W"
        AL.Append "H"
        AL.Append "I"
        AL.Append "B"
        AL.Append "O"
        AL.Append "P"
        AL.Append "l"
        AL.Append "p"
        AL.Append "q"
        AL.Append "r"
        AL.Append "U"
        AL.Append "C"
        AL.Append "D"
        AL.Append "e"
        AL.Append "f"
        AL.Append "M"
        AL.Append "Q"
        AL.Append "R"
        AL.Append "S"
        AL.Append "T"
        AL.Sort vbTextCompare
        If StrComp(AL.JoinString, "AaBbCcDdEeFfGgHhIiJjKkLlMmNnOoPpQqRrSsTtUuVvWwXxYyZz", vbTextCompare) <> 0 Then
            TestSort = False
            Debug.Print "MergeSort CompareMethod Text"
        End If
 
    Debug.Print "TestSort: " & TestSort

End Function
