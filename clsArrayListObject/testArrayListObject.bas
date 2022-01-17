Attribute VB_Name = "testArrayListObject"
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
'  Module Name: testArrayListObject
'  Module Description: Unit tests for clsArrayListObject class.
'  Module Version: 1.0
'  Module License: MIT
'  Module Author: Peter Roach; PeterRoach.Code@gmail.com
'  Module Contents:
'        TestArrayListObject
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
'        TestToArray
'        TestJoinString
'        TestCallMethod
'        TestSort

Private Const DEFAULT_CAPACITY& = 10

Private Function EG(Msg$) As clsExample
    Set EG = New clsExample
    EG.Message = Msg
End Function


'Example Usage=========================================================
'======================================================================

Public Sub Example()
    
    Dim C1 As Collection
    Set C1 = New Collection
    C1.Add 1
    
    Dim C3 As Collection
    Set C3 = New Collection
    C3.Add 1
    C3.Add 2
    C3.Add 3
    
    Dim C2 As Collection
    Set C2 = New Collection
    C2.Add 1
    C2.Add 2
    
    Dim AL As clsArrayListObject
    Set AL = New clsArrayListObject
    AL.Append C1
    AL.Append C2
    AL.Append C3
    
    AL.Sort "Count", True
    
    Debug.Print AL.JoinString("Count", ", ", True)
    
End Sub


'Unit Tests============================================================
'======================================================================

Public Function TestArrayListObject() As Boolean

    TestArrayListObject = _
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
    TestArrayListObject = TestArrayListObject And _
        TestClear And _
        TestReverse And _
        TestToArray And _
        TestJoinString And _
        TestCallMethod And _
        TestSort

    Debug.Print "TestArrayListObject: " & TestArrayListObject

End Function

Private Function TestCapacity() As Boolean

    TestCapacity = True

    Dim AL As clsArrayListObject
    Set AL = New clsArrayListObject
    
    Dim AL1 As clsArrayListObject
    Set AL1 = New clsArrayListObject
    
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
    AL.Append New clsExample
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
        AL.Append New clsExample
    Next i
    If AL.Capacity <> DEFAULT_CAPACITY * 2 + 2 Then
        TestCapacity = False
        Debug.Print "Append"
    End If
    
    'AppendArray
    AL.Reinitialize
    Dim Arr(0 To 10) As clsExample
    AL.AppendArray Arr
    If AL.Capacity <> DEFAULT_CAPACITY * 2 + 2 Then
        TestCapacity = False
        Debug.Print "AppendArray"
    End If
    
    'AppendArrayList
    AL.Reinitialize
    AL1.AppendArray Arr
    AL.AppendArrayList AL1
    If AL.Capacity <> DEFAULT_CAPACITY * 2 + 2 Then
        TestCapacity = False
        Debug.Print "AppendArrayList"
    End If
    
    'Insert
    AL.Reinitialize
    For i = 1 To 11
        AL.Insert 0, New clsExample
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
    AL.Append New clsExample
    AL.Remove 0
    If AL.Capacity <> DEFAULT_CAPACITY Then
        TestCapacity = False
        Debug.Print "Remove"
    End If
    
    'RemoveRange
    AL.Reinitialize
    AL.Append New clsExample
    AL.Append New clsExample
    AL.Append New clsExample
    AL.RemoveRange 0, 1
    If AL.Capacity <> DEFAULT_CAPACITY Then
        TestCapacity = False
        Debug.Print "RemoveRange"
    End If
    
    'RemoveFirst
    AL.Reinitialize
    AL.Append EG("A")
    AL.Append EG("B")
    AL.Append EG("C")
    AL.RemoveFirst "C", "Message"
    If AL.Capacity <> DEFAULT_CAPACITY Then
        TestCapacity = False
        Debug.Print "RemoveFirst"
    End If
    
    'RemoveLast
    AL.Reinitialize
    AL.Append EG("A")
    AL.Append EG("B")
    AL.Append EG("C")
    AL.RemoveLast "A", "Message"
    If AL.Capacity <> DEFAULT_CAPACITY Then
        TestCapacity = False
        Debug.Print "RemoveLast"
    End If
    
    'RemoveAll
    AL.Reinitialize
    AL.Append EG("A")
    AL.Append EG("A")
    AL.Append EG("A")
    AL.RemoveLast "A", "Message"
    If AL.Capacity <> DEFAULT_CAPACITY Then
        TestCapacity = False
        Debug.Print "RemoveAll"
    End If
    
    'ReplaceAll
    AL.Reinitialize
    AL.Append EG("A")
    AL.Append EG("A")
    AL.Append EG("A")
    AL.ReplaceAll "A", CStr("B"), "Message"
    If AL.Capacity <> DEFAULT_CAPACITY Then
        TestCapacity = False
        Debug.Print "ReplaceAll"
    End If
    
    'IndexOf
    AL.Reinitialize
    AL.Append EG("A")
    AL.Append EG("A")
    AL.Append EG("A")
    If AL.IndexOf("A", "Message") <> 0 Or AL.Capacity <> DEFAULT_CAPACITY Then
        TestCapacity = False
        Debug.Print "IndexOf"
    End If
    
    'LastIndexOf
    AL.Reinitialize
    AL.Append EG("A")
    AL.Append EG("A")
    AL.Append EG("A")
    If AL.LastIndexOf("A", "Message") <> 2 Or AL.Capacity <> DEFAULT_CAPACITY Then
        TestCapacity = False
        Debug.Print "LastIndexOf"
    End If
    
    'Count
    AL.Reinitialize
    AL.Append EG("A")
    AL.Append EG("A")
    AL.Append EG("B")
    If AL.Count("A", "Message") <> 2 Or AL.Capacity <> DEFAULT_CAPACITY Then
        TestCapacity = False
        Debug.Print "Count"
    End If
    
    'Contains
    AL.Reinitialize
    AL.Append EG("A")
    AL.Append EG("B")
    AL.Append EG("C")
    If AL.Contains("A", "Message") <> True Or AL.Capacity <> DEFAULT_CAPACITY Then
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
    AL.Append EG("A")
    AL.Append EG("B")
    AL.Append EG("C")
    AL.Reverse
    If AL.JoinString("Message") <> "CBA" Or AL.Capacity <> DEFAULT_CAPACITY Then
        TestCapacity = False
        Debug.Print "Reverse"
    End If
    
    'ToArray
    AL.Reinitialize
    AL.Append EG("A")
    AL.Append EG("B")
    AL.Append EG("C")
    Dim ToArr() As Object
    ToArr = AL.ToArray
    Dim ToArrString$
    ToArrString = ToArr(0).Message & ToArr(1).Message & ToArr(2).Message
    If AL.JoinString("Message") <> ToArrString Or AL.Capacity <> DEFAULT_CAPACITY Then
        TestCapacity = False
        Debug.Print "ToArray"
    End If
    
    'JoinString
    AL.Reinitialize
    AL.Append New clsExample
    AL.Append New clsExample
    AL.Append New clsExample
    If AL.JoinString("Message") <> "" Or AL.Capacity <> DEFAULT_CAPACITY Then
        TestCapacity = False
        Debug.Print "JoinString"
    End If
    
    'Sort
    AL.Reinitialize
    AL.Append New clsExample
    AL.Append New clsExample
    AL.Append New clsExample
    AL.Sort "Message"
    If AL.JoinString("Message") <> "" Or AL.Capacity <> DEFAULT_CAPACITY Then
        TestCapacity = False
        Debug.Print "Sort"
    End If

    Debug.Print "TestCapacity: " & TestCapacity

End Function

Private Function TestSize() As Boolean

    TestSize = True
    
    Dim AL As clsArrayListObject
    Set AL = New clsArrayListObject
    
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
    AL.Append EG("A")
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
    AL.Append EG("A")
    If AL.Size <> 1 Then
        TestSize = False
        Debug.Print "Append"
    End If
    
    'AppendArray
    Dim Arr(0 To 2) As Object
    AL.AppendArray Arr
    If AL.Size <> 4 Then
        TestSize = False
        Debug.Print "AppendArray"
    End If
    
    'AppendArrayList
    Dim AL1 As clsArrayListObject
    Set AL1 = New clsArrayListObject
    AL1.AppendArray Arr
    AL.AppendArrayList AL1
    If AL.Size <> 7 Then
        TestSize = False
        Debug.Print "AppendArrayList"
    End If
    
    'Insert
    AL.Insert 0, EG("A")
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
    AL.Append EG("A")
    AL.Append EG("B")
    AL.Append EG("A")
    AL.RemoveFirst "A", "Message"
    If AL.Size <> 2 Then
        TestSize = False
        Debug.Print "RemoveFirst"
    End If
    
    'RemoveLast
    AL.Reinitialize
    AL.Append EG("A")
    AL.Append EG("B")
    AL.Append EG("A")
    AL.RemoveLast "A", "Message"
    If AL.Size <> 2 Then
        TestSize = False
        Debug.Print "RemoveLast"
    End If
    
    'RemoveAll
    AL.Reinitialize
    AL.Append EG("A")
    AL.Append EG("B")
    AL.Append EG("A")
    AL.RemoveAll "A", "Message"
    If AL.Size <> 1 Then
        TestSize = False
        Debug.Print "RemoveAll"
    End If
    
    'ReplaceAll
    AL.Reinitialize
    AL.Append EG("A")
    AL.Append EG("B")
    AL.Append EG("A")
    AL.ReplaceAll "A", "B", "Message"
    If AL.Size <> 3 Then
        TestSize = False
        Debug.Print "ReplaceAll"
    End If
    
    'IndexOf
    AL.Reinitialize
    AL.Append EG("A")
    AL.Append EG("B")
    AL.Append EG("A")
    If AL.IndexOf("A", "Message") <> 0 Or AL.Size <> 3 Then
        TestSize = False
        Debug.Print "IndexOf"
    End If
    
    'LastIndexOf
    AL.Reinitialize
    AL.Append EG("A")
    AL.Append EG("B")
    AL.Append EG("A")
    If AL.LastIndexOf("A", "Message") <> 2 Or AL.Size <> 3 Then
        TestSize = False
        Debug.Print "LastIndexOf"
    End If
    
    'Count
    AL.Reinitialize
    AL.Append EG("A")
    AL.Append EG("B")
    AL.Append EG("A")
    If AL.Count("A", "Message") <> 2 Or AL.Size <> 3 Then
        TestSize = False
        Debug.Print "Count"
    End If
    
    'Contains
    AL.Reinitialize
    AL.Append EG("A")
    AL.Append EG("B")
    AL.Append EG("A")
    If AL.Contains("A", "Message") <> True Or AL.Size <> 3 Then
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
    AL.Append EG("A")
    AL.Append EG("B")
    AL.Append EG("C")
    AL.Reverse
    If AL.JoinString("Message") <> "CBA" Or AL.Size <> 3 Then
        TestSize = False
        Debug.Print "Reverse"
    End If
    
    'ToArray
    AL.Reinitialize
    AL.Append EG("A")
    AL.Append EG("B")
    AL.Append EG("C")
    Dim ToArr() As Object
    ToArr = AL.ToArray
    Dim ToArrString$
    ToArrString = ToArr(0).Message & ToArr(1).Message & ToArr(2).Message
    If AL.JoinString("Message") <> ToArrString Or AL.Size <> 3 Then
        TestSize = False
        Debug.Print "ToStringArray"
    End If
    
    'JoinString
    AL.Reinitialize
    AL.Append EG("A")
    AL.Append EG("B")
    AL.Append EG("C")
    If AL.JoinString("Message") <> "ABC" Or AL.Size <> 3 Then
        TestSize = False
        Debug.Print "JoinString"
    End If
    
    'Sort
    AL.Reinitialize
    AL.Append EG("C")
    AL.Append EG("B")
    AL.Append EG("A")
    AL.Sort "Message"
    If AL.JoinString("Message") <> "ABC" Or AL.Size <> 3 Then
        TestSize = False
        Debug.Print "Sort"
    End If
    
    Debug.Print "TestSize: " & TestSize

End Function

Private Function TestItem() As Boolean

    TestItem = True

    Dim AL As clsArrayListObject
    Set AL = New clsArrayListObject

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
        Set AL.Item(0) = EG("A")
        If Err.NUMBER <> 9 Then
            TestItem = False
            Debug.Print "Empty Let"
        End If
        On Error GoTo 0

    'Non-Empty
    AL.Append EG("A")
        'Get
        If AL.Item(0).Message <> "A" Then
            TestItem = False
            Debug.Print "Non-Empty Get"
        End If
        'Let
        Set AL.Item(0) = EG("B")
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
    
    Dim AL As clsArrayListObject
    Set AL = New clsArrayListObject
    
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
    
    Dim AL As clsArrayListObject
    Set AL = New clsArrayListObject
    
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
    AL.Append EG("A")
    AL.Append EG("B")
    AL.Append EG("C")
    AL.ShrinkCapacity 3
    If AL.Size <> 3 Then
        TestShrinkCapacity = False
        Debug.Print "Size"
    End If
    
    'Less than size
    AL.Reinitialize
    AL.Append EG("A")
    AL.Append EG("B")
    AL.Append EG("C")
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
    
    Dim AL As clsArrayListObject
    Set AL = New clsArrayListObject
    
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

    Dim AL As clsArrayListObject
    Set AL = New clsArrayListObject

    'Size 0
    AL.TrimToSize
    If AL.Capacity <> DEFAULT_CAPACITY Or AL.Size <> 0 Then
        TestTrimToSize = False
        Debug.Print "Size 0"
    End If

    'Size 1
    AL.Append EG("A")
    AL.TrimToSize
    If AL.Capacity <> 1 Or AL.Size <> 1 Then
        TestTrimToSize = False
        Debug.Print "Size 1"
    End If

    'Size > 1
    AL.Reinitialize
    AL.Append EG("A")
    AL.Append EG("B")
    AL.Append EG("C")
    AL.TrimToSize
    If AL.Capacity <> 3 Or AL.Size <> 3 Then
        TestTrimToSize = False
        Debug.Print "Size > 1"
    End If

    'Reinitialize
    AL.Reinitialize
    AL.GrowCapacity DEFAULT_CAPACITY * 2 + 2
    AL.Append EG("A")
    AL.Append EG("B")
    AL.Append EG("C")
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
    
    Dim AL As clsArrayListObject
    Set AL = New clsArrayListObject
    
    'Default
    AL.GrowCapacity DEFAULT_CAPACITY * 2 + 2
    AL.Append EG("A")
    AL.Append EG("B")
    AL.Append EG("C")
    AL.Reinitialize
    If AL.Capacity <> DEFAULT_CAPACITY Or AL.Size <> 0 Then
        TestReinitialize = False
        Debug.Print "Default"
    End If
    
    'Less than default
    AL.Reinitialize
    AL.GrowCapacity DEFAULT_CAPACITY * 2 + 2
    AL.Append EG("A")
    AL.Append EG("B")
    AL.Append EG("C")
    AL.Reinitialize DEFAULT_CAPACITY - 1
    If AL.Capacity <> DEFAULT_CAPACITY - 1 Or AL.Size <> 0 Then
        TestReinitialize = False
        Debug.Print "Less than default"
    End If
    
    'More than default
    AL.Reinitialize
    AL.GrowCapacity DEFAULT_CAPACITY * 2 + 2
    AL.Append EG("A")
    AL.Append EG("B")
    AL.Append EG("C")
    AL.Reinitialize DEFAULT_CAPACITY + 1
    If AL.Capacity <> DEFAULT_CAPACITY + 1 Or AL.Size <> 0 Then
        TestReinitialize = False
        Debug.Print "More than default"
    End If
    
    'Invalid 0
    AL.Reinitialize
    AL.GrowCapacity DEFAULT_CAPACITY * 2 + 2
    AL.Append EG("A")
    AL.Append EG("B")
    AL.Append EG("C")
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
    AL.Append EG("A")
    AL.Append EG("B")
    AL.Append EG("C")
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

    Dim AL As clsArrayListObject
    Set AL = New clsArrayListObject

    'Empty
    AL.Append EG("A")
    If AL.JoinString("Message") <> "A" Or AL.Size <> 1 Then
        TestAppend = False
        Debug.Print "Empty"
    End If

    'Non empty
    AL.Append EG("B")
    If AL.JoinString("Message") <> "AB" Or AL.Size <> 2 Then
        TestAppend = False
        Debug.Print "Non Empty"
    End If

    'Until capacity
    Dim i&
    AL.Reinitialize
    For i = 1 To 11
        AL.Append EG("A")
    Next i
    If AL.JoinString("Message") <> String$(11, "A") Or _
    AL.Capacity <> DEFAULT_CAPACITY * 2 + 2 Or _
    AL.Size <> 11 Then
        TestAppend = False
        Debug.Print "Until capacity"
    End If

    Debug.Print "TestAppend: " & TestAppend

End Function

Private Function TestAppendArray() As Boolean

    TestAppendArray = True

    Dim AL As clsArrayListObject
    Set AL = New clsArrayListObject

    Dim Arr(0 To 2) As Object

    'Empty
    Set Arr(0) = EG("A")
    Set Arr(1) = EG("B")
    Set Arr(2) = EG("C")
    AL.AppendArray Arr
    If AL.JoinString("Message") <> "ABC" Or AL.Size <> 3 Then
        TestAppendArray = False
        Debug.Print "Empty"
    End If
    
    'Non empty
    Set Arr(0) = EG("X")
    Set Arr(1) = EG("Y")
    Set Arr(2) = EG("Z")
    AL.AppendArray Arr
    If AL.JoinString("Message") <> "ABCXYZ" Or AL.Size <> 6 Then
        TestAppendArray = False
        Debug.Print "Non empty"
    End If
    
    'Until capacity
    Set Arr(0) = EG("L")
    Set Arr(1) = EG("M")
    Set Arr(2) = EG("N")
    AL.AppendArray Arr
    Set Arr(0) = EG("T")
    Set Arr(1) = EG("U")
    Set Arr(2) = EG("V")
    AL.AppendArray Arr
    If AL.JoinString("Message") <> "ABCXYZLMNTUV" Or _
    AL.Size <> 12 Or _
    AL.Capacity <> DEFAULT_CAPACITY * 2 + 2 Then
        TestAppendArray = False
        Debug.Print "Until capacity"
    End If
    
    'Empty Array
    Dim Arr1() As Object
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
    
    Dim AL As clsArrayListObject
    Set AL = New clsArrayListObject
    
    Dim AL1 As clsArrayListObject
    Set AL1 = New clsArrayListObject

    'Empty
    AL1.Append EG("A")
    AL1.Append EG("B")
    AL1.Append EG("C")
    AL.AppendArrayList AL1
    If AL.JoinString("Message") <> "ABC" Or AL.Size <> 3 Then
        TestAppendArrayList = False
        Debug.Print "Empty"
    End If
    
    'Non empty
    AL1.Reinitialize
    AL1.Append EG("X")
    AL1.Append EG("Y")
    AL1.Append EG("Z")
    AL.AppendArrayList AL1
    If AL.JoinString("Message") <> "ABCXYZ" Or AL.Size <> 6 Then
        TestAppendArrayList = False
        Debug.Print "Non empty"
    End If
    
    'Until capacity
    AL1.Reinitialize
    AL1.Append EG("L")
    AL1.Append EG("M")
    AL1.Append EG("N")
    AL.AppendArrayList AL1
    AL1.Reinitialize
    AL1.Append EG("T")
    AL1.Append EG("U")
    AL1.Append EG("V")
    AL.AppendArrayList AL1
    If AL.JoinString("Message") <> "ABCXYZLMNTUV" Or _
    AL.Size <> 12 Or _
    AL.Capacity <> DEFAULT_CAPACITY * 2 + 2 Then
        TestAppendArrayList = False
        Debug.Print "Until capacity"
    End If
    
    'Empty ArrayList
    AL.Reinitialize
    AL1.Reinitialize
    AL.AppendArrayList AL1
    If AL.JoinString("Message") <> "" Or _
    AL.Capacity <> DEFAULT_CAPACITY Or _
    AL.Size <> 0 Then
        TestAppendArrayList = False
        Debug.Print "Empty ArrayList"
    End If
    
    Debug.Print "TestAppendArrayList: " & TestAppendArrayList

End Function

Private Function TestInsert() As Boolean

    TestInsert = True
    
    Dim AL As clsArrayListObject
    Set AL = New clsArrayListObject
    
    'Empty
    AL.Insert 0, EG("A")
    If AL.JoinString("Message") <> "A" Or AL.Size <> 1 Then
        TestInsert = False
        Debug.Print "Empty"
    End If
    
    'Non empty
        'Start
        AL.Insert 0, EG("B")
        If AL.JoinString("Message") <> "BA" Or AL.Size <> 2 Then
            TestInsert = False
            Debug.Print "Start"
        End If
        'Middle
        AL.Insert 1, EG("C")
        If AL.JoinString("Message") <> "BCA" Or AL.Size <> 3 Then
            TestInsert = False
            Debug.Print "Middle"
        End If
        'End
        AL.Insert 3, EG("D")
        If AL.JoinString("Message") <> "BCAD" Or AL.Size <> 4 Then
            TestInsert = False
            Debug.Print "End"
        End If
        
    'Invalid lower bound
    On Error Resume Next
    AL.Insert -1, EG("Z")
    If Err.NUMBER <> 9 Then
        TestInsert = False
        Debug.Print "Invalid lower bound"
    End If
    On Error GoTo 0
    
    'Invalid upper bound
    On Error Resume Next
    AL.Insert 5, EG("Z")
    If Err.NUMBER <> 9 Then
        TestInsert = False
        Debug.Print "Invalid upper bound"
    End If
    On Error GoTo 0
    
    Debug.Print "TestInsert: " & TestInsert

End Function

Private Function TestInsertArray() As Boolean

    TestInsertArray = True

    Dim AL As clsArrayListObject
    Set AL = New clsArrayListObject

    Dim Arr(0 To 2) As Object

    'Empty
    Set Arr(0) = EG("A")
    Set Arr(1) = EG("B")
    Set Arr(2) = EG("C")
    AL.InsertArray 0, Arr
    If AL.JoinString("Message") <> "ABC" Or AL.Size <> 3 Or AL.Capacity <> DEFAULT_CAPACITY Then
        TestInsertArray = False
        Debug.Print "Empty"
    End If

    'Non empty
        'Start
        Set Arr(0) = EG("D")
        Set Arr(1) = EG("E")
        Set Arr(2) = EG("F")
        AL.InsertArray 0, Arr
        If AL.JoinString("Message") <> "DEFABC" Or AL.Size <> 6 Or AL.Capacity <> DEFAULT_CAPACITY Then
            TestInsertArray = False
            Debug.Print "Start"
        End If
        'Middle
        Set Arr(0) = EG("G")
        Set Arr(1) = EG("H")
        Set Arr(2) = EG("I")
        AL.InsertArray 3, Arr
        If AL.JoinString("Message") <> "DEFGHIABC" Or AL.Size <> 9 Or AL.Capacity <> DEFAULT_CAPACITY Then
            TestInsertArray = False
            Debug.Print "Middle"
        End If
        'End
        Set Arr(0) = EG("J")
        Set Arr(1) = EG("K")
        Set Arr(2) = EG("L")
        AL.InsertArray 9, Arr
        If AL.JoinString("Message") <> "DEFGHIABCJKL" Or AL.Size <> 12 Or AL.Capacity <> DEFAULT_CAPACITY * 2 + 2 Then
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
    Dim Arr1() As Object
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

    Dim AL As clsArrayListObject
    Set AL = New clsArrayListObject
    
    Dim AL1 As clsArrayListObject
    Set AL1 = New clsArrayListObject
    
    'Empty
    AL1.Reinitialize
    AL1.Append EG("A")
    AL1.Append EG("B")
    AL1.Append EG("C")
    AL.InsertArrayList 0, AL1
    If AL.JoinString("Message") <> "ABC" Or AL.Size <> 3 Or AL.Capacity <> DEFAULT_CAPACITY Then
        TestInsertArrayList = False
        Debug.Print "Empty"
    End If

    'Non empty
        'Start
        AL1.Reinitialize
        AL1.Append EG("D")
        AL1.Append EG("E")
        AL1.Append EG("F")
        AL.InsertArrayList 0, AL1
        If AL.JoinString("Message") <> "DEFABC" Or AL.Size <> 6 Or AL.Capacity <> DEFAULT_CAPACITY Then
            TestInsertArrayList = False
            Debug.Print "Start"
        End If
        'Middle
        AL1.Reinitialize
        AL1.Append EG("G")
        AL1.Append EG("H")
        AL1.Append EG("I")
        AL.InsertArrayList 3, AL1
        If AL.JoinString("Message") <> "DEFGHIABC" Or AL.Size <> 9 Or AL.Capacity <> DEFAULT_CAPACITY Then
            TestInsertArrayList = False
            Debug.Print "Middle"
        End If
        'End
        AL1.Reinitialize
        AL1.Append EG("J")
        AL1.Append EG("K")
        AL1.Append EG("L")
        AL.InsertArrayList 9, AL1
        If AL.JoinString("Message") <> "DEFGHIABCJKL" Or AL.Size <> 12 Or AL.Capacity <> DEFAULT_CAPACITY * 2 + 2 Then
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
    If AL.JoinString("Message") <> "" Or AL.Size <> 0 Or AL.Capacity <> DEFAULT_CAPACITY Then
        TestInsertArrayList = False
        Debug.Print "Empty ArrayList"
    End If

    Debug.Print "TestInsertArrayList: " & TestInsertArrayList

End Function

Private Function TestRemove() As Boolean

    TestRemove = True

    Dim AL As clsArrayListObject
    Set AL = New clsArrayListObject

    'Empty
    On Error Resume Next
    AL.Remove 0
    If Err.NUMBER <> 9 Then
        TestRemove = False
        Debug.Print "Empty"
    End If
    On Error GoTo 0
    
    'Non empty
    AL.Append EG("A")
    AL.Remove 0
    If AL.JoinString("Message") <> "" Or AL.Size <> 0 Then
        TestRemove = False
        Debug.Print "Non empty"
    End If

    'Start
    AL.Reinitialize
    AL.Append EG("A")
    AL.Append EG("B")
    AL.Append EG("C")
    AL.Remove 0
    If AL.JoinString("Message") <> "BC" Or AL.Size <> 2 Then
        TestRemove = False
        Debug.Print "Start"
    End If

    'Middle
    AL.Reinitialize
    AL.Append EG("A")
    AL.Append EG("B")
    AL.Append EG("C")
    AL.Remove 1
    If AL.JoinString("Message") <> "AC" Or AL.Size <> 2 Then
        TestRemove = False
        Debug.Print "Middle"
    End If

    'End
    AL.Reinitialize
    AL.Append EG("A")
    AL.Append EG("B")
    AL.Append EG("C")
    AL.Remove AL.Size - 1
    If AL.JoinString("Message") <> "AB" Or AL.Size <> 2 Then
        TestRemove = False
        Debug.Print "End"
    End If

    'Invalid lower bound
    AL.Reinitialize
    AL.Append EG("A")
    AL.Append EG("B")
    AL.Append EG("C")
    On Error Resume Next
    AL.Remove -1
    If Err.NUMBER <> 9 Then
        TestRemove = False
        Debug.Print "Invalid lower bound"
    End If

    'Invalid upper bound
    AL.Reinitialize
    AL.Append EG("A")
    AL.Append EG("B")
    AL.Append EG("C")
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

    Dim AL As clsArrayListObject
    Set AL = New clsArrayListObject

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
        AL.Append EG("A")
        AL.Append EG("B")
        AL.Append EG("C")
        AL.RemoveRange 0, 1
        If AL.JoinString("Message") <> "C" Or AL.Size <> 1 Then
            TestRemoveRange = False
            Debug.Print "Start"
        End If
        'Middle
        AL.Reinitialize
        AL.Append EG("A")
        AL.Append EG("B")
        AL.Append EG("C")
        AL.Append EG("D")
        AL.RemoveRange 1, 2
        If AL.JoinString("Message") <> "AD" Or AL.Size <> 2 Then
            TestRemoveRange = False
            Debug.Print "Middle"
        End If
        'End
        AL.Reinitialize
        AL.Append EG("A")
        AL.Append EG("B")
        AL.Append EG("C")
        AL.RemoveRange 1, 2
        If AL.JoinString("Message") <> "A" Or AL.Size <> 1 Then
            TestRemoveRange = False
            Debug.Print "End"
        End If
        

    'Invalid lower > upper
    AL.Reinitialize
    AL.Append EG("A")
    AL.Append EG("B")
    AL.Append EG("C")
    On Error Resume Next
    AL.RemoveRange 1, 0
    If Err.NUMBER <> 5 Then
        TestRemoveRange = False
        Debug.Print "Invalid lower > upper"
    End If
    On Error GoTo 0

    'Invalid lower bound
    AL.Reinitialize
    AL.Append EG("A")
    AL.Append EG("B")
    AL.Append EG("C")
    On Error Resume Next
    AL.RemoveRange -1, 0
    If Err.NUMBER <> 9 Then
        TestRemoveRange = False
        Debug.Print "Invalid lower bound"
    End If
    On Error GoTo 0

    'Invalid upper bound
    AL.Reinitialize
    AL.Append EG("A")
    AL.Append EG("B")
    AL.Append EG("C")
    On Error Resume Next
    AL.RemoveRange 2, 3
    If Err.NUMBER <> 9 Then
        TestRemoveRange = False
        Debug.Print "Invalid upper bound"
    End If
    On Error GoTo 0

    Debug.Print "TestRemoveRange: " & TestRemoveRange

End Function

Private Function TestRemoveFirst() As Boolean

    TestRemoveFirst = True

    Dim AL As clsArrayListObject
    Set AL = New clsArrayListObject

    'Empty
    On Error Resume Next
    AL.RemoveFirst "", "Message"
    If Err.NUMBER <> 0 Then
        TestRemoveFirst = False
        Debug.Print "Empty"
    End If
    On Error GoTo 0
    If AL.Size <> 0 Then
        TestRemoveFirst = False
        Debug.Print "Empty"
    End If

    'One
    AL.Append EG("A")
    AL.RemoveFirst "A", "Message"
    If AL.JoinString("Message") <> "" Or AL.Size <> 0 Then
        TestRemoveFirst = False
        Debug.Print "One"
    End If
    AL.Append EG("A")
    AL.RemoveFirst "", "Message"
    If AL.JoinString("Message") <> "A" Or AL.Size <> 1 Then
        TestRemoveFirst = False
        Debug.Print "One"
    End If
    
    'Multiple
    AL.Reinitialize
    AL.Append EG("A")
    AL.Append EG("B")
    AL.Append EG("A")
    AL.RemoveFirst "A", "Message"
    If AL.JoinString("Message") <> "BA" Or AL.Size <> 2 Then
       TestRemoveFirst = False
        Debug.Print "Multiple"
    End If
    
    'First
    AL.Reinitialize
    AL.Append EG("A")
    AL.Append EG("B")
    AL.Append EG("A")
    AL.RemoveFirst "A", "Message"
    If AL.JoinString("Message") <> "BA" Or AL.Size <> 2 Then
       TestRemoveFirst = False
        Debug.Print "First"
    End If

    'Middle
    AL.Reinitialize
    AL.Append EG("B")
    AL.Append EG("A")
    AL.Append EG("B")
    AL.Append EG("A")
    AL.RemoveFirst "A", "Message"
    If AL.JoinString("Message") <> "BBA" Or AL.Size <> 3 Then
       TestRemoveFirst = False
        Debug.Print "Middle"
    End If

    'End
    AL.Reinitialize
    AL.Append EG("B")
    AL.Append EG("B")
    AL.Append EG("A")
    AL.RemoveFirst "A", "Message"
    If AL.JoinString("Message") <> "BB" Or AL.Size <> 2 Then
       TestRemoveFirst = False
        Debug.Print "End"
    End If

    'Not there
    AL.Reinitialize
    AL.Append EG("A")
    AL.Append EG("B")
    AL.Append EG("C")
    AL.RemoveFirst "Z", "Message"
    If AL.JoinString("Message") <> "ABC" Or AL.Size <> 3 Then
       TestRemoveFirst = False
        Debug.Print "Not there"
    End If

    'CompareMethod
    AL.Reinitialize
    AL.Append EG("A")
    AL.Append EG("B")
    AL.Append EG("C")
    AL.RemoveFirst "a", "Message"
    If AL.JoinString("Message") <> "ABC" Or AL.Size <> 3 Then
       TestRemoveFirst = False
        Debug.Print "Binary"
    End If
    
    'CallType
    AL.Reinitialize
    AL.Append EG("A")
    AL.Append EG("B")
    AL.Append EG("C")
    AL.RemoveFirst 1, "GetValue", True
    If AL.JoinString("GetValue", , True) <> "11" Or AL.Size <> 2 Then
       TestRemoveFirst = False
        Debug.Print "CallType"
    End If
    
    Debug.Print "TestRemoveFirst: " & TestRemoveFirst

End Function

Private Function TestRemoveLast() As Boolean

    TestRemoveLast = True
    
    Dim AL As clsArrayListObject
    Set AL = New clsArrayListObject
    
    'Empty
    On Error Resume Next
    AL.RemoveLast "", "Message"
    If Err.NUMBER <> 0 Then
        TestRemoveLast = False
        Debug.Print "Empty"
    End If
    On Error GoTo 0
    If AL.Size <> 0 Then
        TestRemoveLast = False
        Debug.Print "Empty"
    End If

    'One
    AL.Append EG("A")
    AL.RemoveLast "A", "Message"
    If AL.JoinString("Message") <> "" Or AL.Size <> 0 Then
        TestRemoveLast = False
        Debug.Print "One"
    End If
    AL.Append EG("A")
    AL.RemoveLast "", "Message"
    If AL.JoinString("Message") <> "A" Or AL.Size <> 1 Then
        TestRemoveLast = False
        Debug.Print "One"
    End If
    
    'Multiple
    AL.Reinitialize
    AL.Append EG("A")
    AL.Append EG("B")
    AL.Append EG("A")
    AL.RemoveLast "A", "Message"
    If AL.JoinString("Message") <> "AB" Or AL.Size <> 2 Then
       TestRemoveLast = False
        Debug.Print "Multiple"
    End If
    
    'First
    AL.Reinitialize
    AL.Append EG("A")
    AL.Append EG("B")
    AL.Append EG("C")
    AL.RemoveLast "A", "Message"
    If AL.JoinString("Message") <> "BC" Or AL.Size <> 2 Then
       TestRemoveLast = False
        Debug.Print "First"
    End If

    'Middle
    AL.Reinitialize
    AL.Append EG("A")
    AL.Append EG("B")
    AL.Append EG("A")
    AL.Append EG("B")
    AL.RemoveLast "A", "Message"
    If AL.JoinString("Message") <> "ABB" Or AL.Size <> 3 Then
       TestRemoveLast = False
        Debug.Print "Middle"
    End If

    'End
    AL.Reinitialize
    AL.Append EG("A")
    AL.Append EG("B")
    AL.Append EG("B")
    AL.Append EG("A")
    AL.RemoveLast "A", "Message"
    If AL.JoinString("Message") <> "ABB" Or AL.Size <> 3 Then
       TestRemoveLast = False
        Debug.Print "End"
    End If

    'Not there
    AL.Reinitialize
    AL.Append EG("A")
    AL.Append EG("B")
    AL.Append EG("C")
    AL.RemoveLast "Z", "Message"
    If AL.JoinString("Message") <> "ABC" Or AL.Size <> 3 Then
       TestRemoveLast = False
        Debug.Print "Not there"
    End If
    
    'CallType
    AL.Reinitialize
    AL.Append EG("A")
    AL.Append EG("B")
    AL.Append EG("C")
    AL.RemoveLast 1, "GetValue", True
    If AL.JoinString("GetValue", , True) <> "11" Or AL.Size <> 2 Then
       TestRemoveLast = False
        Debug.Print "CallType"
    End If
    
    Debug.Print "TestRemoveLast: " & TestRemoveLast

End Function

Private Function TestRemoveAll() As Boolean

    TestRemoveAll = True
    
    Dim AL As clsArrayListObject
    Set AL = New clsArrayListObject
    
    'Empty
    On Error Resume Next
    AL.RemoveAll "", "Message"
    If Err.NUMBER <> 0 Then
        TestRemoveAll = False
        Debug.Print "Empty"
    End If
    On Error GoTo 0
    If AL.Size <> 0 Then
        TestRemoveAll = False
        Debug.Print "Empty"
    End If

    'One
    AL.Append EG("A")
    AL.RemoveAll "A", "Message"
    If AL.JoinString("Message") <> "" Or AL.Size <> 0 Then
        TestRemoveAll = False
        Debug.Print "One"
    End If
    AL.Append EG("A")
    AL.RemoveAll "", "Message"
    If AL.JoinString("Message") <> "A" Or AL.Size <> 1 Then
        TestRemoveAll = False
        Debug.Print "One"
    End If
    
    'Multiple
    AL.Reinitialize
    AL.Append EG("A")
    AL.Append EG("B")
    AL.Append EG("A")
    AL.RemoveAll "A", "Message"
    If AL.JoinString("Message") <> "B" Or AL.Size <> 1 Then
       TestRemoveAll = False
       Debug.Print "Multiple"
    End If
    
    'First
    AL.Reinitialize
    AL.Append EG("A")
    AL.Append EG("B")
    AL.Append EG("C")
    AL.RemoveAll "A", "Message"
    If AL.JoinString("Message") <> "BC" Or AL.Size <> 2 Then
       TestRemoveAll = False
       Debug.Print "First"
    End If

    'Middle
    AL.Reinitialize
    AL.Append EG("A")
    AL.Append EG("B")
    AL.Append EG("A")
    AL.Append EG("B")
    AL.RemoveAll "A", "Message"
    If AL.JoinString("Message") <> "BB" Or AL.Size <> 2 Then
       TestRemoveAll = False
       Debug.Print "Middle"
    End If

    'End
    AL.Reinitialize
    AL.Append EG("A")
    AL.Append EG("B")
    AL.Append EG("B")
    AL.Append EG("A")
    AL.RemoveAll "A", "Message"
    If AL.JoinString("Message") <> "BB" Or AL.Size <> 2 Then
       TestRemoveAll = False
       Debug.Print "End"
    End If

    'Not there
    AL.Reinitialize
    AL.Append EG("A")
    AL.Append EG("B")
    AL.Append EG("C")
    AL.RemoveAll "Z", "Message"
    If AL.JoinString("Message") <> "ABC" Or AL.Size <> 3 Then
       TestRemoveAll = False
       Debug.Print "Not there"
    End If
    
    'CallType
    AL.Reinitialize
    AL.Append EG("A")
    AL.Append EG("B")
    AL.Append EG("C")
    AL.RemoveAll 1, "GetValue", True
    If AL.JoinString("GetValue", , True) <> "" Or AL.Size <> 0 Then
       TestRemoveAll = False
        Debug.Print "CallType"
    End If
    
    Debug.Print "TestRemoveAll: " & TestRemoveAll

End Function

Private Function TestReplaceAll() As Boolean

    TestReplaceAll = True
    
    Dim AL As clsArrayListObject
    Set AL = New clsArrayListObject
    
    'Empty
    On Error Resume Next
    AL.ReplaceAll "", "A", "Message"
    If Err.NUMBER <> 0 Then
        TestReplaceAll = False
        Debug.Print "Empty"
    End If
    On Error GoTo 0
    If AL.Size <> 0 Then
        TestReplaceAll = False
        Debug.Print "Empty"
    End If

    'One
    AL.Append EG("A")
    AL.ReplaceAll "A", "B", "Message"
    If AL.JoinString("Message") <> "B" Or AL.Size <> 1 Then
        TestReplaceAll = False
        Debug.Print "One"
    End If
    AL.ReplaceAll "", "A", "Message"
    If AL.JoinString("Message") <> "B" Or AL.Size <> 1 Then
        TestReplaceAll = False
        Debug.Print "One"
    End If
    
    'Multiple
    AL.Reinitialize
    AL.Append EG("A")
    AL.Append EG("B")
    AL.Append EG("A")
    AL.ReplaceAll "A", "Z", "Message"
    If AL.JoinString("Message") <> "ZBZ" Or AL.Size <> 3 Then
       TestReplaceAll = False
       Debug.Print "Multiple"
    End If
    
    'First
    AL.Reinitialize
    AL.Append EG("A")
    AL.Append EG("B")
    AL.Append EG("C")
    AL.ReplaceAll "A", "Z", "Message"
    If AL.JoinString("Message") <> "ZBC" Or AL.Size <> 3 Then
       TestReplaceAll = False
       Debug.Print "First"
    End If

    'Middle
    AL.Reinitialize
    AL.Append EG("A")
    AL.Append EG("B")
    AL.Append EG("A")
    AL.Append EG("B")
    AL.ReplaceAll "A", "Z", "Message"
    If AL.JoinString("Message") <> "ZBZB" Or AL.Size <> 4 Then
       TestReplaceAll = False
       Debug.Print "Middle"
    End If

    'End
    AL.Reinitialize
    AL.Append EG("A")
    AL.Append EG("B")
    AL.Append EG("B")
    AL.Append EG("A")
    AL.ReplaceAll "A", "Z", "Message"
    If AL.JoinString("Message") <> "ZBBZ" Or AL.Size <> 4 Then
       TestReplaceAll = False
       Debug.Print "End"
    End If

    'Not there
    AL.Reinitialize
    AL.Append EG("A")
    AL.Append EG("B")
    AL.Append EG("C")
    AL.ReplaceAll "Z", "", "Message"
    If AL.JoinString("Message") <> "ABC" Or AL.Size <> 3 Then
       TestReplaceAll = False
       Debug.Print "Not there"
    End If
    
    Debug.Print "TestReplaceAll: " & TestReplaceAll

End Function

Private Function TestIndexOf() As Boolean

    TestIndexOf = True
    
    Dim AL As clsArrayListObject
    Set AL = New clsArrayListObject
    
    'Empty
    If AL.IndexOf("", "Message") <> -1 Then
        TestIndexOf = False
        Debug.Print "Empty"
    End If

    'One
    AL.Append EG("A")
    If AL.IndexOf("A", "Message") <> 0 Then
        TestIndexOf = False
        Debug.Print "One"
    End If
    If AL.IndexOf("", "Message") <> -1 Then
        TestIndexOf = False
        Debug.Print "One"
    End If
    
    'Multiple
    AL.Reinitialize
    AL.Append EG("A")
    AL.Append EG("B")
    AL.Append EG("A")
    If AL.IndexOf("A", "Message") <> 0 Then
       TestIndexOf = False
       Debug.Print "Multiple"
    End If
    
    'First
    AL.Reinitialize
    AL.Append EG("A")
    AL.Append EG("B")
    AL.Append EG("C")
    If AL.IndexOf("A", "Message") <> 0 Then
       TestIndexOf = False
       Debug.Print "First"
    End If

    'Middle
    AL.Reinitialize
    AL.Append EG("B")
    AL.Append EG("A")
    AL.Append EG("B")
    AL.Append EG("A")
    If AL.IndexOf("A", "Message") <> 1 Then
       TestIndexOf = False
       Debug.Print "Middle"
    End If

    'End
    AL.Reinitialize
    AL.Append EG("B")
    AL.Append EG("B")
    AL.Append EG("A")
    If AL.IndexOf("A", "Message") <> 2 Then
       TestIndexOf = False
       Debug.Print "End"
    End If

    'Not there
    AL.Reinitialize
    AL.Append EG("A")
    AL.Append EG("B")
    AL.Append EG("C")
    If AL.IndexOf("Z", "Message") <> -1 Then
       TestIndexOf = False
       Debug.Print "Not there"
    End If

    'From
    AL.Reinitialize
    AL.Append EG("A")
    AL.Append EG("A")
    AL.Append EG("B")
    AL.Append EG("C")
    AL.Append EG("A")
    If AL.IndexOf("A", "Message", 0) <> 0 Then
       TestIndexOf = False
       Debug.Print "From"
    End If
    If AL.IndexOf("A", "Message", 1) <> 1 Then
       TestIndexOf = False
       Debug.Print "From"
    End If
    If AL.IndexOf("A", "Message", 2) <> 4 Then
       TestIndexOf = False
       Debug.Print "From"
    End If
    If AL.IndexOf("A", "Message", 3) <> 4 Then
       TestIndexOf = False
       Debug.Print "From"
    End If
    If AL.IndexOf("A", "Message", 4) <> 4 Then
       TestIndexOf = False
       Debug.Print "From"
    End If
    On Error Resume Next
    AL.IndexOf "A", "Message", 5
    If Err.NUMBER <> 9 Then
       TestIndexOf = False
       Debug.Print "From"
    End If
    On Error GoTo 0
    
    'From negative
    AL.Reinitialize
    AL.Append EG("A")
    AL.Append EG("A")
    AL.Append EG("B")
    AL.Append EG("C")
    AL.Append EG("A")
    If AL.IndexOf("A", "Message", -1) <> 4 Then
       TestIndexOf = False
       Debug.Print "From negative"
    End If
    If AL.IndexOf("A", "Message", -2) <> 4 Then
       TestIndexOf = False
       Debug.Print "From negative"
    End If
    If AL.IndexOf("A", "Message", -3) <> 4 Then
       TestIndexOf = False
       Debug.Print "From negative"
    End If
    If AL.IndexOf("A", "Message", -4) <> 1 Then
       TestIndexOf = False
       Debug.Print "From negative"
    End If
    If AL.IndexOf("A", "Message", -5) <> 0 Then
       TestIndexOf = False
       Debug.Print "From negative"
    End If
    On Error Resume Next
    AL.IndexOf "A", "Message", -6
    If Err.NUMBER <> 9 Then
       TestIndexOf = False
       Debug.Print "From"
    End If
    On Error GoTo 0
    
    'CallType
    AL.Reinitialize
    AL.Append EG("A")
    AL.Append EG("B")
    AL.Append EG("C")
    If AL.IndexOf(1, "GetValue", , True) <> 0 Then
       TestIndexOf = False
       Debug.Print "CallType"
    End If
    
    Debug.Print "TestIndexOf: " & TestIndexOf

End Function

Private Function TestLastIndexOf() As Boolean

    TestLastIndexOf = True

    Dim AL As clsArrayListObject
    Set AL = New clsArrayListObject

    'Empty
    If AL.LastIndexOf("", "Message") <> -1 Then
        TestLastIndexOf = False
        Debug.Print "Empty"
    End If

    'One
    AL.Append EG("A")
    If AL.LastIndexOf("A", "Message") <> 0 Then
        TestLastIndexOf = False
        Debug.Print "One"
    End If
    If AL.LastIndexOf("", "Message") <> -1 Then
        TestLastIndexOf = False
        Debug.Print "One"
    End If

    'Multiple
    AL.Reinitialize
    AL.Append EG("A")
    AL.Append EG("B")
    AL.Append EG("A")
    If AL.LastIndexOf("A", "Message") <> 2 Then
       TestLastIndexOf = False
       Debug.Print "Multiple"
    End If

    'First
    AL.Reinitialize
    AL.Append EG("A")
    AL.Append EG("B")
    AL.Append EG("C")
    If AL.LastIndexOf("A", "Message") <> 0 Then
       TestLastIndexOf = False
       Debug.Print "First"
    End If

    'Middle
    AL.Reinitialize
    AL.Append EG("A")
    AL.Append EG("B")
    AL.Append EG("A")
    AL.Append EG("B")
    If AL.LastIndexOf("A", "Message") <> 2 Then
       TestLastIndexOf = False
       Debug.Print "Middle"
    End If

    'End
    AL.Reinitialize
    AL.Append EG("B")
    AL.Append EG("B")
    AL.Append EG("A")
    If AL.LastIndexOf("A", "Message") <> 2 Then
       TestLastIndexOf = False
       Debug.Print "End"
    End If

    'Not there
    AL.Reinitialize
    AL.Append EG("A")
    AL.Append EG("B")
    AL.Append EG("C")
    If AL.LastIndexOf("Z", "Message") <> -1 Then
       TestLastIndexOf = False
       Debug.Print "Not there"
    End If

    'From
    AL.Reinitialize
    AL.Append EG("A")
    AL.Append EG("A")
    AL.Append EG("B")
    AL.Append EG("A")
    AL.Append EG("C")
    If AL.LastIndexOf("A", "Message", 4) <> 3 Then
       TestLastIndexOf = False
       Debug.Print "From"
    End If
    If AL.LastIndexOf("A", "Message", 3) <> 3 Then
       TestLastIndexOf = False
       Debug.Print "From"
    End If
    If AL.LastIndexOf("A", "Message", 2) <> 1 Then
       TestLastIndexOf = False
       Debug.Print "From"
    End If
    If AL.LastIndexOf("A", "Message", 1) <> 1 Then
       TestLastIndexOf = False
       Debug.Print "From"
    End If
    If AL.LastIndexOf("A", "Message", 0) <> 0 Then
       TestLastIndexOf = False
       Debug.Print "From"
    End If
    On Error Resume Next
    AL.LastIndexOf "A", "Message", 5
    If Err.NUMBER <> 9 Then
       TestLastIndexOf = False
       Debug.Print "From"
    End If
    On Error GoTo 0

    'From negative
    AL.Reinitialize
    AL.Append EG("A")
    AL.Append EG("A")
    AL.Append EG("B")
    AL.Append EG("C")
    AL.Append EG("A")
    If AL.LastIndexOf("A", "Message", -1) <> 4 Then
       TestLastIndexOf = False
       Debug.Print "From negative"
    End If
    If AL.LastIndexOf("A", "Message", -2) <> 1 Then
       TestLastIndexOf = False
       Debug.Print "From negative"
    End If
    If AL.LastIndexOf("A", "Message", -3) <> 1 Then
       TestLastIndexOf = False
       Debug.Print "From negative"
    End If
    If AL.LastIndexOf("A", "Message", -4) <> 1 Then
       TestLastIndexOf = False
       Debug.Print "From negative"
    End If
    If AL.LastIndexOf("A", "Message", -5) <> 0 Then
       TestLastIndexOf = False
       Debug.Print "From negative"
    End If
    On Error Resume Next
    AL.LastIndexOf "A", "Message", -6
    If Err.NUMBER <> 9 Then
       TestLastIndexOf = False
       Debug.Print "From"
    End If
    On Error GoTo 0
    
    'CallType
    AL.Reinitialize
    AL.Append EG("A")
    AL.Append EG("B")
    AL.Append EG("C")
    If AL.LastIndexOf(1, "GetValue", , True) <> 2 Then
       TestLastIndexOf = False
       Debug.Print "CallType"
    End If
    
    Debug.Print "TestLastIndexOf: " & TestLastIndexOf

End Function

Private Function TestCount() As Boolean

    TestCount = True
    
    Dim AL As clsArrayListObject
    Set AL = New clsArrayListObject
    
    'Empty
    If AL.Count("A", "Message") <> 0 Then
        TestCount = False
        Debug.Print "Empty"
    End If
    If AL.Count("", "Message") <> 0 Then
        TestCount = False
        Debug.Print "Empty"
    End If
    
    'One
    AL.Append EG("A")
    If AL.Count("A", "Message") <> 1 Then
        TestCount = False
        Debug.Print "One"
    End If
    If AL.Count("", "Message") <> 0 Then
        TestCount = False
        Debug.Print "One"
    End If
    
    'Multiple
    AL.Append EG("A")
    If AL.Count("A", "Message") <> 2 Then
        TestCount = False
        Debug.Print "Multiple"
    End If
    If AL.Count("", "Message") <> 0 Then
        TestCount = False
        Debug.Print "Multiple"
    End If
    
    'First
    AL.Reinitialize
    AL.Append EG("A")
    AL.Append EG("B")
    AL.Append EG("C")
    If AL.Count("A", "Message") <> 1 Then
        TestCount = False
        Debug.Print "First"
    End If
    
    'Middle
    AL.Reinitialize
    AL.Append EG("B")
    AL.Append EG("A")
    AL.Append EG("C")
    If AL.Count("A", "Message") <> 1 Then
        TestCount = False
        Debug.Print "Middle"
    End If
    
    'End
    AL.Reinitialize
    AL.Append EG("C")
    AL.Append EG("B")
    AL.Append EG("A")
    If AL.Count("A", "Message") <> 1 Then
        TestCount = False
        Debug.Print "End"
    End If
    
    'Not there
    If AL.Count("Z", "Message") <> 0 Then
        TestCount = False
        Debug.Print "Not there"
    End If
    
    'CallType
    AL.Reinitialize
    AL.Append EG("A")
    AL.Append EG("B")
    AL.Append EG("C")
    If AL.Count(1, "GetValue", True) <> 3 Then
        TestCount = False
        Debug.Print "CallType"
    End If
    
    Debug.Print "TestCount: " & TestCount

End Function

Private Function TestContains() As Boolean

    TestContains = True

    Dim AL As clsArrayListObject
    Set AL = New clsArrayListObject
    
    'Empty
    If AL.Contains("A", "Message") <> False Then
        TestContains = False
        Debug.Print "Empty"
    End If
    If AL.Contains("", "Message") <> False Then
        TestContains = False
        Debug.Print "Empty"
    End If
    
    'One
    AL.Append EG("A")
    If AL.Contains("A", "Message") <> True Then
        TestContains = False
        Debug.Print "One"
    End If
    If AL.Contains("", "Message") <> False Then
        TestContains = False
        Debug.Print "One"
    End If
    
    'Multiple
    AL.Append EG("A")
    If AL.Contains("A", "Message") <> True Then
        TestContains = False
        Debug.Print "Multiple"
    End If
    If AL.Contains("", "Message") <> False Then
        TestContains = False
        Debug.Print "Multiple"
    End If
    
    'First
    AL.Reinitialize
    AL.Append EG("A")
    AL.Append EG("B")
    AL.Append EG("C")
    If AL.Contains("A", "Message") <> True Then
        TestContains = False
        Debug.Print "First"
    End If
    
    'Middle
    AL.Reinitialize
    AL.Append EG("B")
    AL.Append EG("A")
    AL.Append EG("C")
    If AL.Contains("A", "Message") <> True Then
        TestContains = False
        Debug.Print "Middle"
    End If
    
    'End
    AL.Reinitialize
    AL.Append EG("C")
    AL.Append EG("B")
    AL.Append EG("A")
    If AL.Contains("A", "Message") <> True Then
        TestContains = False
        Debug.Print "End"
    End If
    
    'Not there
    If AL.Contains("Z", "Message") <> False Then
        TestContains = False
        Debug.Print "Not there"
    End If
    
    'CallType
    AL.Reinitialize
    AL.Append EG("A")
    AL.Append EG("B")
    AL.Append EG("C")
    If AL.Contains(1, "GetValue", True) <> True Then
        TestContains = False
        Debug.Print "CallType"
    End If

    Debug.Print "TestContains: " & TestContains

End Function

Private Function TestClear() As Boolean

    TestClear = True
    
    Dim AL As clsArrayListObject
    Set AL = New clsArrayListObject
    
    'Empty
    AL.Clear
    If AL.JoinString("Message") <> "" Or AL.Size <> 0 Then
        TestClear = False
        Debug.Print "Empty"
    End If
    
    'One
    AL.Append EG("A")
    AL.Clear
    If AL.JoinString("Message") <> "" Or AL.Size <> 0 Then
        TestClear = False
        Debug.Print "One"
    End If

    'Multiple
    AL.Append EG("A")
    AL.Append EG("B")
    AL.Append EG("C")
    AL.Clear
    If AL.JoinString("Message") <> "" Or AL.Size <> 0 Then
        TestClear = False
        Debug.Print "Multiple"
    End If
    
    'Capacity persists
    Dim i&
    For i = 1 To 11
        AL.Append EG("A")
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
    
    Dim AL As clsArrayListObject
    Set AL = New clsArrayListObject
    
    'Empty
    On Error Resume Next
    AL.Reverse
    If Err.NUMBER <> 0 Then
        TestReverse = False
        Debug.Print "Empty"
    End If
    On Error GoTo 0
    
    'One
    AL.Append EG("A")
    AL.Reverse
    If AL.JoinString("Message") <> "A" Then
        TestReverse = False
        Debug.Print "One"
    End If
    
    'Multiple
    AL.Append EG("B")
    AL.Append EG("C")
    AL.Reverse
    If AL.JoinString("Message") <> "CBA" Then
        TestReverse = False
        Debug.Print "Multiple"
    End If
    
    'Even
    AL.Reinitialize
    AL.Append EG("A")
    AL.Append EG("B")
    AL.Append EG("C")
    AL.Append EG("D")
    AL.Reverse
    If AL.JoinString("Message") <> "DCBA" Then
        TestReverse = False
        Debug.Print "Even"
    End If
    
    'Odd
    AL.Reinitialize
    AL.Append EG("A")
    AL.Append EG("B")
    AL.Append EG("C")
    AL.Reverse
    If AL.JoinString("Message") <> "CBA" Then
        TestReverse = False
        Debug.Print "Odd"
    End If
    
    Debug.Print "TestReverse: " & TestReverse

End Function

Private Function TestToArray() As Boolean

    TestToArray = True
    
    Dim AL As clsArrayListObject
    Set AL = New clsArrayListObject
    
    'Empty
    Dim Arr() As Object
    Arr = AL.ToArray
    On Error Resume Next
    Debug.Print LBound(Arr)
    If Err.NUMBER <> 9 Then
        TestToArray = False
        Debug.Print "Empty"
    End If
    On Error GoTo 0
    
    'One
    AL.Append EG("A")
    Arr = AL.ToArray
    If UBound(Arr) - LBound(Arr) + 1 <> AL.Size Then
        TestToArray = False
        Debug.Print "One"
    End If
    If "A" <> AL.JoinString("Message") Then
        TestToArray = False
        Debug.Print "One"
    End If
    
    'Multiple
    AL.Append EG("B")
    AL.Append EG("C")
    Arr = AL.ToArray
    If UBound(Arr) - LBound(Arr) + 1 <> AL.Size Then
        TestToArray = False
        Debug.Print "Multiple"
    End If
    If "ABC" <> AL.JoinString("Message") Then
        TestToArray = False
        Debug.Print "Multiple"
    End If
    
    Debug.Print "TestToArray: " & TestToArray

End Function

Private Function TestJoinString() As Boolean

    TestJoinString = True
    
    Dim AL As clsArrayListObject
    Set AL = New clsArrayListObject
    
    'No delimiter
        'Empty
        If AL.JoinString("Message") <> "" Then
            TestJoinString = False
            Debug.Print "No delimiter Empty"
        End If
        'One
        AL.Append EG("A")
        If AL.JoinString("Message") <> "A" Then
            TestJoinString = False
            Debug.Print "No delimiter One"
        End If
        'Multiple
        AL.Append EG("B")
        AL.Append EG("C")
        If AL.JoinString("Message") <> "ABC" Then
            TestJoinString = False
            Debug.Print "No delimiter Multiple"
        End If
    
    'Delimiter
        AL.Reinitialize
        'Empty
        If AL.JoinString("Message", ",") <> "" Then
            TestJoinString = False
            Debug.Print "Delimiter Empty"
        End If
        'One
        AL.Append EG("A")
        If AL.JoinString("Message", ",") <> "A" Then
            TestJoinString = False
            Debug.Print "Delimiter One"
        End If
        'Multiple
        AL.Append EG("B")
        AL.Append EG("C")
        If AL.JoinString("Message", ",") <> "A,B,C" Then
            TestJoinString = False
            Debug.Print "Delimiter Multiple"
        End If
        
    'CallType
    AL.Reinitialize
    AL.Append EG("A")
    AL.Append EG("B")
    AL.Append EG("C")
    If AL.JoinString("GetValue", , True) <> "111" Then
        TestJoinString = False
        Debug.Print "CallType"
    End If
    
    Debug.Print "TestJoinString: " & TestJoinString

End Function

Private Function TestCallMethod() As Boolean

    TestCallMethod = True
    
    Dim AL As clsArrayListObject
    Set AL = New clsArrayListObject
    
    'Empty
    On Error Resume Next
    AL.CallMethod "MethodSub"
    If Err.NUMBER <> 0 Then
        TestCallMethod = False
        Debug.Print "Empty"
    End If
    On Error GoTo 0
    
    AL.Append EG("A")
    AL.Append EG("B")
    AL.Append EG("C")
    
    'Sub
    On Error Resume Next
    AL.CallMethod "MethodSub"
    If Err.NUMBER <> 0 Then
        TestCallMethod = False
        Debug.Print "Sub"
    End If
    On Error GoTo 0
    
    'Function
    On Error Resume Next
    AL.CallMethod "MethodFunction"
    If Err.NUMBER <> 0 Then
        TestCallMethod = False
        Debug.Print "Function"
    End If
    On Error GoTo 0
    
    'Arg
    On Error Resume Next
    AL.CallMethod "MethodArg"
    If Err.NUMBER <> 449 Then
        TestCallMethod = False
        Debug.Print "Arg"
    End If
    On Error GoTo 0
    
    Debug.Print "TestCallMethod: " & TestCallMethod
    
End Function

Private Function TestSort() As Boolean

    TestSort = True

    Dim AL As clsArrayListObject
    Set AL = New clsArrayListObject

    AL.Reinitialize
    '(Insertion Sort)
        'Empty
        AL.Sort "Message"
        If AL.JoinString("Message") <> "" Then
            TestSort = False
            Debug.Print "Insertion Sort Empty"
        End If
        'One
        AL.Append EG("A")
        AL.Sort "Message"
        If AL.JoinString("Message") <> "A" Then
            TestSort = False
            Debug.Print "Insertion Sort One"
        End If
        'Many
        AL.Reinitialize
        AL.Append EG("A")
        AL.Append EG("M")
        AL.Append EG("F")
        AL.Append EG("I")
        AL.Append EG("J")
        AL.Append EG("U")
        AL.Append EG("K")
        AL.Append EG("W")
        AL.Append EG("R")
        AL.Append EG("Z")
        AL.Append EG("B")
        AL.Append EG("X")
        AL.Append EG("S")
        AL.Append EG("E")
        AL.Append EG("Y")
        AL.Append EG("G")
        AL.Append EG("H")
        AL.Append EG("N")
        AL.Append EG("O")
        AL.Append EG("P")
        AL.Append EG("C")
        AL.Append EG("D")
        AL.Append EG("Q")
        AL.Append EG("L")
        AL.Append EG("T")
        AL.Append EG("V")
        AL.Sort "Message"
        If AL.JoinString("Message") <> "ABCDEFGHIJKLMNOPQRSTUVWXYZ" Then
            TestSort = False
            Debug.Print "Insertion Sort Many"
        End If
        'Binary
        AL.Reinitialize
        AL.Append EG("l")
        AL.Append EG("M")
        AL.Append EG("N")
        AL.Append EG("B")
        AL.Append EG("f")
        AL.Append EG("G")
        AL.Append EG("F")
        AL.Append EG("H")
        AL.Append EG("D")
        AL.Append EG("e")
        AL.Append EG("O")
        AL.Append EG("c")
        AL.Append EG("d")
        AL.Append EG("p")
        AL.Append EG("J")
        AL.Append EG("K")
        AL.Append EG("A")
        AL.Append EG("n")
        AL.Append EG("E")
        AL.Append EG("P")
        AL.Append EG("a")
        AL.Append EG("L")
        AL.Append EG("b")
        AL.Append EG("C")
        AL.Append EG("m")
        AL.Append EG("o")
        AL.Append EG("g")
        AL.Append EG("h")
        AL.Append EG("i")
        AL.Append EG("j")
        AL.Append EG("k")
        AL.Append EG("I")
        AL.Sort "Message"
        If AL.JoinString("Message") <> "ABCDEFGHIJKLMNOPabcdefghijklmnop" Then
            TestSort = False
            Debug.Print "Insertion Sort CompareMethod Binary"
        End If

    AL.Reinitialize
    '(Merge Sort)
        'Empty
        AL.Sort "Message"
        If AL.JoinString("Message") <> "" Then
            TestSort = False
            Debug.Print "Merge Sort Empty"
        End If
        'One
        AL.Append EG("A")
        AL.Sort "Message"
        If AL.JoinString("Message") <> "A" Then
            TestSort = False
            Debug.Print "Merge Sort One"
        End If
        'Many
        AL.Reinitialize
        AL.Append EG("A")
        AL.Append EG("M")
        AL.Append EG("F")
        AL.Append EG("I")
        AL.Append EG("J")
        AL.Append EG("U")
        AL.Append EG("K")
        AL.Append EG("W")
        AL.Append EG("R")
        AL.Append EG("Z")
        AL.Append EG("B")
        AL.Append EG("X")
        AL.Append EG("S")
        AL.Append EG("E")
        AL.Append EG("Y")
        AL.Append EG("G")
        AL.Append EG("H")
        AL.Append EG("N")
        AL.Append EG("O")
        AL.Append EG("P")
        AL.Append EG("C")
        AL.Append EG("D")
        AL.Append EG("Q")
        AL.Append EG("L")
        AL.Append EG("T")
        AL.Append EG("V")
        AL.Append EG("A")
        AL.Append EG("M")
        AL.Append EG("F")
        AL.Append EG("I")
        AL.Append EG("J")
        AL.Append EG("U")
        AL.Append EG("K")
        AL.Append EG("W")
        AL.Append EG("R")
        AL.Append EG("Z")
        AL.Append EG("B")
        AL.Append EG("X")
        AL.Append EG("S")
        AL.Append EG("E")
        AL.Append EG("Y")
        AL.Append EG("G")
        AL.Append EG("H")
        AL.Append EG("N")
        AL.Append EG("O")
        AL.Append EG("P")
        AL.Append EG("C")
        AL.Append EG("D")
        AL.Append EG("Q")
        AL.Append EG("L")
        AL.Append EG("T")
        AL.Append EG("V")
        AL.Sort "Message"
        If AL.JoinString("Message") <> "AABBCCDDEEFFGGHHIIJJKKLLMMNNOOPPQQRRSSTTUUVVWWXXYYZZ" Then
            TestSort = False
            Debug.Print "Merge Sort Many"
        End If
        'Binary
        AL.Reinitialize
        AL.Append EG("w")
        AL.Append EG("n")
        AL.Append EG("z")
        AL.Append EG("Z")
        AL.Append EG("a")
        AL.Append EG("y")
        AL.Append EG("N")
        AL.Append EG("J")
        AL.Append EG("h")
        AL.Append EG("i")
        AL.Append EG("j")
        AL.Append EG("b")
        AL.Append EG("c")
        AL.Append EG("d")
        AL.Append EG("x")
        AL.Append EG("X")
        AL.Append EG("V")
        AL.Append EG("s")
        AL.Append EG("o")
        AL.Append EG("A")
        AL.Append EG("m")
        AL.Append EG("t")
        AL.Append EG("u")
        AL.Append EG("E")
        AL.Append EG("F")
        AL.Append EG("G")
        AL.Append EG("L")
        AL.Append EG("k")
        AL.Append EG("K")
        AL.Append EG("v")
        AL.Append EG("Y")
        AL.Append EG("g")
        AL.Append EG("W")
        AL.Append EG("H")
        AL.Append EG("I")
        AL.Append EG("B")
        AL.Append EG("O")
        AL.Append EG("P")
        AL.Append EG("l")
        AL.Append EG("p")
        AL.Append EG("q")
        AL.Append EG("r")
        AL.Append EG("U")
        AL.Append EG("C")
        AL.Append EG("D")
        AL.Append EG("e")
        AL.Append EG("f")
        AL.Append EG("M")
        AL.Append EG("Q")
        AL.Append EG("R")
        AL.Append EG("S")
        AL.Append EG("T")
        AL.Sort "Message"
        If AL.JoinString("Message") <> "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz" Then
            TestSort = False
            Debug.Print "MergeSort CompareMethod Binary"
        End If
 
    'CallType
    AL.Reinitialize
    AL.Append EG("C")
    AL.Append EG("B")
    AL.Append EG("A")
    AL.Sort "GetMessage", True
    If AL.JoinString("Message") <> "ABC" Then
        TestSort = False
        Debug.Print "CallType"
    End If
        
    Debug.Print "TestSort: " & TestSort

End Function
