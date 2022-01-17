Attribute VB_Name = "testStringBuilder"
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
'  Module Name: testStringBuilder
'  Module Description: Unit tests for clsStringBuilder class.
'  Module Version: 1.0
'  Module License: MIT
'  Module Author: Peter Roach; PeterRoach.Code@gmail.com
'  Module Contents:
'      TestStringBuilder
'      TestCapacity
'      TestSize
'      TestReinitialize
'      TestEnsureCapacity
'      TestGrowCapacity
'      TestShrinkCapacity
'      TestTrimToSize
'      TestCharAt
'      TestAppend
'      TestInsert
'      TestRemove
'      TestCharPositionOf
'      TestLastCharPositionOf
'      TestReplace
'      TestReverse
'      TestClear
'      TestSubstring
'      TestToString


Private Const DEFAULT_CAPACITY& = 16


'Example Usage=========================================================
'======================================================================

Public Sub Example()

    Dim SB As clsStringBuilder
    Set SB = New clsStringBuilder

    SB.Append "ABCDABCD"

    Debug.Print SB.CharAt(0)

    Debug.Print SB.CharPositionOf("C")

    Debug.Print SB.LastCharPositionOf("C")

    Debug.Print SB.Substring(0, 3)

    SB.Insert 4, ","

    Debug.Print SB.ToString

    SB.Remove 4, 4

    Debug.Print SB.ToString

    Debug.Print SB.Size, SB.Capacity

    SB.Reverse

    Debug.Print SB.ToString

    Set SB = Nothing

End Sub


'Unit Tests============================================================
'======================================================================

Public Function testStringBuilder() As Boolean

    testStringBuilder = _
        TestCapacity And _
        TestSize And _
        TestReinitialize And _
        TestEnsureCapacity And _
        TestGrowCapacity And _
        TestShrinkCapacity And _
        TestTrimToSize And _
        TestCharAt And _
        TestAppend And _
        TestInsert And _
        TestRemove And _
        TestCharPositionOf And _
        TestLastCharPositionOf And _
        TestReplace And _
        TestReverse And _
        TestClear And _
        TestSubstring And _
        TestToString

    Debug.Print "TestStringBuilder: " & testStringBuilder

End Function

Private Function TestCapacity() As Boolean

    TestCapacity = True
    
    Dim SB As clsStringBuilder
    Set SB = New clsStringBuilder
    
    'Default
    If SB.Capacity <> DEFAULT_CAPACITY Then
        TestCapacity = False
        Debug.Print "Default"
    End If
    
    'Reinitialize
    SB.Append String$(DEFAULT_CAPACITY + 1, "A")
    SB.Reinitialize
    If SB.Capacity <> DEFAULT_CAPACITY Then
        TestCapacity = False
        Debug.Print "Reinitialize"
    End If
    
    'EnsureCapacity
    SB.EnsureCapacity DEFAULT_CAPACITY + 1
    If SB.Capacity <> DEFAULT_CAPACITY * 2 + 2 Then
        TestCapacity = False
        Debug.Print "EnsureCapacity"
    End If
    
    SB.Reinitialize
    
    'GrowCapacity
    SB.GrowCapacity DEFAULT_CAPACITY + 1
    If SB.Capacity <> DEFAULT_CAPACITY + 1 Then
        TestCapacity = False
        Debug.Print "GrowCapacity"
    End If
    
    
    'ShrinkCapacity
    SB.ShrinkCapacity DEFAULT_CAPACITY
    If SB.Capacity <> DEFAULT_CAPACITY Then
        TestCapacity = False
        Debug.Print "ShrinkCapacity"
    End If
    
    'TrimToSize
    SB.Append "A"
    SB.TrimToSize
    If SB.Capacity <> 1 Then
        TestCapacity = False
        Debug.Print "TrimToSize"
    End If
    
    SB.Reinitialize
    
    'CharAt
    SB.Append "ABC"
    If SB.CharAt(0) <> "A" Or SB.Capacity <> DEFAULT_CAPACITY Then
        TestCapacity = False
        Debug.Print "CharAt"
    End If
    
    SB.Reinitialize
    
    'Append
    SB.Append "A"
    If SB.Capacity <> DEFAULT_CAPACITY Then
        TestCapacity = False
        Debug.Print "Append"
    End If
    SB.Append String$(16, "A")
    If SB.Capacity <> DEFAULT_CAPACITY * 2 + 2 Then
        TestCapacity = False
        Debug.Print "Append"
    End If
    
    SB.Reinitialize
    
    'Insert
    SB.Insert 0, "A"
    If SB.Capacity <> DEFAULT_CAPACITY Then
        TestCapacity = False
        Debug.Print "Insert"
    End If
    SB.Insert 0, String$(16, "A")
    If SB.Capacity <> DEFAULT_CAPACITY * 2 + 2 Then
        TestCapacity = False
        Debug.Print "Insert"
    End If
    
    SB.Reinitialize
    
    'Remove
    SB.Append "ABCDEFGHIJKLMNOPQ"
    SB.Remove 0, 16
    If SB.Capacity <> DEFAULT_CAPACITY * 2 + 2 Then
        TestCapacity = False
        Debug.Print "Remove"
    End If
    
    SB.Reinitialize
    SB.Append "ABC"
    
    'CharPositionOf
    If SB.CharPositionOf("A") <> 0 Or SB.Capacity <> DEFAULT_CAPACITY Then
        TestCapacity = False
        Debug.Print "CharPositionOf"
    End If
    
    'LastCharPositionOf
    If SB.LastCharPositionOf("A") <> 0 Or SB.Capacity <> DEFAULT_CAPACITY Then
        TestCapacity = False
        Debug.Print "LastCharPositionOf"
    End If
    
    SB.Reinitialize
    
    'Replace
    SB.Append "ABC"
    SB.Replace 0, 2, String$(17, "A")
    If SB.Capacity <> DEFAULT_CAPACITY * 2 + 2 Then
        TestCapacity = False
        Debug.Print "Replace"
    End If
    
    SB.Reinitialize
    SB.Append "ABC"
    
    'Reverse
    SB.Reverse
    If SB.Capacity <> DEFAULT_CAPACITY Then
        TestCapacity = False
        Debug.Print "Reverse"
    End If
    
    'Clear
    SB.Clear
    If SB.Capacity <> DEFAULT_CAPACITY Then
        TestCapacity = False
        Debug.Print "Clear"
    End If
    SB.Append String$(17, "A")
    SB.Clear
    If SB.Capacity <> DEFAULT_CAPACITY * 2 + 2 Then
        TestCapacity = False
        Debug.Print "Clear"
    End If
    
    SB.Reinitialize
    SB.Append "ABC"
    
    'Substring
    If SB.Substring(0, 1) <> "AB" Or SB.Capacity <> DEFAULT_CAPACITY Then
        TestCapacity = False
        Debug.Print "Substring"
    End If
    
    'ToString
    If SB.ToString <> "ABC" Or SB.Capacity <> DEFAULT_CAPACITY Then
        TestCapacity = False
        Debug.Print "ToString"
    End If

    Debug.Print "TestCapacity: " & TestCapacity

End Function

Private Function TestSize() As Boolean

    TestSize = True
    
    Dim SB As clsStringBuilder
    Set SB = New clsStringBuilder
    
    'Default
    If SB.Size <> 0 Then
        TestSize = False
        Debug.Print "Default"
    End If
    
    'Reinitialize
    SB.Reinitialize
    If SB.Size <> 0 Then
        TestSize = False
        Debug.Print "Reinitialize"
    End If
    SB.Append "ABC"
    SB.Reinitialize
    If SB.Size <> 0 Then
        TestSize = False
        Debug.Print "Reinitialize"
    End If
    
    'EnsureCapacity
    SB.EnsureCapacity SB.Capacity + 1
    If SB.Size <> 0 Then
        TestSize = False
        Debug.Print "EnsureCapacity"
    End If

    'GrowCapacity
    SB.GrowCapacity SB.Capacity + 1
    If SB.Size <> 0 Then
        TestSize = False
        Debug.Print "GrowCapacity"
    End If

    'ShrinkCapacity
    SB.ShrinkCapacity SB.Capacity - 1
    If SB.Size <> 0 Then
        TestSize = False
        Debug.Print "ShrinkCapacity"
    End If

    SB.Reinitialize

    'TrimToSize
    SB.Append "ABC"
    SB.TrimToSize
    If SB.Size <> 3 Then
        TestSize = False
        Debug.Print "TrimToSize"
    End If

    'CharAt
    If SB.CharAt(0) <> "A" Or SB.Size <> 3 Then
        TestSize = False
        Debug.Print "CharAt"
    End If

    SB.Reinitialize

    'Append
    SB.Append "A"
    If SB.Size <> 1 Then
        TestSize = False
        Debug.Print "Append"
    End If
    SB.Append "BC"
    If SB.Size <> 3 Then
        TestSize = False
        Debug.Print "Append"
    End If

    SB.Reinitialize

    'Insert
    SB.Insert 0, "A"
    If SB.Size <> 1 Then
        TestSize = False
        Debug.Print "Insert"
    End If
    SB.Insert 0, "BC"
    If SB.Size <> 3 Then
        TestSize = False
        Debug.Print "Insert"
    End If

    'Remove
    SB.Remove 0, 2
    If SB.Size <> 0 Then
        TestSize = False
        Debug.Print "Remove"
    End If

    SB.Reinitialize
    SB.Append "ABC"

    'CharPositionOf
    If SB.CharPositionOf("A") <> 0 Or SB.Size <> 3 Then
        TestSize = False
        Debug.Print "CharPositionOf"
    End If

    'LastCharPositionOf
    If SB.LastCharPositionOf("A") <> 0 Or SB.Size <> 3 Then
        TestSize = False
        Debug.Print "LastCharPositionOf"
    End If

    'Replace
    SB.Replace 0, 2, String$(17, "A")
    If SB.Size <> 17 Then
        TestSize = False
        Debug.Print "Replace"
    End If

    SB.Reinitialize
    SB.Append "ABC"

    'Reverse
    SB.Reverse
    If SB.Size <> 3 Then
        TestSize = False
        Debug.Print "Reverse"
    End If

    'Clear
    SB.Clear
    If SB.Size <> 0 Then
        TestSize = False
        Debug.Print "Clear"
    End If

    SB.Append "ABC"

    'Substring
    If SB.Substring(0, 1) <> "AB" Or SB.Size <> 3 Then
        TestSize = False
        Debug.Print "Substring"
    End If

    'ToString
    If SB.ToString <> "ABC" Or SB.Size <> 3 Then
        TestSize = False
        Debug.Print "ToString"
    End If

    Debug.Print "TestSize: " & TestSize

End Function

Private Function TestReinitialize() As Boolean

    TestReinitialize = True

    Dim SB As clsStringBuilder
    Set SB = New clsStringBuilder

    'Default
    SB.Reinitialize
    If SB.Capacity <> DEFAULT_CAPACITY Or SB.Size <> 0 Then
        TestReinitialize = False
        Debug.Print "Default"
    End If

    'More
    SB.Reinitialize 32
    If SB.Capacity <> 32 Or SB.Size <> 0 Then
        TestReinitialize = False
        Debug.Print "More"
    End If

    'Less
    SB.Reinitialize 8
    If SB.Capacity <> 8 Or SB.Size <> 0 Then
        TestReinitialize = False
        Debug.Print "Less"
    End If

    'Size
    SB.Append "ABC"
    SB.Reinitialize
    If SB.Capacity <> DEFAULT_CAPACITY Or SB.Size <> 0 Then
        TestReinitialize = False
        Debug.Print "Size"
    End If

    'Invalid 0
    On Error Resume Next
    SB.Reinitialize 0
    If Err.Number <> 5 Then
        TestReinitialize = False
        Debug.Print "Invalid 0"
    End If
    On Error GoTo 0

    'Invalid Negative
    On Error Resume Next
    SB.Reinitialize 0
    If Err.Number <> 5 Then
        TestReinitialize = False
        Debug.Print "Invalid Negative"
    End If
    On Error GoTo 0

    Debug.Print "TestReinitialize: " & TestReinitialize

End Function

Private Function TestEnsureCapacity() As Boolean

    TestEnsureCapacity = True

    Dim SB As clsStringBuilder
    Set SB = New clsStringBuilder

    'Less
    SB.EnsureCapacity SB.Capacity - 1
    If SB.Capacity <> DEFAULT_CAPACITY Then
        TestEnsureCapacity = False
        Debug.Print "Less"
    End If
    
    'Same
    SB.EnsureCapacity SB.Capacity
    If SB.Capacity <> DEFAULT_CAPACITY Then
        TestEnsureCapacity = False
        Debug.Print "Same"
    End If
    
    'More
    SB.EnsureCapacity SB.Capacity + 1
    If SB.Capacity <> DEFAULT_CAPACITY * 2 + 2 Then
        TestEnsureCapacity = False
        Debug.Print "More"
    End If
    
    SB.Reinitialize
    
    'More than double
    SB.EnsureCapacity SB.Capacity * 2 + 3
    If SB.Capacity <> DEFAULT_CAPACITY * 2 + 3 Then
        TestEnsureCapacity = False
        Debug.Print "More than double"
    End If
    
    SB.Reinitialize
    
    'Invalid 0 - implicitly not possible
    SB.EnsureCapacity 0
    If SB.Capacity <> DEFAULT_CAPACITY Then
        TestEnsureCapacity = False
        Debug.Print "Invalid 0"
    End If
    
    'Invalid Negative - implicitly not possible
    SB.EnsureCapacity -1
    If SB.Capacity <> DEFAULT_CAPACITY Then
        TestEnsureCapacity = False
        Debug.Print "Invalid Negative"
    End If
    
    Debug.Print "TestEnsureCapacity: " & TestEnsureCapacity

End Function

Private Function TestGrowCapacity() As Boolean

    TestGrowCapacity = True

    Dim SB As clsStringBuilder
    Set SB = New clsStringBuilder

    'Less
    SB.GrowCapacity SB.Capacity - 1
    If SB.Capacity <> DEFAULT_CAPACITY Then
        TestGrowCapacity = False
        Debug.Print "Less"
    End If
    
    'Same
    SB.GrowCapacity SB.Capacity
    If SB.Capacity <> DEFAULT_CAPACITY Then
        TestGrowCapacity = False
        Debug.Print "Same"
    End If
    
    'More
    SB.GrowCapacity SB.Capacity + 1
    If SB.Capacity <> DEFAULT_CAPACITY + 1 Then
        TestGrowCapacity = False
        Debug.Print "More"
    End If
    
    SB.Reinitialize
    
    'Invalid 0 - implicitly not possible
    SB.GrowCapacity 0
    If SB.Capacity <> DEFAULT_CAPACITY Then
        TestGrowCapacity = False
        Debug.Print "Invalid 0"
    End If
    
    'Invalid Negative - implicitly not possible
    SB.GrowCapacity -1
    If SB.Capacity <> DEFAULT_CAPACITY Then
        TestGrowCapacity = False
        Debug.Print "Invalid Negative"
    End If
    
    Debug.Print "TestGrowCapacity: " & TestGrowCapacity

End Function

Private Function TestShrinkCapacity() As Boolean

    TestShrinkCapacity = True

    Dim SB As clsStringBuilder
    Set SB = New clsStringBuilder

    'More
    SB.ShrinkCapacity SB.Capacity + 1
    If SB.Capacity <> DEFAULT_CAPACITY Then
        TestShrinkCapacity = False
        Debug.Print "More"
    End If
    
    'Same
    SB.ShrinkCapacity SB.Capacity
    If SB.Capacity <> DEFAULT_CAPACITY Then
        TestShrinkCapacity = False
        Debug.Print "Same"
    End If
    
    'Less
    SB.ShrinkCapacity SB.Capacity - 1
    If SB.Capacity <> DEFAULT_CAPACITY - 1 Then
        TestShrinkCapacity = False
        Debug.Print "Less"
    End If
    
    'Less than size
    SB.Append "ABC"
    SB.ShrinkCapacity 2
    If SB.Capacity <> 3 Then
        TestShrinkCapacity = False
        Debug.Print "Less than size"
    End If
    
    SB.Reinitialize
    
    'Invalid 0
    On Error Resume Next
    SB.ShrinkCapacity 0
    If Err.Number <> 5 Then
        TestShrinkCapacity = False
        Debug.Print "Invalid 0"
    End If
    On Error GoTo 0
    
    'Invalid Negative
    On Error Resume Next
    SB.ShrinkCapacity -1
    If Err.Number <> 5 Then
        TestShrinkCapacity = False
        Debug.Print "Invalid Negative"
    End If
    On Error GoTo 0
    
    Debug.Print "TestShrinkCapacity: " & TestShrinkCapacity
    
End Function

Private Function TestTrimToSize() As Boolean

    TestTrimToSize = True

    Dim SB As clsStringBuilder
    Set SB = New clsStringBuilder

    'Size 0
    SB.TrimToSize
    If SB.Capacity <> DEFAULT_CAPACITY Then
        TestTrimToSize = False
        Debug.Print "Size 0"
    End If
    
    'Size > 0
    SB.Append "ABC"
    SB.TrimToSize
    If SB.Capacity <> 3 Then
        TestTrimToSize = False
        Debug.Print "Size > 0"
    End If
    
    Debug.Print "TestTrimToSize: " & TestTrimToSize

End Function

Private Function TestCharAt() As Boolean

    TestCharAt = True

    Dim SB As clsStringBuilder
    Set SB = New clsStringBuilder

    'Empty
    On Error Resume Next
    Debug.Print SB.CharAt(0)
    If Err.Number <> 9 Then
        TestCharAt = False
        Debug.Print "Empty"
    End If
    On Error GoTo 0
    
    SB.Append "ABC"
    
    'First
    If SB.CharAt(0) <> "A" Then
        TestCharAt = False
        Debug.Print "First"
    End If
    
    'Middle
    If SB.CharAt(1) <> "B" Then
        TestCharAt = False
        Debug.Print "Middle"
    End If
    
    'Last
    If SB.CharAt(2) <> "C" Then
        TestCharAt = False
        Debug.Print "Last"
    End If
    
    'Invalid lower bound
    On Error Resume Next
    Debug.Print SB.CharAt(-1)
    If Err.Number <> 9 Then
        TestCharAt = False
        Debug.Print "Invalid lower bound"
    End If
    On Error GoTo 0
    
    'Invalid upper bound
    On Error Resume Next
    Debug.Print SB.CharAt(3)
    If Err.Number <> 9 Then
        TestCharAt = False
        Debug.Print "Invalid upper bound"
    End If
    On Error GoTo 0
    
    Debug.Print "TestCharAt: " & TestCharAt

End Function

Private Function TestAppend() As Boolean

    TestAppend = True

    Dim SB As clsStringBuilder
    Set SB = New clsStringBuilder

    'Empty string
    SB.Append ""
    If SB.ToString <> "" Or SB.Size <> 0 Then
        TestAppend = False
        Debug.Print "Empty string"
    End If
    
    'One char
    SB.Append "A"
    If SB.ToString <> "A" Or SB.Size <> 1 Then
        TestAppend = False
        Debug.Print "One char"
    End If
    
    'Multiple char
    SB.Append "BC"
    If SB.ToString <> "ABC" Or SB.Size <> 3 Then
        TestAppend = False
        Debug.Print "Multiple char"
    End If
    
    Debug.Print "TestAppend: " & TestAppend

End Function

Private Function TestInsert() As Boolean

    TestInsert = True
    
    Dim SB As clsStringBuilder
    Set SB = New clsStringBuilder
    
    'Empty string
    SB.Insert 0, ""
    If SB.ToString <> "" Or SB.Size <> 0 Then
        TestInsert = False
        Debug.Print "Empty string"
    End If
    
    'One char
    SB.Insert 0, "A"
    If SB.ToString <> "A" Or SB.Size <> 1 Then
        TestInsert = False
        Debug.Print "One char"
    End If
    
    SB.Reinitialize
    SB.Append "ABC"
    
    'One First
    SB.Insert 0, "X"
    If SB.ToString <> "XABC" Or SB.Size <> 4 Then
        TestInsert = False
        Debug.Print "One First"
    End If
    
    'One Middle
    SB.Insert 2, "Y"
    If SB.ToString <> "XAYBC" Or SB.Size <> 5 Then
        TestInsert = False
        Debug.Print "One Middle"
    End If
    
    'One Last
    SB.Insert SB.Size, "Z"
    If SB.ToString <> "XAYBCZ" Or SB.Size <> 6 Then
        TestInsert = False
        Debug.Print "One Last"
    End If
    
    SB.Reinitialize
    
    'Multiple chars
    SB.Insert 0, "ABC"
    If SB.ToString <> "ABC" Or SB.Size <> 3 Then
        TestInsert = False
        Debug.Print "Multiple chars"
    End If
    
    'Multiple First
    SB.Insert 0, "LM"
    If SB.ToString <> "LMABC" Or SB.Size <> 5 Then
        TestInsert = False
        Debug.Print "Multiple First"
    End If
    
    'Multiple Middle
    SB.Insert 3, "QR"
    If SB.ToString <> "LMAQRBC" Or SB.Size <> 7 Then
        TestInsert = False
        Debug.Print "Multiple Middle"
    End If
    
    'Multiple Last
    SB.Insert SB.Size, "YZ"
    If SB.ToString <> "LMAQRBCYZ" Or SB.Size <> 9 Then
        TestInsert = False
        Debug.Print "Multiple Last"
    End If
    
    'Invalid lower bound
    On Error Resume Next
    SB.Insert -1, "A"
    If Err.Number <> 9 Then
        TestInsert = False
        Debug.Print "Invalid lower bound"
    End If
    On Error GoTo 0
    
    'Invalid upper bound
    On Error Resume Next
    SB.Insert SB.Size + 1, "A"
    If Err.Number <> 9 Then
        TestInsert = False
        Debug.Print "Invalid upper bound"
    End If
    On Error GoTo 0
    
    Debug.Print "TestInsert: " & TestInsert

End Function

Private Function TestRemove() As Boolean

    TestRemove = True
    
    Dim SB As clsStringBuilder
    Set SB = New clsStringBuilder
    
    'Empty
    On Error Resume Next
    SB.Remove 0, 0
    If Err.Number <> 9 Then
        TestRemove = False
        Debug.Print "Empty"
    End If
    On Error GoTo 0
    
    'One
    SB.Append "A"
    SB.Remove 0, 0
    If SB.ToString <> "" Or SB.Size <> 0 Then
        TestRemove = False
        Debug.Print "One"
    End If
    
    'Many Start
    SB.Reinitialize
    SB.Append "ABCD"
    SB.Remove 0, 1
    If SB.ToString <> "CD" Or SB.Size <> 2 Then
        TestRemove = False
        Debug.Print "Many Start"
    End If
    'Many Middle
    SB.Reinitialize
    SB.Append "ABCD"
    SB.Remove 1, 2
    If SB.ToString <> "AD" Or SB.Size <> 2 Then
        TestRemove = False
        Debug.Print "Many Middle"
    End If
    'Many End
    SB.Reinitialize
    SB.Append "ABCD"
    SB.Remove SB.Size - 2, SB.Size - 1
    If SB.ToString <> "AB" Or SB.Size <> 2 Then
        TestRemove = False
        Debug.Print "Many End"
    End If

    SB.Reinitialize
    SB.Append "ABC"
    
    'Invalid lower > upper
    On Error Resume Next
    SB.Remove 1, 0
    If Err.Number <> 5 Then
        TestRemove = False
        Debug.Print "Invalid lower > upper"
    End If
    On Error GoTo 0
    
    'Invalid lower bound
    On Error Resume Next
    SB.Remove -1, 0
    If Err.Number <> 9 Then
        TestRemove = False
        Debug.Print "Invalid lower bound"
    End If
    On Error GoTo 0
    
    'Invalid upper bound
    On Error Resume Next
    SB.Remove 0, 3
    If Err.Number <> 9 Then
        TestRemove = False
        Debug.Print "Invalid upper bound"
    End If
    On Error GoTo 0
    
    Debug.Print "TestRemove: " & TestRemove

End Function

Private Function TestCharPositionOf() As Boolean

    TestCharPositionOf = True
    
    Dim SB As clsStringBuilder
    Set SB = New clsStringBuilder
    
    'Empty
    If SB.CharPositionOf("") <> -1 Then
        TestCharPositionOf = False
        Debug.Print "Empty"
    End If
    
    'One
    SB.Append "A"
    If SB.CharPositionOf("") <> -1 Then
        TestCharPositionOf = False
        Debug.Print "One"
    End If
    If SB.CharPositionOf("A") <> 0 Then
        TestCharPositionOf = False
        Debug.Print "One"
    End If
    
    SB.Reinitialize
    SB.Append "ABC"
    
    'One First
    If SB.CharPositionOf("A") <> 0 Then
        TestCharPositionOf = False
        Debug.Print "One First"
    End If
    
    'One Middle
    If SB.CharPositionOf("B") <> 1 Then
        TestCharPositionOf = False
        Debug.Print "One Middle"
    End If
    
    'One Last
    If SB.CharPositionOf("C") <> 2 Then
        TestCharPositionOf = False
        Debug.Print "One Last"
    End If
    
    SB.Reinitialize
    SB.Append "ABCD"

    'Multiple First
    If SB.CharPositionOf("AB") <> 0 Then
        TestCharPositionOf = False
        Debug.Print "Multiple First"
    End If
    
    'Multiple Middle
    If SB.CharPositionOf("BC") <> 1 Then
        TestCharPositionOf = False
        Debug.Print "Multiple Middle"
    End If
    
    'Multiple Last
    If SB.CharPositionOf("CD") <> 2 Then
        TestCharPositionOf = False
        Debug.Print "Multiple Last"
    End If
    
    SB.Reinitialize
    SB.Append "ABA"
    
    'Duplicate
    If SB.CharPositionOf("A") <> 0 Then
        TestCharPositionOf = False
        Debug.Print "Duplicate"
    End If
        
    'From
    If SB.CharPositionOf("A", 1) <> 2 Then
        TestCharPositionOf = False
        Debug.Print "From"
    End If
    
    'From Negative
    If SB.CharPositionOf("A", -1) <> 2 Then
        TestCharPositionOf = False
        Debug.Print "From Negative"
    End If
    
    'Invalid From upper bound
    On Error Resume Next
    Debug.Print SB.CharPositionOf("A", 3)
    If Err.Number <> 9 Then
        TestCharPositionOf = False
        Debug.Print "Invalid From upper bound"
    End If
    On Error GoTo 0
    
    'Invalid From lower bound
    On Error Resume Next
    Debug.Print SB.CharPositionOf("A", -4)
    If Err.Number <> 9 Then
        TestCharPositionOf = False
        Debug.Print "Invalid From lower bound"
    End If
    On Error GoTo 0
    
    'Not there
    If SB.CharPositionOf("Z") <> -1 Then
        TestCharPositionOf = False
        Debug.Print "Not there"
    End If
    
    Debug.Print "TestCharPositionOf: " & TestCharPositionOf

End Function

Private Function TestLastCharPositionOf() As Boolean

    TestLastCharPositionOf = True
    
    Dim SB As clsStringBuilder
    Set SB = New clsStringBuilder
    
    'Empty
    If SB.LastCharPositionOf("") <> -1 Then
        TestLastCharPositionOf = False
        Debug.Print "Empty"
    End If
    
    'One
    SB.Append "A"
    If SB.LastCharPositionOf("") <> -1 Then
        TestLastCharPositionOf = False
        Debug.Print "One"
    End If
    If SB.LastCharPositionOf("A") <> 0 Then
        TestLastCharPositionOf = False
        Debug.Print "One"
    End If
    
    SB.Reinitialize
    SB.Append "ABC"
    
    'One First
    If SB.LastCharPositionOf("A") <> 0 Then
        TestLastCharPositionOf = False
        Debug.Print "One First"
    End If
    
    'One Middle
    If SB.LastCharPositionOf("B") <> 1 Then
        TestLastCharPositionOf = False
        Debug.Print "One Middle"
    End If
    
    'One Last
    If SB.LastCharPositionOf("C") <> 2 Then
        TestLastCharPositionOf = False
        Debug.Print "One Last"
    End If
    
    SB.Reinitialize
    SB.Append "ABCD"

    'Multiple First
    If SB.LastCharPositionOf("AB") <> 0 Then
        TestLastCharPositionOf = False
        Debug.Print "Multiple First"
    End If
    
    'Multiple Middle
    If SB.LastCharPositionOf("BC") <> 1 Then
        TestLastCharPositionOf = False
        Debug.Print "Multiple Middle"
    End If
    
    'Multiple Last
    If SB.LastCharPositionOf("CD") <> 2 Then
        TestLastCharPositionOf = False
        Debug.Print "Multiple Last"
    End If
    
    SB.Reinitialize
    SB.Append "ABA"
    
    'Duplicate
    If SB.LastCharPositionOf("A") <> 2 Then
        TestLastCharPositionOf = False
        Debug.Print "Duplicate"
    End If
    
    'From
    If SB.LastCharPositionOf("A", 1) <> 0 Then
        TestLastCharPositionOf = False
        Debug.Print "From"
    End If
    
    'From Negative
    If SB.LastCharPositionOf("A", -1) <> 2 Then
        TestLastCharPositionOf = False
        Debug.Print "From Negative"
    End If
    If SB.LastCharPositionOf("A", -2) <> 0 Then
        TestLastCharPositionOf = False
        Debug.Print "From Negative"
    End If
    
    'Invalid From upper bound
    On Error Resume Next
    Debug.Print SB.LastCharPositionOf("A", 3)
    If Err.Number <> 9 Then
        TestLastCharPositionOf = False
        Debug.Print "Invalid From upper bound"
    End If
    On Error GoTo 0
    
    'Invalid From lower bound
    On Error Resume Next
    Debug.Print SB.LastCharPositionOf("A", -4)
    If Err.Number <> 9 Then
        TestLastCharPositionOf = False
        Debug.Print "Invalid From lower bound"
    End If
    On Error GoTo 0
    
    'Not there
    If SB.LastCharPositionOf("Z") <> -1 Then
        TestLastCharPositionOf = False
        Debug.Print "Not there"
    End If
    
    Debug.Print "TestLastCharPositionOf: " & TestLastCharPositionOf

End Function

Private Function TestReplace() As Boolean

    TestReplace = True
    
    Dim SB As clsStringBuilder
    Set SB = New clsStringBuilder
    
    'Empty
    On Error Resume Next
    SB.Replace 0, 0, "A"
    If Err.Number <> 9 Then
        TestReplace = False
        Debug.Print "Empty"
    End If
    On Error GoTo 0
    
    'One
    SB.Append "A"
    SB.Replace 0, 0, "Z"
    If SB.ToString <> "Z" Or SB.Size <> 1 Then
        TestReplace = False
        Debug.Print "One"
    End If
    
    'Many
    SB.Append "YX"
    SB.Replace 0, 2, "ABC"
    If SB.ToString <> "ABC" Or SB.Size <> 3 Then
        TestReplace = False
        Debug.Print "Many"
    End If
    
    'Replacement empty
    SB.Reinitialize
    SB.Append "ABC"
    SB.Replace 0, 2, ""
    If SB.ToString <> "" Or SB.Size <> 0 Then
        TestReplace = False
        Debug.Print "Replacement Empty"
    End If
    
    'Replacement shorter
    SB.Reinitialize
    SB.Append "ABC"
    SB.Replace 0, 2, "XY"
    If SB.ToString <> "XY" Or SB.Size <> 2 Then
        TestReplace = False
        Debug.Print "Replacement shorter"
    End If
    
    'Replacement longer
    SB.Reinitialize
    SB.Append "ABC"
    SB.Replace 0, 2, "LMNOP"
    If SB.ToString <> "LMNOP" Or SB.Size <> 5 Then
        TestReplace = False
        Debug.Print "Replacement longer"
    End If
    
    SB.Reinitialize
    SB.Append "ABC"
    
    'Invalid lower > upper
    On Error Resume Next
    SB.Replace 1, 0, ""
    If Err.Number <> 5 Then
        TestReplace = False
        Debug.Print "Invalid lower > upper"
    End If
    On Error GoTo 0
    
    'Invalid lower
    On Error Resume Next
    SB.Replace -1, 0, ""
    If Err.Number <> 9 Then
        TestReplace = False
        Debug.Print "Invalid lower"
    End If
    On Error GoTo 0
    
    'Invalid upper
    On Error Resume Next
    SB.Replace 0, 3, ""
    If Err.Number <> 9 Then
        TestReplace = False
        Debug.Print "Invalid upper"
    End If
    On Error GoTo 0
    
    Debug.Print "TestReplace: " & TestReplace

End Function

Private Function TestReverse() As Boolean

    TestReverse = True
    
    Dim SB As clsStringBuilder
    Set SB = New clsStringBuilder
    
    'Empty
    SB.Reverse
    If SB.ToString <> "" Or SB.Size <> 0 Then
        TestReverse = False
        Debug.Print "Empty"
    End If
    
    'One
    SB.Append "A"
    SB.Reverse
    If SB.ToString <> "A" Or SB.Size <> 1 Then
        TestReverse = False
        Debug.Print "One"
    End If
    
    'Odd
    SB.Append "BC"
    SB.Reverse
    If SB.ToString <> "CBA" Or SB.Size <> 3 Then
        TestReverse = False
        Debug.Print "Odd"
    End If
    
    'Even
    SB.Insert 0, "D"
    SB.Reverse
    If SB.ToString <> "ABCD" Or SB.Size <> 4 Then
        TestReverse = False
        Debug.Print "Even"
    End If
    
    Debug.Print "TestReverse: " & TestReverse

End Function

Private Function TestClear() As Boolean

    TestClear = True
    
    Dim SB As clsStringBuilder
    Set SB = New clsStringBuilder
    
    'Empty
    SB.Clear
    If SB.Capacity <> DEFAULT_CAPACITY Or SB.Size <> 0 Or SB.ToString <> "" Then
        TestClear = False
        Debug.Print "Empty"
    End If
    
    'One
    SB.Append "A"
    SB.Clear
    If SB.Capacity <> DEFAULT_CAPACITY Or SB.Size <> 0 Or SB.ToString <> "" Then
        TestClear = False
        Debug.Print "One"
    End If
    
    'Many
    SB.Append "ABC"
    SB.Clear
    If SB.Capacity <> DEFAULT_CAPACITY Or SB.Size <> 0 Or SB.ToString <> "" Then
        TestClear = False
        Debug.Print "Many"
    End If
    
    'Capacity holds
    SB.GrowCapacity DEFAULT_CAPACITY + 1
    SB.Clear
    If SB.Capacity <> DEFAULT_CAPACITY + 1 Then
        TestClear = False
        Debug.Print "Capacity holds"
    End If
    
    Debug.Print "TestClear: " & TestClear

End Function

Private Function TestSubstring() As Boolean

    TestSubstring = True
    
    Dim SB As clsStringBuilder
    Set SB = New clsStringBuilder
    
    'Empty
    On Error Resume Next
    SB.Substring 0, 0
    If Err.Number <> 9 Then
        TestSubstring = False
        Debug.Print "Empty"
    End If
    On Error GoTo 0
    
    'One
    SB.Append "A"
    If SB.Substring(0, 0) <> "A" Then
        TestSubstring = False
        Debug.Print "One"
    End If
    
    SB.Append "BCD"
    
    'First
    If SB.Substring(0, 1) <> "AB" Then
        TestSubstring = False
        Debug.Print "First"
    End If
    
    'Middle
    If SB.Substring(1, 2) <> "BC" Then
        TestSubstring = False
        Debug.Print "Middle"
    End If
    
    'End
    If SB.Substring(2, 3) <> "CD" Then
        TestSubstring = False
        Debug.Print "End"
    End If
    
    'Invalid lower > upper
    On Error Resume Next
    SB.Substring 1, 0
    If Err.Number <> 5 Then
        TestSubstring = False
        Debug.Print "Invalid lower > upper"
    End If
    On Error GoTo 0
    
    'Invalid lower bound
    On Error Resume Next
    SB.Substring -1, 0
    If Err.Number <> 9 Then
        TestSubstring = False
        Debug.Print "Invalid lower bound"
    End If
    On Error GoTo 0
    
    'Invalid upper bound
    On Error Resume Next
    SB.Substring 0, 4
    If Err.Number <> 9 Then
        TestSubstring = False
        Debug.Print "Invalid upper bound"
    End If
    On Error GoTo 0
    
    Debug.Print "TestSubstring: " & TestSubstring

End Function

Private Function TestToString() As Boolean

    TestToString = True
    
    Dim SB As clsStringBuilder
    Set SB = New clsStringBuilder
    
    'Empty
    If SB.ToString <> "" Then
        TestToString = False
        Debug.Print "Empty"
    End If
    
    'One
    SB.Append "A"
    If SB.ToString <> "A" Then
        TestToString = False
        Debug.Print "One"
    End If
    
    'Many
    SB.Append "BC"
    If SB.ToString <> "ABC" Then
        TestToString = False
        Debug.Print "Many"
    End If
    
    Debug.Print "TestToString: " & TestToString

End Function
