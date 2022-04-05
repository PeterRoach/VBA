VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsArrayListString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

'  Module Type: Class
'  Module Name: clsArrayListString
'  Module Description: ArrayList implementation for Strings.
'  Module Version: 1.1
'  Module License: MIT
'  Module Author: Peter Roach; PeterRoach.Code@gmail.com
'  Module Contents:
'  ----------------------------------------
'    Public Properties:
'       Capacity
'       Size
'       Item
'    Public Methods:
'       GrowCapacity
'       ShrinkCapacity
'       EnsureCapacity
'       TrimToSize
'       Reinitialize
'       Append
'       AppendArray
'       AppendArrayList
'       Insert
'       InsertArray
'       InsertArrayList
'       Remove
'       RemoveRange
'       RemoveFirst
'       RemoveLast
'       RemoveAll
'       ReplaceAll
'       IndexOf
'       LastIndexOf
'       Count
'       Contains
'       Clear
'       Reverse
'       Sort
'       ToStringArray
'       JoinString
'  ----------------------------------------

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'  Example Usage:

'        Public Sub Example()
'
'            Dim AL As clsArrayListString
'            Set AL = New clsArrayListString
'
'            AL.Append "A"
'            AL.Append "B"
'            AL.Append "C"
'
'            Dim Arr$(0 To 2)
'            Arr(0) = "A"
'            Arr(0) = "B"
'            Arr(0) = "C"
'            AL.AppendArray Arr
'
'            Dim AL1 As clsArrayListString
'            Set AL1 = New clsArrayListString
'            AL1.Append "A"
'            AL1.Append "B"
'            AL1.Append "C"
'            AL.AppendArrayList AL1
'
'            AL.Insert 0, "A"
'            AL.InsertArray 0, Arr
'            AL.InsertArrayList 0, AL1
'
'            AL.Remove 0
'            AL.RemoveRange 0, 1
'            AL.RemoveFirst "B"
'            AL.RemoveLast "B"
'            AL.RemoveAll "C"
'            AL.ReplaceAll "D", "Z"
'
'            Debug.Print AL.Contains("A")
'            Debug.Print AL.Count("A")
'            Debug.Print AL.IndexOf("A")
'            Debug.Print AL.LastIndexOf("A")
'
'            AL.Reverse
'            AL.Sort
'
'            AL.Clear
'            AL.Reinitialize
'
'            AL.Append "A"
'            AL.Append "B"
'            AL.Append "C"
'
'            Debug.Print AL.JoinString()
'            Debug.Print AL.JoinString(",")
'
'            Dim Arr1$()
'            Arr1 = AL.ToStringArray()
'
'        End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' ***See testArrayListString module for Unit Tests


'Private===============================================================
'======================================================================

Private Const DEFAULT_CAPACITY& = 10

Private pStringArray$()
Private pSize&

Private Sub Class_Initialize()
    ReDim pStringArray$(0 To DEFAULT_CAPACITY - 1)
    pSize = 0
End Sub


'Public Properties=====================================================
'======================================================================

Public Property Get Capacity&()
    Capacity = UBound(pStringArray) - LBound(pStringArray) + 1
End Property

Public Property Get Size&()
    Size = pSize
End Property

Public Property Get Item$(Index&)
    If Index < 0 Or Index > pSize - 1 Then
        Err.Raise 9
    End If
    Item = pStringArray(Index)
End Property

Public Property Let Item(Index&, Text$)
    If Index < 0 Or Index > pSize - 1 Then
        Err.Raise 9
    End If
    pStringArray(Index) = Text
End Property


'Public Methods========================================================
'======================================================================

Public Sub GrowCapacity(Capacity&)
    If Capacity > Me.Capacity Then
        ReDim Preserve pStringArray$(0 To Capacity - 1)
    End If
End Sub

Public Sub ShrinkCapacity(Capacity&)
    If Capacity < 1 Then
        Err.Raise 5
    End If
    If Capacity < Me.Capacity Then
        If Capacity > pSize Then
            ReDim Preserve pStringArray$(0 To Capacity - 1)
        Else
            ReDim Preserve pStringArray$(0 To pSize - 1)
        End If
    End If
End Sub

Public Sub EnsureCapacity(MinimumCapacity&)
    Dim CurrentCapacity&
    CurrentCapacity = Me.Capacity
    If MinimumCapacity > CurrentCapacity Then
        Dim DoubleCapacity&
        DoubleCapacity = CurrentCapacity * 2 + 2
        If DoubleCapacity > MinimumCapacity Then
            Me.GrowCapacity DoubleCapacity
        Else
            Me.GrowCapacity MinimumCapacity
        End If
    End If
End Sub

Public Sub TrimToSize()
    If pSize > 0 Then
        ReDim Preserve pStringArray$(0 To pSize - 1)
    Else
        Me.Reinitialize
    End If
End Sub

Public Sub Reinitialize(Optional Capacity& = DEFAULT_CAPACITY)
    If Capacity < 1 Then
        Err.Raise 5
    End If
    ReDim pStringArray$(0 To Capacity - 1)
    pSize = 0
End Sub

Public Sub Append(Text$)
    EnsureCapacity pSize + 1
    pStringArray(pSize) = Text
    pSize = pSize + 1
End Sub

Public Sub AppendArray(Arr$())
    EnsureCapacity pSize + UBound(Arr) - LBound(Arr) + 1
    Dim i&
    For i = LBound(Arr) To UBound(Arr)
        pStringArray(pSize) = Arr(i)
        pSize = pSize + 1
    Next i
End Sub

Public Sub AppendArrayList(AL As clsArrayListString)
    EnsureCapacity pSize + AL.Size
    Dim i&
    For i = 0 To AL.Size - 1
        pStringArray(pSize) = AL.Item(i)
        pSize = pSize + 1
    Next i
End Sub

Public Sub Insert(Index&, Text$)
    If Index < 0 Or Index > pSize Then
        Err.Raise 9
    End If
    EnsureCapacity pSize + 1
    Dim i&
    For i = pSize - 1 To Index Step -1
        pStringArray(i + 1) = pStringArray(i)
    Next i
    pStringArray(Index) = Text
    pSize = pSize + 1
End Sub

Public Sub InsertArray(Index&, Arr$())
    If Index < 0 Or Index > pSize Then
        Err.Raise 9
    End If
    Dim ArrSize&
    ArrSize = UBound(Arr) - LBound(Arr) + 1
    EnsureCapacity pSize + ArrSize
    Dim i&
    For i = pSize - 1 To Index Step -1
        pStringArray(i + ArrSize) = pStringArray(i)
    Next i
    Dim j&
    j = LBound(Arr)
    For i = Index To Index + ArrSize - 1
        pStringArray(i) = Arr(j)
        j = j + 1
    Next i
    pSize = pSize + ArrSize
End Sub

Public Sub InsertArrayList(Index&, AL As clsArrayListString)
    If Index < 0 Or Index > pSize Then
        Err.Raise 9
    End If
    EnsureCapacity pSize + AL.Size
    Dim ALSize&
    ALSize = AL.Size
    Dim i&
    For i = pSize - 1 To Index Step -1
        pStringArray(i + ALSize) = pStringArray(i)
    Next i
    Dim j&
    j = 0
    For i = Index To Index + ALSize - 1
        pStringArray(i) = AL.Item(j)
        j = j + 1
    Next i
    pSize = pSize + ALSize
End Sub

Public Sub Remove(Index&)
    If Index < 0 Or Index > pSize - 1 Then
        Err.Raise 9
    End If
    Dim i&
    For i = Index To pSize - 2
        pStringArray(i) = pStringArray(i + 1)
    Next i
    pStringArray(pSize - 1) = vbNullString
    pSize = pSize - 1
End Sub

Public Sub RemoveRange(LowerIndex&, UpperIndex&)
    If LowerIndex > UpperIndex Then
        Err.Raise 5
    End If
    If LowerIndex < 0 Or LowerIndex > pSize - 1 Then
        Err.Raise 9
    End If
    If UpperIndex < 0 Or UpperIndex > pSize - 1 Then
        Err.Raise 9
    End If
    Dim i&
    Dim j&
    j = LowerIndex
    For i = UpperIndex + 1 To pSize - 1
        pStringArray(j) = pStringArray(i)
        j = j + 1
    Next i
    For i = j To pSize - 1
        pStringArray(i) = vbNullString
    Next i
    pSize = pSize - (UpperIndex - LowerIndex + 1)
End Sub

Public Sub RemoveFirst(Text$, _
Optional CompareMethod As VbCompareMethod = vbBinaryCompare)
    Dim i&
    Dim j&
    For i = 0 To pSize - 1
        If StrComp(pStringArray(i), Text, CompareMethod) = 0 Then
            For j = i To pSize - 2
                pStringArray(j) = pStringArray(j + 1)
            Next j
            pStringArray(pSize - 1) = vbNullString
            pSize = pSize - 1
            Exit Sub
        End If
    Next i
End Sub

Public Sub RemoveLast(Text$, _
Optional CompareMethod As VbCompareMethod = vbBinaryCompare)
    Dim i&
    Dim j&
    For i = pSize - 1 To 0 Step -1
        If StrComp(pStringArray(i), Text, CompareMethod) = 0 Then
            For j = i To pSize - 2
                pStringArray(j) = pStringArray(j + 1)
            Next j
            pStringArray(pSize - 1) = vbNullString
            pSize = pSize - 1
            Exit Sub
        End If
    Next i
End Sub

Public Sub RemoveAll(Text$, _
Optional CompareMethod As VbCompareMethod = vbBinaryCompare)
    Dim i&
    Dim j&
    Dim Offset&
    i = 0
    Offset = 0
    Do While i < pSize
        If StrComp(pStringArray(i), Text, CompareMethod) = 0 Then
            Offset = Offset + 1
        Else
            If Offset > 0 Then
                pStringArray(i - Offset) = pStringArray(i)
            End If
        End If
        i = i + 1
    Loop
    If Offset > 0 Then
        For i = pSize - Offset To pSize - 1
            pStringArray(i) = vbNullString
        Next i
        pSize = pSize - Offset
    End If
End Sub

Public Sub ReplaceAll(Text$, ReplacementText$, _
Optional CompareMethod As VbCompareMethod = vbBinaryCompare)
    Dim i&
    For i = 0 To pSize - 1
        If StrComp(pStringArray(i), Text, CompareMethod) = 0 Then
            pStringArray(i) = ReplacementText
        End If
    Next i
End Sub

Public Function IndexOf&(Text$, Optional FromIndex& = 0, _
Optional CompareMethod As VbCompareMethod = vbBinaryCompare)
    IndexOf = -1
    If Not pSize > 0 Then
        Exit Function
    End If
    Dim StartIndex&
    If FromIndex < 0 Then
        StartIndex = pSize + FromIndex
    Else
        StartIndex = FromIndex
    End If
    If StartIndex < 0 Or StartIndex > pSize - 1 Then
        Err.Raise 9
    End If
    Dim i&
    For i = StartIndex To pSize - 1
        If StrComp(pStringArray(i), Text, CompareMethod) = 0 Then
            IndexOf = i
            Exit Function
        End If
    Next i
End Function

Public Function LastIndexOf&(Text$, Optional FromIndex& = -1, _
Optional CompareMethod As VbCompareMethod = vbBinaryCompare)
    LastIndexOf = -1
    If Not pSize > 0 Then
        Exit Function
    End If
    Dim StartIndex&
    If FromIndex < 0 Then
        StartIndex = pSize + FromIndex
    Else
        StartIndex = FromIndex
    End If
    If StartIndex < 0 Or StartIndex > pSize - 1 Then
        Err.Raise 9
    End If
    Dim i&
    For i = StartIndex To 0 Step -1
        If StrComp(pStringArray(i), Text, CompareMethod) = 0 Then
            LastIndexOf = i
            Exit Function
        End If
    Next i
End Function

Public Function Count&(Text$, _
Optional CompareMethod As VbCompareMethod = vbBinaryCompare)
    Dim i&
    For i = 0 To pSize - 1
        If StrComp(pStringArray(i), Text, CompareMethod) = 0 Then
            Count = Count + 1
        End If
    Next i
End Function

Public Function Contains(Text$, _
Optional CompareMethod As VbCompareMethod = vbBinaryCompare) As Boolean
    Dim i&
    For i = 0 To pSize - 1
        If StrComp(pStringArray(i), Text, CompareMethod) = 0 Then
            Contains = True
            Exit Function
        End If
    Next i
End Function

Public Sub Clear()
    If Not pSize > 0 Then
        Exit Sub
    End If
    ReDim pStringArray$(0 To Me.Capacity - 1)
    pSize = 0
End Sub

Public Sub Reverse()
    Dim i&
    Dim j&
    Dim Tmp$
    If pSize > 0 Then
        j = pSize
        For i = LBound(pStringArray) To (pSize - 1) \ 2
            j = j - 1
            Tmp = pStringArray(i)
            pStringArray(i) = pStringArray(j)
            pStringArray(j) = Tmp
        Next i
    End If
End Sub

Public Function ToStringArray() As String()
    If Not pSize > 0 Then
        Exit Function
    End If
    Dim Arr$()
    ReDim Arr$(0 To pSize - 1)
    Dim i&
    For i = 0 To pSize - 1
        Arr(i) = pStringArray(i)
    Next i
    ToStringArray = Arr
End Function

Public Function JoinString$(Optional Delimiter As String = "")
    JoinString = Join(Me.ToStringArray, Delimiter)
End Function


'Sorting===============================================================
'======================================================================

Public Sub Sort( _
Optional CompareMethod As VbCompareMethod = vbBinaryCompare)
    Const RUN& = 32
    Dim N&
    Dim L&
    Dim M&
    Dim R&
    Dim i&
    For i = 0 To pSize - 1 Step RUN
        InsertionSort i, Min(i + RUN - 1, pSize - 1), CompareMethod
    Next i
    N = RUN
    Do While N < pSize
        L = 0
        Do While L < pSize
            M = L + N - 1
            R = Min(L + 2 * N - 1, pSize - 1)
            If M < R Then
                Merge L, M, R, CompareMethod
            End If
            L = L + 2 * N
        Loop
        N = N * 2
    Loop
End Sub

Private Function Min&(a&, b&)
    If a < b Then
        Min = a
    Else
        Min = b
    End If
End Function

Private Sub InsertionSort(L&, R&, _
Optional CompareMethod As VbCompareMethod = vbBinaryCompare)
    Dim i&
    Dim j&
    Dim Element
    For i = L + 1 To R
        Element = pStringArray(i)
        j = i - 1
        Do While StrComp(Element, pStringArray(j), CompareMethod) = -1
            pStringArray(j + 1) = pStringArray(j)
            j = j - 1
            If j < L Then
                Exit Do
            End If
        Loop
        pStringArray(j + 1) = Element
    Next i
End Sub

Private Sub Merge(L&, M&, R&, _
Optional CompareMethod As VbCompareMethod = vbBinaryCompare)
    Dim Temp()
    ReDim Temp(0 To (R - L + 1) - 1)
    Dim i&
    Dim j&
    Dim k&
    i = L
    j = M + 1
    k = 0
    Do While i <= M And j <= R
        If StrComp(pStringArray(i), pStringArray(j), _
        CompareMethod) < 1 Then
            Temp(k) = pStringArray(i)
            k = k + 1
            i = i + 1
        Else
            Temp(k) = pStringArray(j)
            k = k + 1
            j = j + 1
        End If
    Loop
    Do While i <= M
        Temp(k) = pStringArray(i)
        k = k + 1
        i = i + 1
    Loop
    Do While j <= R
        Temp(k) = pStringArray(j)
        k = k + 1
        j = j + 1
    Loop
    For i = L To R
        pStringArray(i) = Temp(i - L)
    Next i
End Sub
