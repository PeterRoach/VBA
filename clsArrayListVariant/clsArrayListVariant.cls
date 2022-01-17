VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsArrayListVariant"
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
'  Module Name: clsArrayListVariant
'  Module Description: ArrayList implementation for value-type Variants.
'  Module Version: 1.0
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
'       ToArray
'       JoinString
'  ----------------------------------------

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'  Example Usage:

'        Public Sub Example()
'
'        End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' ***See testArrayListVariant module for Unit Tests


'Private===============================================================
'======================================================================

Private Const DEFAULT_CAPACITY& = 10

Private pArray()
Private pSize&

Private Sub Class_Initialize()
    ReDim pArray(0 To DEFAULT_CAPACITY - 1)
    pSize = 0
End Sub

Private Function ValidateVariant(Value As Variant) As Boolean

    If IsObject(Value) Then
        Exit Function
    End If

    If VarType(Value) = vbUserDefinedType Then
        Exit Function
    End If

    If IsArray(Value) Then
        Exit Function
    End If

    If IsNull(Value) Then
        Exit Function
    End If

    ValidateVariant = True

End Function

'Public Properties=====================================================
'======================================================================

Public Property Get Capacity&()
    Capacity = UBound(pArray) - LBound(pArray) + 1
End Property

Public Property Get Size&()
    Size = pSize
End Property

Public Property Get Item(Index&)
    If Index < 0 Or Index > pSize - 1 Then
        Err.Raise 9
    End If
    Item = pArray(Index)
End Property

Public Property Let Item(Index&, Value)
    If Not ValidateVariant(Value) Then
        Err.Raise 5
    End If
    If Index < 0 Or Index > pSize - 1 Then
        Err.Raise 9
    End If
    pArray(Index) = Value
End Property

Public Property Set Item(Index&, Value)
    If Not ValidateVariant(Value) Then
        Err.Raise 5
    End If
    If Index < 0 Or Index > pSize - 1 Then
        Err.Raise 9
    End If
    Err.Raise 5
End Property

'Public Methods========================================================
'======================================================================

Public Sub GrowCapacity(Capacity&)
    If Capacity > Me.Capacity Then
        ReDim Preserve pArray(0 To Capacity - 1)
    End If
End Sub

Public Sub ShrinkCapacity(Capacity&)
    If Capacity < 1 Then
        Err.Raise 5
    End If
    If Capacity < Me.Capacity Then
        If Capacity > pSize Then
            ReDim Preserve pArray(0 To Capacity - 1)
        Else
            ReDim Preserve pArray(0 To pSize - 1)
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
        ReDim Preserve pArray(0 To pSize - 1)
    Else
        Me.Reinitialize
    End If
End Sub

Public Sub Reinitialize(Optional Capacity& = DEFAULT_CAPACITY)
    If Capacity < 1 Then
        Err.Raise 5
    End If
    ReDim pArray(0 To Capacity - 1)
    pSize = 0
End Sub

Public Sub Append(Value)
    If Not ValidateVariant(Value) Then
        Err.Raise 5
    End If
    EnsureCapacity pSize + 1
    pArray(pSize) = Value
    pSize = pSize + 1
End Sub

Public Sub AppendArray(Arr)
    If Not IsArray(Arr) Then
        Err.Raise 5
    End If
    Dim i&
    Dim L&
    Dim U&
    L = LBound(Arr)
    U = UBound(Arr)
    For i = L To U
        If Not ValidateVariant(Arr(i)) Then
            Err.Raise 5
        End If
    Next i
    EnsureCapacity pSize + U - L + 1
    For i = L To U
        pArray(pSize) = Arr(i)
        pSize = pSize + 1
    Next i
End Sub

Public Sub AppendArrayList(AL As clsArrayListVariant)
    EnsureCapacity pSize + AL.Size
    Dim i&
    For i = 0 To AL.Size - 1
        pArray(pSize) = AL.Item(i)
        pSize = pSize + 1
    Next i
End Sub

Public Sub Insert(Index&, Value)
    If Not ValidateVariant(Value) Then
        Err.Raise 5
    End If
    If Index < 0 Or Index > pSize Then
        Err.Raise 9
    End If
    EnsureCapacity pSize + 1
    Dim i&
    For i = pSize - 1 To Index Step -1
        pArray(i + 1) = pArray(i)
    Next i
    pArray(Index) = Value
    pSize = pSize + 1
End Sub

Public Sub InsertArray(Index&, Arr)
    If Not IsArray(Arr) Then
        Err.Raise 5
    End If
    Dim i&
    Dim L&
    Dim U&
    L = LBound(Arr)
    U = UBound(Arr)
    For i = L To U
        If Not ValidateVariant(Arr(i)) Then
            Err.Raise 5
        End If
    Next i
    If Index < 0 Or Index > pSize Then
        Err.Raise 9
    End If
    Dim ArrSize&
    ArrSize = U - L + 1
    EnsureCapacity pSize + ArrSize
    For i = pSize - 1 To Index Step -1
        pArray(i + ArrSize) = pArray(i)
    Next i
    Dim j&
    j = L
    For i = Index To Index + ArrSize - 1
        pArray(i) = Arr(j)
        j = j + 1
    Next i
    pSize = pSize + ArrSize
End Sub

Public Sub InsertArrayList(Index&, AL As clsArrayListVariant)
    If Index < 0 Or Index > pSize Then
        Err.Raise 9
    End If
    EnsureCapacity pSize + AL.Size
    Dim ALSize&
    ALSize = AL.Size
    Dim i&
    For i = pSize - 1 To Index Step -1
        pArray(i + ALSize) = pArray(i)
    Next i
    Dim j&
    j = 0
    For i = Index To Index + ALSize - 1
        pArray(i) = AL.Item(j)
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
        pArray(i) = pArray(i + 1)
    Next i
    pArray(pSize - 1) = Empty
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
        pArray(j) = pArray(i)
        j = j + 1
    Next i
    For i = j To pSize - 1
        pArray(i) = Empty
    Next i
    pSize = pSize - (UpperIndex - LowerIndex + 1)
End Sub

Public Sub RemoveFirst(Value)
    If Not ValidateVariant(Value) Then
        Err.Raise 5
    End If
    Dim i&
    Dim j&
    For i = 0 To pSize - 1
        If pArray(i) = Value Then
            For j = i To pSize - 2
                pArray(j) = pArray(j + 1)
            Next j
            pArray(pSize - 1) = Empty
            pSize = pSize - 1
            Exit Sub
        End If
    Next i
End Sub

Public Sub RemoveLast(Value)
    If Not ValidateVariant(Value) Then
        Err.Raise 5
    End If
    Dim i&
    Dim j&
    For i = pSize - 1 To 0 Step -1
        If pArray(i) = Value Then
            For j = i To pSize - 2
                pArray(j) = pArray(j + 1)
            Next j
            pArray(pSize - 1) = Empty
            pSize = pSize - 1
            Exit Sub
        End If
    Next i
End Sub

Public Sub RemoveAll(Value)
    If Not ValidateVariant(Value) Then
        Err.Raise 5
    End If
    Dim i&
    Dim j&
    Dim Offset&
    i = 0
    Offset = 0
    Do While i < pSize
        If pArray(i) = Value Then
            Offset = Offset + 1
        Else
            If Offset > 0 Then
                pArray(i - Offset) = pArray(i)
            End If
        End If
        i = i + 1
    Loop
    If Offset > 0 Then
        For i = pSize - Offset To pSize - 1
            pArray(i) = Empty
        Next i
        pSize = pSize - Offset
    End If
End Sub

Public Sub ReplaceAll(Value, ReplacementValue)
    If Not ValidateVariant(Value) Then
        Err.Raise 5
    End If
    If Not ValidateVariant(ReplacementValue) Then
        Err.Raise 5
    End If
    Dim i&
    For i = 0 To pSize - 1
        If pArray(i) = Value Then
            pArray(i) = ReplacementValue
        End If
    Next i
End Sub

Public Function IndexOf&(Value, Optional FromIndex& = 0)
    If Not ValidateVariant(Value) Then
        Err.Raise 5
    End If
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
        If pArray(i) = Value Then
            IndexOf = i
            Exit Function
        End If
    Next i
End Function

Public Function LastIndexOf&(Value, Optional FromIndex& = -1)
    If Not ValidateVariant(Value) Then
        Err.Raise 5
    End If
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
        If pArray(i) = Value Then
            LastIndexOf = i
            Exit Function
        End If
    Next i
End Function

Public Function Count&(Value)
    If Not ValidateVariant(Value) Then
        Err.Raise 5
    End If
    Dim i&
    For i = 0 To pSize - 1
        If pArray(i) = Value Then
            Count = Count + 1
        End If
    Next i
End Function

Public Function Contains(Value) As Boolean
    If Not ValidateVariant(Value) Then
        Err.Raise 5
    End If
    Dim i&
    For i = 0 To pSize - 1
        If pArray(i) = Value Then
            Contains = True
            Exit Function
        End If
    Next i
End Function

Public Sub Clear()
    If Not pSize > 0 Then
        Exit Sub
    End If
    ReDim pArray(0 To Me.Capacity - 1)
    pSize = 0
End Sub

Public Sub Reverse()
    Dim i&
    Dim j&
    Dim Tmp
    If pSize > 0 Then
        j = pSize
        For i = LBound(pArray) To (pSize - 1) \ 2
            j = j - 1
            Tmp = pArray(i)
            pArray(i) = pArray(j)
            pArray(j) = Tmp
        Next i
    End If
End Sub

Public Function ToArray() As Variant()
    If Not pSize > 0 Then
        Exit Function
    End If
    Dim Arr()
    ReDim Arr(0 To pSize - 1)
    Dim i&
    For i = 0 To pSize - 1
        Arr(i) = pArray(i)
    Next i
    ToArray = Arr
End Function

Public Function JoinString$(Optional Delimiter As String = "")
    JoinString = Join(Me.ToArray, Delimiter)
End Function


'Sorting===============================================================
'======================================================================

Public Sub Sort()
    If Not pSize > 0 Then
        Exit Sub
    End If
    If pSize < 51 Then
        pInsertionSort
    Else
        pMergeSort
    End If
End Sub

Private Sub pInsertionSort()
    Dim i&
    Dim j&
    Dim Element
    For i = 1 To pSize - 1
        Element = pArray(i)
        j = i - 1
        Do While Element < pArray(j)
            pArray(j + 1) = pArray(j)
            j = j - 1
            If j < 0 Then
                Exit Do
            End If
        Loop
        pArray(j + 1) = Element
    Next i
End Sub

Private Sub pMergeSort()
    MergeSortR pArray, 0, pSize - 1
End Sub

Private Sub MergeSortR(Arr(), First&, Last&)
    Dim Middle&
    If First < Last Then
        Middle = (First + Last) \ 2
        MergeSortR Arr, First, Middle
        MergeSortR Arr, Middle + 1, Last
        Merge Arr, First, Middle, Last
    End If
End Sub

Private Sub Merge(Arr(), First&, Middle&, Last&)
    Dim Temp()
    ReDim Temp(0 To (Last - First + 1) - 1)
    Dim i&
    Dim j&
    Dim k&
    i = First
    j = Middle + 1
    k = 0
    Do While i <= Middle And j <= Last
        If Arr(i) <= Arr(j) Then
            Temp(k) = Arr(i)
            k = k + 1
            i = i + 1
        Else
            Temp(k) = Arr(j)
            k = k + 1
            j = j + 1
        End If
    Loop
    Do While i <= Middle
        Temp(k) = Arr(i)
        k = k + 1
        i = i + 1
    Loop
    Do While j <= Last
        Temp(k) = Arr(j)
        k = k + 1
        j = j + 1
    Loop
    For i = First To Last
        Arr(i) = Temp(i - First)
    Next i
End Sub
