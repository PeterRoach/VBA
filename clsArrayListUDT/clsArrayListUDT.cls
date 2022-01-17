VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsArrayListUDT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'INSTRUCTIONS:
    '-> Replace TExample with a public user-defined type. Ctrl + h
    '-> Replace clsArrayListUDT with clsArrayList[typename] and rename class.

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
'  Module Name: clsArrayListUDT
'  Module Description: ArrayList implementation for User-Defined Types.
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
'       Clear
'       Reverse
'       ToArray
'  ----------------------------------------

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'  Example Usage:

'    Public Sub Example()
'
'        Dim AL As clsArrayListUDT
'        Set AL = New clsArrayListObject
'
'        Dim E As TExample
'        E.Message = "Hello, World!"
'        E.Flag = True
'
'        AL.Append E
'        AL.Append E
'        AL.Append E
'
'        Set AL = Nothing
'
'    End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' ***See testArrayListUDT module for Unit Tests


'Private===============================================================
'======================================================================

Private Const DEFAULT_CAPACITY& = 10

Private pArray() As TExample
Private pSize&

Private Sub Class_Initialize()
    ReDim pArray(0 To DEFAULT_CAPACITY - 1)
    pSize = 0
End Sub


'Public Properties=====================================================
'======================================================================

Public Property Get Capacity&()
    Capacity = UBound(pArray) - LBound(pArray) + 1
End Property

Public Property Get Size&()
    Size = pSize
End Property

Public Property Get Item(Index&) As TExample
    If Index < 0 Or Index > pSize - 1 Then
        Err.Raise 9
    End If
    Item = pArray(Index)
End Property

Public Property Let Item(Index&, Typ As TExample)
    If Index < 0 Or Index > pSize - 1 Then
        Err.Raise 9
    End If
    pArray(Index) = Typ
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

Public Sub Append(Typ As TExample)
    EnsureCapacity pSize + 1
    pArray(pSize) = Typ
    pSize = pSize + 1
End Sub

Public Sub AppendArray(Arr() As TExample)
    Dim L&
    Dim U&
    L = LBound(Arr)
    U = UBound(Arr)
    EnsureCapacity pSize + U - L + 1
    Dim i&
    For i = L To U
        pArray(pSize) = Arr(i)
        pSize = pSize + 1
    Next i
End Sub

Public Sub AppendArrayList(AL As clsArrayListUDT)
    EnsureCapacity pSize + AL.Size
    Dim i&
    For i = 0 To AL.Size - 1
        pArray(pSize) = AL.Item(i)
        pSize = pSize + 1
    Next i
End Sub

Public Sub Insert(Index&, Typ As TExample)
    If Index < 0 Or Index > pSize Then
        Err.Raise 9
    End If
    EnsureCapacity pSize + 1
    Dim i&
    For i = pSize - 1 To Index Step -1
        pArray(i + 1) = pArray(i)
    Next i
    pArray(Index) = Typ
    pSize = pSize + 1
End Sub

Public Sub InsertArray(Index&, Arr() As TExample)
    If Index < 0 Or Index > pSize Then
        Err.Raise 9
    End If
    Dim ArrSize&
    ArrSize = UBound(Arr) - LBound(Arr) + 1
    EnsureCapacity pSize + ArrSize
    Dim i&
    For i = pSize - 1 To Index Step -1
        pArray(i + ArrSize) = pArray(i)
    Next i
    Dim j&
    j = LBound(Arr)
    For i = Index To Index + ArrSize - 1
        pArray(i) = Arr(j)
        j = j + 1
    Next i
    pSize = pSize + ArrSize
End Sub

Public Sub InsertArrayList(Index&, AL As clsArrayListUDT)
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
    Dim T As TExample
    pArray(pSize - 1) = T
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
    Dim T As TExample
    For i = j To pSize - 1
        pArray(i) = T
    Next i
    pSize = pSize - (UpperIndex - LowerIndex + 1)
End Sub

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
    Dim Tmp As TExample
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

Public Function ToArray() As TExample()
    If Not pSize > 0 Then
        Exit Function
    End If
    Dim Arr() As TExample
    ReDim Arr(0 To pSize - 1)
    Dim i&
    For i = 0 To pSize - 1
        Arr(i) = pArray(i)
    Next i
    ToArray = Arr
End Function
