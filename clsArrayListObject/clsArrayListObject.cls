VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsArrayListObject"
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
'  Module Name: clsArrayListObject
'  Module Description: ArrayList implementation for Objects.
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
'       CallMethod
'  ----------------------------------------

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'  Example Usage:
'
'        Public Sub Example()
'
'            Dim AL As clsArrayListObject
'            Set AL = New clsArrayListObject
'
'            Dim E As clsExample
'            Set E = New clsExample
'            E.Message = "A"
'
'            Dim Arr() As Object
'            ReDim Arr(0 To 2)
'            Set Arr(0) = New clsExample
'            Set Arr(1) = New clsExample
'            Set Arr(2) = New clsExample
'            Arr(0).Message = "A"
'            Arr(1).Message = "B"
'            Arr(2).Message = "C"
'
'            Dim AL1 As clsArrayListObject
'            Set AL1 = New clsArrayListObject
'            AL1.Append New clsExample
'            AL1.Append New clsExample
'            AL1.Append New clsExample
'            AL1.Item(0).Message = "A"
'            AL1.Item(1).Message = "A"
'            AL1.Item(2).Message = "A"
'
'            AL.Append E
'            AL.AppendArray Arr
'            AL.AppendArrayList AL1
'
'            'AL.Insert 0, E
'            'AL.InsertArray 0, Arr
'            'AL.InsertArrayList 0, AL1
'
'            Debug.Print AL.Count("A", "Message")
'            Debug.Print AL.Contains("A", "Message")
'            Debug.Print AL.JoinString("Message", ", ")
'
'            AL.Remove 0
'            AL.RemoveFirst "A", "Message"
'            AL.RemoveLast "A", "Message"
'            AL.ReplaceAll "A", "B", "Message"
'            AL.RemoveAll "B", "Message"
'
'            AL.Clear
'
'            AL.AppendArray Arr
'
'            AL.Reinitialize
'
'            Set Arr(0) = New clsExample
'            Set Arr(1) = New clsExample
'            Set Arr(2) = New clsExample
'            Arr(0).Message = "C"
'            Arr(1).Message = "B"
'            Arr(2).Message = "A"
'
'            AL.AppendArray Arr
'
'            Debug.Print AL.JoinString("Message", ", ")
'            AL.Sort "Message"
'            Debug.Print AL.JoinString("Message", ", ")
'            AL.Reverse
'            Debug.Print AL.JoinString("Message", ", ")
'
'            AL.CallMethod "MethodSub"
'
'        End Sub
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' ***See testArrayListObject module for Unit Tests


'Private===============================================================
'======================================================================

Private Const DEFAULT_CAPACITY& = 10

Private pArray() As Object
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

Public Property Get Item(Index&) As Object
    If Index < 0 Or Index > pSize - 1 Then
        Err.Raise 9
    End If
    Set Item = pArray(Index)
End Property

Public Property Set Item(Index&, Obj As Object)
    If Index < 0 Or Index > pSize - 1 Then
        Err.Raise 9
    End If
    Set pArray(Index) = Obj
End Property

Public Property Let Item(Index&, Obj)
    If Not IsObject(Obj) Then
        Err.Raise 5
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

Public Sub Append(Obj As Object)
    EnsureCapacity pSize + 1
    Set pArray(pSize) = Obj
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
    If U < L Then
        Err.Raise 9
    End If
    For i = L To U
        If Not IsObject(Arr(i)) Then
            Err.Raise 5
        End If
    Next i
    EnsureCapacity pSize + U - L + 1
    For i = L To U
        Set pArray(pSize) = Arr(i)
        pSize = pSize + 1
    Next i
End Sub

Public Sub AppendArrayList(AL As clsArrayListObject)
    EnsureCapacity pSize + AL.Size
    Dim i&
    For i = 0 To AL.Size - 1
        Set pArray(pSize) = AL.Item(i)
        pSize = pSize + 1
    Next i
End Sub

Public Sub Insert(Index&, Obj As Object)
    If Index < 0 Or Index > pSize Then
        Err.Raise 9
    End If
    EnsureCapacity pSize + 1
    Dim i&
    For i = pSize - 1 To Index Step -1
        Set pArray(i + 1) = pArray(i)
    Next i
    Set pArray(Index) = Obj
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
    If U < L Then
        Err.Raise 9
    End If
    For i = L To U
        If Not IsObject(Arr(i)) Then
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
        Set pArray(i + ArrSize) = pArray(i)
    Next i
    Dim j&
    j = L
    For i = Index To Index + ArrSize - 1
        Set pArray(i) = Arr(j)
        j = j + 1
    Next i
    pSize = pSize + ArrSize
End Sub

Public Sub InsertArrayList(Index&, AL As clsArrayListObject)
    If Index < 0 Or Index > pSize Then
        Err.Raise 9
    End If
    EnsureCapacity pSize + AL.Size
    Dim ALSize&
    ALSize = AL.Size
    Dim i&
    For i = pSize - 1 To Index Step -1
        Set pArray(i + ALSize) = pArray(i)
    Next i
    Dim j&
    j = 0
    For i = Index To Index + ALSize - 1
        Set pArray(i) = AL.Item(j)
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
        Set pArray(i) = pArray(i + 1)
    Next i
    Set pArray(pSize - 1) = Nothing
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
        Set pArray(j) = pArray(i)
        j = j + 1
    Next i
    For i = j To pSize - 1
        Set pArray(i) = Nothing
    Next i
    pSize = pSize - (UpperIndex - LowerIndex + 1)
End Sub

Public Sub RemoveFirst(Value, Member$, _
Optional MemberIsMethod As Boolean = False)
    Dim CallType As VbCallType
    If MemberIsMethod Then
        CallType = VbMethod
    Else
        CallType = VbGet
    End If
    Dim i&
    Dim j&
    For i = 0 To pSize - 1
        If CallByName(pArray(i), Member, CallType) = Value Then
            For j = i To pSize - 2
                Set pArray(j) = pArray(j + 1)
            Next j
            Set pArray(pSize - 1) = Nothing
            pSize = pSize - 1
            Exit Sub
        End If
    Next i
End Sub

Public Sub RemoveLast(Value, Member$, _
Optional MemberIsMethod As Boolean = False)
    Dim CallType As VbCallType
    If MemberIsMethod Then
        CallType = VbMethod
    Else
        CallType = VbGet
    End If
    Dim i&
    Dim j&
    For i = pSize - 1 To 0 Step -1
        If CallByName(pArray(i), Member, CallType) = Value Then
            For j = i To pSize - 2
                Set pArray(j) = pArray(j + 1)
            Next j
            Set pArray(pSize - 1) = Nothing
            pSize = pSize - 1
            Exit Sub
        End If
    Next i
End Sub

Public Sub RemoveAll(Value, Member$, _
Optional MemberIsMethod As Boolean = False)
    Dim CallType As VbCallType
    If MemberIsMethod Then
        CallType = VbMethod
    Else
        CallType = VbGet
    End If
    Dim i&
    Dim j&
    Dim Offset&
    i = 0
    Offset = 0
    Do While i < pSize
        If CallByName(pArray(i), Member, CallType) = Value Then
            Offset = Offset + 1
        Else
            If Offset > 0 Then
                Set pArray(i - Offset) = pArray(i)
            End If
        End If
        i = i + 1
    Loop
    If Offset > 0 Then
        For i = pSize - Offset To pSize - 1
            Set pArray(i) = Nothing
        Next i
        pSize = pSize - Offset
    End If
End Sub

Public Sub ReplaceAll(Value, ReplacementValue, Member$)
    Dim i&
    For i = 0 To pSize - 1
        If CallByName(pArray(i), Member, VbGet) = Value Then
            Select Case VarType(ReplacementValue)
                Case VbVarType.vbBoolean
                    CallByName pArray(i), Member, VbLet, CBool(ReplacementValue)
                Case VbVarType.vbByte
                    CallByName pArray(i), Member, VbLet, CByte(ReplacementValue)
                Case VbVarType.vbCurrency
                    CallByName pArray(i), Member, VbLet, CCur(ReplacementValue)
                Case VbVarType.vbDate
                    CallByName pArray(i), Member, VbLet, CDate(ReplacementValue)
                Case VbVarType.vbDecimal
                    CallByName pArray(i), Member, VbLet, CDec(ReplacementValue)
                Case VbVarType.vbDouble
                    CallByName pArray(i), Member, VbLet, CDbl(ReplacementValue)
                Case VbVarType.vbInteger
                    CallByName pArray(i), Member, VbLet, CInt(ReplacementValue)
                Case VbVarType.vbLong
                    CallByName pArray(i), Member, VbLet, CLng(ReplacementValue)
                #If VBA7 = 1 And Win64 = 1 Then
                Case VbVarType.vbLongLong
                    CallByName pArray(i), Member, VbLet, CLngLng(ReplacementValue)
                #End If
                Case VbVarType.vbSingle
                    CallByName pArray(i), Member, VbLet, CSng(ReplacementValue)
                Case VbVarType.vbString
                    CallByName pArray(i), Member, VbLet, CStr(ReplacementValue)
                Case VbVarType.vbVariant
                    CallByName pArray(i), Member, VbLet, CVar(ReplacementValue)
            End Select
        End If
    Next i
End Sub

Public Function IndexOf&(Value, Member$, Optional FromIndex& = 0, _
Optional MemberIsMethod As Boolean = False)
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
    Dim CallType As VbCallType
    If MemberIsMethod Then
        CallType = VbMethod
    Else
        CallType = VbGet
    End If
    Dim i&
    For i = StartIndex To pSize - 1
        If CallByName(pArray(i), Member, CallType) = Value Then
            IndexOf = i
            Exit Function
        End If
    Next i
End Function

Public Function LastIndexOf&(Value, Member$, Optional FromIndex& = -1, _
Optional MemberIsMethod As Boolean = False)
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
    Dim CallType As VbCallType
    If MemberIsMethod Then
        CallType = VbMethod
    Else
        CallType = VbGet
    End If
    Dim i&
    For i = StartIndex To 0 Step -1
        If CallByName(pArray(i), Member, CallType) = Value Then
            LastIndexOf = i
            Exit Function
        End If
    Next i
End Function

Public Function Count&(Value, Member$, _
Optional MemberIsMethod As Boolean = False)
    Dim CallType As VbCallType
    If MemberIsMethod Then
        CallType = VbMethod
    Else
        CallType = VbGet
    End If
    Dim i&
    For i = 0 To pSize - 1
        If CallByName(pArray(i), Member, CallType) = Value Then
            Count = Count + 1
        End If
    Next i
End Function

Public Function Contains(Value, Member$, _
Optional MemberIsMethod As Boolean = False) As Boolean
    Dim CallType As VbCallType
    If MemberIsMethod Then
        CallType = VbMethod
    Else
        CallType = VbGet
    End If
    Dim i&
    For i = 0 To pSize - 1
        If CallByName(pArray(i), Member, CallType) = Value Then
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
    Dim Tmp As Object
    If pSize > 0 Then
        j = pSize
        For i = LBound(pArray) To (pSize - 1) \ 2
            j = j - 1
            Set Tmp = pArray(i)
            Set pArray(i) = pArray(j)
            Set pArray(j) = Tmp
        Next i
    End If
End Sub

Public Function ToArray() As Object()
    If Not pSize > 0 Then
        Exit Function
    End If
    Dim Arr() As Object
    ReDim Arr(0 To pSize - 1)
    Dim i&
    For i = 0 To pSize - 1
        Set Arr(i) = pArray(i)
    Next i
    ToArray = Arr
End Function

Public Function JoinString$(Member$, Optional Delimiter As String = "", _
Optional MemberIsMethod As Boolean = False)
    If Not pSize > 0 Then
        Exit Function
    End If
    Dim CallType As VbCallType
    If MemberIsMethod Then
        CallType = VbMethod
    Else
        CallType = VbGet
    End If
    Dim Arr()
    ReDim Arr(0 To pSize - 1)
    Dim i&
    For i = 0 To pSize - 1
        Arr(i) = CallByName(pArray(i), Member, CallType)
    Next i
    JoinString = Join(Arr, Delimiter)
End Function

Public Sub CallMethod(Method$)
    Dim i&
    For i = 0 To pSize - 1
        CallByName pArray(i), Method, VbMethod
    Next i
End Sub


'Sorting===============================================================
'======================================================================

Public Sub Sort(Member$, Optional MemberIsMethod As Boolean = False)
    If Not pSize > 0 Then
        Exit Sub
    End If
    Dim CallType As VbCallType
    If MemberIsMethod Then
        CallType = VbMethod
    Else
        CallType = VbGet
    End If
    If pSize < 51 Then
        pInsertionSort Member, CallType
    Else
        pMergeSort Member, CallType
    End If
End Sub

Private Sub pInsertionSort(Member$, CallType As VbCallType)
    Dim i&
    Dim j&
    Dim Element As Object
    For i = 1 To pSize - 1
        Set Element = pArray(i)
        j = i - 1
        Do While CallByName(Element, Member, CallType) < _
        CallByName(pArray(j), Member, CallType)
            Set pArray(j + 1) = pArray(j)
            j = j - 1
            If j < 0 Then
                Exit Do
            End If
        Loop
        Set pArray(j + 1) = Element
    Next i
End Sub

Private Sub pMergeSort(Member$, CallType As VbCallType)
    MergeSortR pArray, 0, pSize - 1, Member, CallType
End Sub

Private Sub MergeSortR(Arr() As Object, First&, Last&, Member$, _
CallType As VbCallType)
    Dim Middle&
    If First < Last Then
        Middle = (First + Last) \ 2
        MergeSortR Arr, First, Middle, Member, CallType
        MergeSortR Arr, Middle + 1, Last, Member, CallType
        Merge Arr, First, Middle, Last, Member, CallType
    End If
End Sub

Private Sub Merge(Arr() As Object, First&, Middle&, Last&, Member$, _
CallType As VbCallType)
    Dim Temp() As Object
    ReDim Temp(0 To (Last - First + 1) - 1)
    Dim i&
    Dim j&
    Dim k&
    i = First
    j = Middle + 1
    k = 0
    Do While i <= Middle And j <= Last
        If CallByName(Arr(i), Member, CallType) <= _
        CallByName(Arr(j), Member, CallType) Then
            Set Temp(k) = Arr(i)
            k = k + 1
            i = i + 1
        Else
            Set Temp(k) = Arr(j)
            k = k + 1
            j = j + 1
        End If
    Loop
    Do While i <= Middle
        Set Temp(k) = Arr(i)
        k = k + 1
        i = i + 1
    Loop
    Do While j <= Last
        Set Temp(k) = Arr(j)
        k = k + 1
        j = j + 1
    Loop
    For i = First To Last
        Set Arr(i) = Temp(i - First)
    Next i
End Sub
