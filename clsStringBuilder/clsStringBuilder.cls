VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsStringBuilder"
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
'  Module Name: clsStringBuilder
'  Module Description: StringBuilder class implementation.
'  Module Version: 1.0
'  Module License: MIT
'  Module Author: Peter Roach; PeterRoach.Code@gmail.com
'  Module Contents:
'  ----------------------------------------
'    Public Properties:
'        Capacity
'        Size
'    Public Methods:
'        Get Capacity
'        Get Size
'        Reinitialize
'        EnsureCapacity
'        GrowCapacity
'        ShrinkCapacity
'        TrimToSize
'        CharAt
'        Append
'        Insert
'        Remove
'        CharPositionOf
'        LastCharPositionOf
'        Replace
'        Reverse
'        Clear
'        Substring
'        ToString
'  ----------------------------------------

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'  Example Usage:
'
'    Public Sub Example()
'
'        Dim SB As clsStringBuilder
'        Set SB = New clsStringBuilder
'
'        SB.Append "ABCDABCD"
'
'        Debug.Print SB.CharAt(0)
'
'        Debug.Print SB.CharPositionOf("C")
'
'        Debug.Print SB.LastCharPositionOf("C")
'
'        Debug.Print SB.Substring(0, 3)
'
'        SB.Insert 4, ","
'
'        Debug.Print SB.ToString
'
'        SB.Remove 4, 4
'
'        Debug.Print SB.ToString
'
'        Debug.Print SB.Size, SB.Capacity
'
'        SB.Reverse
'
'        Debug.Print SB.ToString
'
'        Set SB = Nothing
'
'    End Sub
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' ***See testStringBuilder module for Unit Tests for this class module


'Private===============================================================
'======================================================================

Private Const DEFAULT_CAPACITY& = 16

Private pByteArray() As Byte
Private pSize&

Private Sub Class_Initialize()
    ReDim pByteArray(0 To DEFAULT_CAPACITY * 2 - 1) As Byte
    pSize = 0
End Sub


'Public Properties=====================================================
'======================================================================

Public Property Get Capacity&()
    Capacity = (UBound(pByteArray) - LBound(pByteArray) + 1) / 2
End Property

Public Property Get Size&()
    Size = pSize / 2
End Property


'Public Methods========================================================
'======================================================================

Public Sub Reinitialize(Optional Capacity& = DEFAULT_CAPACITY)
    If Capacity < 1 Then
        Err.Raise 5
    End If
    ReDim pByteArray(0 To Capacity * 2 - 1) As Byte
    pSize = 0
End Sub

Public Sub EnsureCapacity(MinimumCapacity&)
    Dim CurrentCapacity&
    Dim DoubleCapacity&
    CurrentCapacity = Me.Capacity
    DoubleCapacity = CurrentCapacity * 2 + 2
    If MinimumCapacity > CurrentCapacity Then
        If DoubleCapacity > MinimumCapacity Then
            Me.GrowCapacity DoubleCapacity
        Else
            Me.GrowCapacity MinimumCapacity
        End If
    End If
End Sub

Public Sub GrowCapacity(Capacity&)
    If Capacity > Me.Capacity Then
        ReDim Preserve pByteArray(0 To Capacity * 2 - 1) As Byte
    End If
End Sub

Public Sub ShrinkCapacity(Capacity&)
    If Capacity < 1 Then
        Err.Raise 5
    End If
    If Capacity < Me.Capacity Then
        If Capacity * 2 > pSize Then
            ReDim Preserve pByteArray(0 To Capacity * 2 - 1) As Byte
        Else
            ReDim Preserve pByteArray(0 To pSize - 1) As Byte
        End If
    End If
End Sub

Public Sub TrimToSize()
    If pSize > 0 Then
        ReDim Preserve pByteArray(0 To pSize - 1) As Byte
    Else
        Me.Reinitialize
    End If
End Sub

Public Function CharAt$(CharPosition As Long)
    Dim CharBytes(0 To 1) As Byte
    Dim Index&
    Index = CharPosition * 2
    If Index < 0 Or Index > pSize - 1 Then
        Err.Raise 9
    End If
    CharBytes(0) = pByteArray(Index)
    CharBytes(1) = pByteArray(Index + 1)
    CharAt = CharBytes
End Function

Public Sub Append(Text$)
    Dim TmpByteArray() As Byte
    Dim i&
    If Len(Text) = 0 Then
        Exit Sub
    End If
    TmpByteArray = Text
    Me.EnsureCapacity _
    pSize / 2 + (UBound(TmpByteArray) - LBound(TmpByteArray) + 1) / 2
    For i = LBound(TmpByteArray) To UBound(TmpByteArray)
        pByteArray(pSize) = TmpByteArray(i)
        pSize = pSize + 1
    Next i
End Sub

Public Sub Insert(CharPosition&, Text$)
    Dim TmpByteArray() As Byte
    Dim TmpByteArraySize&
    Dim i&
    Dim Index&
    Index = CharPosition * 2
    If Index < 0 Or Index > pSize Then
        Err.Raise 9
    End If
    If Len(Text) = 0 Then
        Exit Sub
    End If
    TmpByteArray = Text
    TmpByteArraySize = UBound(TmpByteArray) - LBound(TmpByteArray) + 1
    Me.EnsureCapacity pSize / 2 + TmpByteArraySize / 2
    For i = pSize - 1 To Index Step -1
        pByteArray(i + TmpByteArraySize) = pByteArray(i)
    Next i
    For i = LBound(TmpByteArray) To UBound(TmpByteArray)
        pByteArray(Index) = TmpByteArray(i)
        Index = Index + 1
    Next i
    pSize = pSize + TmpByteArraySize
End Sub

Public Sub Remove(CharPositionStart&, CharPositionEnd&)
    Dim StartIndex&
    Dim EndIndex&
    Dim i&
    Dim j&
    If CharPositionStart > CharPositionEnd Then
        Err.Raise 5
    End If
    StartIndex = CharPositionStart * 2
    EndIndex = CharPositionEnd * 2
    If StartIndex < 0 Or StartIndex > pSize - 1 Then
        Err.Raise 9
    End If
    If EndIndex < 0 Or EndIndex > pSize - 1 Then
        Err.Raise 9
    End If
    j = StartIndex
    For i = EndIndex + 2 To pSize - 1
        pByteArray(j) = pByteArray(i)
        j = j + 1
    Next i
    For i = j To pSize - 1
        pByteArray(i) = 0
    Next i
    pSize = pSize - (EndIndex - StartIndex + 2)
End Sub

Public Function CharPositionOf&(Text$, Optional FromCharPosition& = 0)
    CharPositionOf = -1
    If Text = "" Then
        Exit Function
    End If
    Dim FromIndex&
    If FromCharPosition < 0 Then
        FromIndex = Me.Size + FromCharPosition + 1
    Else
        FromIndex = FromCharPosition + 1
    End If
    If FromIndex < 1 Then
        Err.Raise 9
    End If
    Dim S$
    S = Me.ToString
    If FromIndex > Len(S) Then
        Err.Raise 9
    End If
    Dim i&
    i = InStr(FromIndex, S, Text)
    If i < 0 Then
        CharPositionOf = i
    Else
        CharPositionOf = i - 1
    End If
End Function

Public Function LastCharPositionOf&(Text$, Optional FromCharPosition& = -1)
    LastCharPositionOf = -1
    If Text = "" Then
        Exit Function
    End If
    Dim FromIndex&
    If FromCharPosition < 0 Then
        FromIndex = Me.Size + FromCharPosition + 1
    Else
        FromIndex = FromCharPosition + 1
    End If
    If FromIndex < 1 Then
        Err.Raise 9
    End If
    Dim S$
    S = Me.ToString
    If FromIndex > Len(S) Then
        Err.Raise 9
    End If
    Dim i&
    i = InStrRev(S, Text, FromIndex)
    If i < 0 Then
        LastCharPositionOf = i
    Else
        LastCharPositionOf = i - 1
    End If
End Function

Public Sub Replace(CharPositionStart&, CharPositionEnd&, Text$)
    Dim StartIndex&
    Dim EndIndex&
    Dim TmpByteArray() As Byte
    Dim TmpByteArraySize&
    Dim ReplaceSize&
    Dim OffsetSize&
    Dim i&
    Dim j&
    If CharPositionStart > CharPositionEnd Then
        Err.Raise 5
    End If
    StartIndex = CharPositionStart * 2
    EndIndex = CharPositionEnd * 2
    If StartIndex < 0 Or StartIndex > pSize - 1 Then
        Err.Raise 9
    End If
    If EndIndex < 0 Or EndIndex > pSize - 1 Then
        Err.Raise 9
    End If
    TmpByteArray = Text
    TmpByteArraySize = UBound(TmpByteArray) - LBound(TmpByteArray) + 1
    Me.EnsureCapacity Me.Size + TmpByteArraySize / 2
    ReplaceSize = EndIndex - StartIndex + 2
    If ReplaceSize > TmpByteArraySize Then
        OffsetSize = ReplaceSize - TmpByteArraySize
        For i = EndIndex + 2 To pSize - 1
            pByteArray(i - OffsetSize) = pByteArray(i)
            pByteArray(i) = 0
        Next i
    ElseIf ReplaceSize < TmpByteArraySize Then
        OffsetSize = TmpByteArraySize - ReplaceSize
        For i = pSize - 1 To EndIndex + 2 Step -1
            pByteArray(i + OffsetSize) = pByteArray(i)
            pByteArray(i) = 0
        Next i
    End If
    j = StartIndex
    For i = LBound(TmpByteArray) To UBound(TmpByteArray)
        pByteArray(j) = TmpByteArray(i)
        j = j + 1
    Next i
    pSize = pSize - ReplaceSize + TmpByteArraySize
End Sub

Public Sub Reverse()
    Dim i&
    Dim j&
    Dim Tmp As Byte
    If pSize = 0 Then
        Exit Sub
    End If
    j = pSize - 1
    For i = LBound(pByteArray) To (pSize - 1) \ 2 Step 2
        Tmp = pByteArray(i + 1)
        pByteArray(i + 1) = pByteArray(j)
        pByteArray(j) = Tmp
        j = j - 1
        Tmp = pByteArray(i)
        pByteArray(i) = pByteArray(j)
        pByteArray(j) = Tmp
        j = j - 1
    Next i
End Sub

Public Sub Clear()
    ReDim pByteArray(0 To Me.Capacity * 2 - 1) As Byte
    pSize = 0
End Sub

Public Function Substring$(CharPositionStart&, CharPositionEnd&)
    Dim SubStrByteArray() As Byte
    Dim i&
    Dim j&
    Dim StartIndex&
    Dim EndIndex&
    If CharPositionStart > CharPositionEnd Then
        Err.Raise 5
    End If
    StartIndex = CharPositionStart * 2
    EndIndex = CharPositionEnd * 2
    If StartIndex < 0 Or StartIndex > pSize - 1 Then
        Err.Raise 9
    End If
    If EndIndex < 0 Or EndIndex > pSize - 1 Then
        Err.Raise 9
    End If
    ReDim SubStrByteArray(0 To EndIndex - StartIndex + 1) As Byte
    For i = StartIndex To EndIndex
        SubStrByteArray(j) = pByteArray(i)
        j = j + 1
    Next i
    Substring = SubStrByteArray
End Function

Public Function ToString$()
    If pSize = 0 Then
        ToString = ""
        Exit Function
    End If
    Dim ByteArray() As Byte
    ByteArray = pByteArray
    ReDim Preserve ByteArray(0 To pSize - 1) As Byte
    ToString = ByteArray
End Function
