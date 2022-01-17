Attribute VB_Name = "modArray"
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
'  Module Name: modArray
'  Module Description: Contains high-level array functions.
'  Module Version: 1.0
'  Module License: MIT
'  Module Author: Peter Roach; PeterRoach.Code@gmail.com
'  Module Contents:
'  ----------------------------------------
'    Public Procedures:
'       ArrayByte
'       ArrayInteger
'       ArrayLong
'       ArrayLongLong
'       ArrayLongPtr
'       ArraySingle
'       ArrayDouble
'       ArrayDecimal
'       ArrayCurrency
'       ArrayDate
'       ArrayBoolean
'       ArrayString
'       ConcatArray
'       ConcatObjectArray
'       CountArray
'       CountObjectArray
'       CountStringArray
'       DimensionsOfArray
'       GenerateLongArray
'       GenerateDoubleArray
'       GeneratePrimesArray
'       GeneratePatternArray
'       GenerateFibonacciArray
'       GenerateRndArray
'       GenerateRandomLongArray
'       GenerateRandomDoubleArray
'       ReverseArray
'       ReverseObjectArray
'       SearchArray
'       SearchStringArray
'       SearchObjectArray
'       ShuffleArray
'       ShuffleObjectArray
'       SizeOfArray
'       SortArray
'       SortStringArray
'       SortObjectArray
'       ToByteArray
'       ToIntegerArray
'       ToLongArray
'       ToLongLongArray
'       ToLongPtrArray
'       ToSingleArray
'       ToDoubleArray
'       ToDecimalArray
'       ToCurrencyArray
'       ToDateArray
'       ToBooleanArray
'       ToStringArray
'    Private Procedures:
'       pIsPrime
'       pRandomLong
'       pRandomDouble
'       pInsertionSortV
'       pMergeSortV
'       pMergeSortRV
'       pMergeV
'       pInsertionSortS
'       pMergeSortS
'       pMergeSortRS
'       pMergeS
'       pInsertionSortO
'       pMergeSortO
'       pMergeSortRO
'       pMergeO
'    Test Procedures:
'       TestTypedArrayFunctions
'       TestConcatArray
'       TestConcatObjectArray
'       TestCountArray
'       TestCountStringArray
'       TestCountObjectArray
'       TestDimensionsOfArray
'       TestGenerateLongArray
'       TestGenerateDoubleArray
'       TestGeneratePrimesArray
'       TestGeneratePatternArray
'       TestGenerateFibonacciArray
'       TestGenerateRndArray
'       TestGenerateRandomLongArray
'       TestGenerateRandomDoubleArray
'       TestSearchArray
'       TestSearchStringArray
'       TestSearchObjectArray
'       TestShuffleArray
'       TestShuffleObjectArray
'       TestSizeOfArray
'       TestSortArray
'       TestInsertionSortV
'       TestMergeSortV
'       TestSortStringArray
'       TestInsertionSortS
'       TestMergeSortS
'       TestSortObjectArray
'       TestInsertionSortO
'       TestMergeSortO
'       TestReverseArray
'       TestReverseObjectArray
'       TestIsPrime
'       TestRandomLong
'       TestRandomDouble
'       TestToByteArray
'       TestToIntegerArray
'       TestToLongArray
'       TestToLongLongArray
'       TestToLongPtrArray
'       TestToSingleArray
'       TestToDoubleArray
'       TestToDecimalArray
'       TestToCurrencyArray
'       TestToDateArray
'       TestToBooleanArray
'       TestToStringArray
'       TestToVariantArray
'  ----------------------------------------

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'  Example Usage:

    Private Sub Example()

        '-> Get arrays of each type
        '-> Convert arrays to different types
        '-> Generate special kinds arrays
        '-> Get array size and dimensions
        '-> Sort arrays
        '-> Reverse arrays
        '-> Shuffle arrays
        '-> Search for elements in arrays
        '-> Count particular elements in arrays
        '-> Concatenate arrays

        Dim Arr() As Long
        Arr = GenerateRandomLongArray(1, 3, 10)
        Debug.Print "Random: " & Join(ToStringArray(Arr), ", ")

        SortArray Arr
        Debug.Print "Sorted: " & Join(ToStringArray(Arr), ", ")

        ReverseArray Arr
        Debug.Print "Reversed: " & Join(ToStringArray(Arr), ", ")

        ShuffleArray Arr
        Debug.Print "Suffled: " & Join(ToStringArray(Arr), ", ")

        Debug.Print "First index of 1: " & SearchArray(Arr, 1, True) 'search left to right
        Debug.Print "Last index of 1: " & SearchArray(Arr, 1, False) 'search right to left

        Debug.Print "Count of 1: " & CountArray(Arr, 1)

        Dim Arr1() As Long
        Arr1 = GeneratePrimesArray(2, 10)

        ConcatArray Arr, Arr1
        Debug.Print "Array + Primes: " & Join(ToStringArray(Arr), ", ")

    End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


'Private Helper Functions=================================================
'=========================================================================

Private Function pIsPrime(N&) As Boolean
    Select Case N
        Case Is < 1
            Err.Raise 5
        Case 1
            Exit Function
        Case Is < 4
            pIsPrime = True
            Exit Function
        Case Else
            pIsPrime = True
            Dim i&
            For i = 2 To N ^ (1 / 2)
                If N Mod i = 0 Then
                    pIsPrime = False
                    Exit Function
                End If
            Next i
    End Select
End Function

Private Function pRandomLong&(MinValue&, MaxValue&)
    Randomize
    pRandomLong = Int((MaxValue - MinValue + 1) * Rnd + MinValue)
End Function

Private Function pRandomDouble#(MinValue#, MaxValue#)
    Randomize
    pRandomDouble = (MaxValue - MinValue) * Rnd + MinValue
End Function

Private Sub pInsertionSortV(Arr, N&)
    Dim i&
    Dim j&
    Dim Element
    For i = 1 To N - 1
        Element = Arr(i)
        j = i - 1
        Do While Element < Arr(j)
            Arr(j + 1) = Arr(j)
            j = j - 1
            If j < 0 Then
                Exit Do
            End If
        Loop
        Arr(j + 1) = Element
    Next i
End Sub

Private Sub pMergeSortV(Arr, N&)
    pMergeSortRV Arr, 0, N - 1
End Sub

Private Sub pMergeSortRV(Arr, First&, Last&)
    Dim Middle&
    If First < Last Then
        Middle = (First + Last) \ 2
        pMergeSortRV Arr, First, Middle
        pMergeSortRV Arr, Middle + 1, Last
        pMergeV Arr, First, Middle, Last
    End If
End Sub

Private Sub pMergeV(Arr, First&, Middle&, Last&)
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

Private Sub pInsertionSortS(Arr$(), N&, CompareMethod As VbCompareMethod)
    Dim i&
    Dim j&
    Dim Element$
    For i = 1 To N - 1
        Element = Arr(i)
        j = i - 1
        Do While StrComp(Element, Arr(j), CompareMethod) = -1
            Arr(j + 1) = Arr(j)
            j = j - 1
            If j < 0 Then
                Exit Do
            End If
        Loop
        Arr(j + 1) = Element
    Next i
End Sub

Private Sub pMergeSortS(Arr$(), N&, CompareMethod As VbCompareMethod)
    pMergeSortRS Arr, 0, N - 1, CompareMethod
End Sub

Private Sub pMergeSortRS(Arr$(), First&, Last&, _
CompareMethod As VbCompareMethod)
    Dim Middle&
    If First < Last Then
        Middle = (First + Last) \ 2
        pMergeSortRS Arr, First, Middle, CompareMethod
        pMergeSortRS Arr, Middle + 1, Last, CompareMethod
        pMergeS Arr, First, Middle, Last, CompareMethod
    End If
End Sub

Private Sub pMergeS(Arr$(), First&, Middle&, Last&, _
CompareMethod As VbCompareMethod)
    Dim Temp$()
    ReDim Temp$(0 To (Last - First + 1) - 1)
    Dim i&
    Dim j&
    Dim k&
    i = First
    j = Middle + 1
    k = 0
    Do While i <= Middle And j <= Last
        If StrComp(Arr(i), Arr(j), CompareMethod) < 1 Then
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

Private Sub pInsertionSortO(Arr, N&, Member$, CallType As VbCallType)
    Dim i&
    Dim j&
    Dim Element As Object
    For i = 1 To N - 1
        Set Element = Arr(i)
        j = i - 1
        Do While CallByName(Element, Member, CallType) < _
        CallByName(Arr(j), Member, CallType)
            Set Arr(j + 1) = Arr(j)
            j = j - 1
            If j < 0 Then
                Exit Do
            End If
        Loop
        Set Arr(j + 1) = Element
    Next i
End Sub

Private Sub pMergeSortO(Arr, N&, Member$, CallType As VbCallType)
    pMergeSortRO Arr, 0, N - 1, Member, CallType
End Sub

Private Sub pMergeSortRO(Arr, First&, Last&, Member$, _
CallType As VbCallType)
    Dim Middle&
    If First < Last Then
        Middle = (First + Last) \ 2
        pMergeSortRO Arr, First, Middle, Member, CallType
        pMergeSortRO Arr, Middle + 1, Last, Member, CallType
        pMergeO Arr, First, Middle, Last, Member, CallType
    End If
End Sub

Private Sub pMergeO(Arr, First&, Middle&, Last&, Member$, _
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


'Typed Array Functions====================================================
'=========================================================================

Public Function ArrayByte(ParamArray Bytes()) As Byte()
    Dim Arr() As Byte
    ReDim Arr(LBound(Bytes) To UBound(Bytes))
    Dim i&
    For i = LBound(Bytes) To UBound(Bytes)
        Arr(i) = CByte(Bytes(i))
    Next i
    ArrayByte = Arr
End Function

Public Function ArrayInteger(ParamArray Integers()) As Integer()
    Dim Arr%()
    ReDim Arr(LBound(Integers) To UBound(Integers))
    Dim i&
    For i = LBound(Integers) To UBound(Integers)
        Arr(i) = CInt(Integers(i))
    Next i
    ArrayInteger = Arr
End Function

Public Function ArrayLong(ParamArray Longs()) As Long()
    Dim Arr&()
    ReDim Arr(LBound(Longs) To UBound(Longs))
    Dim i&
    For i = LBound(Longs) To UBound(Longs)
        Arr(i) = CLng(Longs(i))
    Next i
    ArrayLong = Arr
End Function

#If VBA7 = 1 And Win64 = 1 Then
    Public Function ArrayLongLong(ParamArray LongLongs()) As LongLong()
        Dim Arr^()
        ReDim Arr(LBound(LongLongs) To UBound(LongLongs))
        Dim i&
        For i = LBound(LongLongs) To UBound(LongLongs)
            Arr(i) = CLngLng(LongLongs(i))
        Next i
        ArrayLongLong = Arr
    End Function
#End If

#If VBA7 Then
    Public Function ArrayLongPtr(ParamArray LongPtrs()) As LongPtr()
        Dim Arr() As LongPtr
        ReDim Arr(LBound(LongPtrs) To UBound(LongPtrs))
        Dim i&
        For i = LBound(LongPtrs) To UBound(LongPtrs)
            Arr(i) = CLngPtr(LongPtrs(i))
        Next i
        ArrayLongPtr = Arr
    End Function
#End If

Public Function ArraySingle(ParamArray Singles()) As Single()
    Dim Arr!()
    ReDim Arr(LBound(Singles) To UBound(Singles))
    Dim i&
    For i = LBound(Singles) To UBound(Singles)
        Arr(i) = CSng(Singles(i))
    Next i
    ArraySingle = Arr
End Function

Public Function ArrayDouble(ParamArray Doubles()) As Double()
    Dim Arr#()
    ReDim Arr(LBound(Doubles) To UBound(Doubles))
    Dim i&
    For i = LBound(Doubles) To UBound(Doubles)
        Arr(i) = CDbl(Doubles(i))
    Next i
    ArrayDouble = Arr
End Function

Public Function ArrayDecimal(ParamArray Decimals()) As Variant()
    Dim Arr()
    ReDim Arr(LBound(Decimals) To UBound(Decimals))
    Dim i&
    For i = LBound(Decimals) To UBound(Decimals)
        Arr(i) = CDec(Decimals(i))
    Next i
    ArrayDecimal = Arr
End Function

Public Function ArrayCurrency(ParamArray Currencys()) As Currency()
    Dim Arr@()
    ReDim Arr(LBound(Currencys) To UBound(Currencys))
    Dim i&
    For i = LBound(Currencys) To UBound(Currencys)
        #If VBA7 = 1 And Win64 = 1 Then
            If VarType(Currencys(i)) = vbLongLong Then
                Arr(i) = CCur(CDec(Currencys(i)))
            Else
                Arr(i) = CCur(Currencys(i))
            End If
        #Else
            Arr(i) = CCur(Currencys(i))
        #End If
    Next i
    ArrayCurrency = Arr
End Function

Public Function ArrayDate(ParamArray Dates()) As Date()
    Dim Arr() As Date
    ReDim Arr(LBound(Dates) To UBound(Dates))
    Dim i&
    For i = LBound(Dates) To UBound(Dates)
        Arr(i) = CDate(Dates(i))
    Next i
    ArrayDate = Arr
End Function

Public Function ArrayBoolean(ParamArray Booleans()) As Boolean()
    Dim Arr() As Boolean
    ReDim Arr(LBound(Booleans) To UBound(Booleans))
    Dim i&
    For i = LBound(Booleans) To UBound(Booleans)
        Arr(i) = CBool(Booleans(i))
    Next i
    ArrayBoolean = Arr
End Function

Public Function ArrayString(ParamArray Strings()) As String()
    Dim Arr$()
    ReDim Arr(LBound(Strings) To UBound(Strings))
    Dim i&
    For i = LBound(Strings) To UBound(Strings)
        Arr(i) = CStr(Strings(i))
    Next i
    ArrayString = Arr
End Function


'Concatenate===========================================================
'======================================================================

Public Sub ConcatArray(Arr, ConcatArr)
    If Not IsArray(Arr) Or Not IsArray(ConcatArr) Then
        Err.Raise 5
    End If
    Dim N1&
    Dim N2&
    N2 = SizeOfArray(ConcatArr)
    If N2 = 0 Then
        Exit Sub
    End If
    N1 = SizeOfArray(Arr)
    Dim L1&
    Dim U1&
    Dim L2&
    Dim U2&
    Dim C&
    If N1 > 0 Then
        L1 = LBound(Arr)
        U1 = UBound(Arr)
        C = U1
    Else
        C = -1
    End If
    L2 = LBound(ConcatArr)
    U2 = UBound(ConcatArr)
    ReDim Preserve Arr(L1 To L1 + N1 + N2 - 1)
    Dim i&
    For i = L2 To U2
        C = C + 1
        Arr(C) = ConcatArr(i)
    Next i
End Sub

Public Sub ConcatObjectArray(Arr, ConcatArr)
    If Not IsArray(Arr) Or Not IsArray(ConcatArr) Then
        Err.Raise 5
    End If
    Dim N1&
    Dim N2&
    N2 = SizeOfArray(ConcatArr)
    If N2 = 0 Then
        Exit Sub
    End If
    N1 = SizeOfArray(Arr)
    Dim L1&
    Dim U1&
    Dim L2&
    Dim U2&
    Dim C&
    Dim i&
    If N1 > 0 Then
        L1 = LBound(Arr)
        U1 = UBound(Arr)
        C = U1
        For i = L1 To U1
            If Not IsObject(Arr(i)) Then
                Err.Raise 5
            End If
        Next i
    Else
        C = -1
    End If
    L2 = LBound(ConcatArr)
    U2 = UBound(ConcatArr)
    For i = L2 To U2
        If Not IsObject(ConcatArr(i)) Then
            Err.Raise 5
        End If
    Next i
    ReDim Preserve Arr(L1 To L1 + N1 + N2 - 1)
    For i = L2 To U2
        C = C + 1
        Set Arr(C) = ConcatArr(i)
    Next i
End Sub


'Count=================================================================
'======================================================================

Public Function CountArray(Arr, Value) As Long
    If Not IsArray(Arr) Then
        Err.Raise 5
    End If
    Dim N&
    N = SizeOfArray(Arr)
    If N = 0 Then
        Exit Function
    End If
    Dim i&
    For i = LBound(Arr) To UBound(Arr)
        If Arr(i) = Value Then
            CountArray = CountArray + 1
        End If
    Next i
End Function

Public Function CountStringArray(Arr$(), Text$, _
Optional CompareMethod As VbCompareMethod = vbBinaryCompare) As Long
    Dim N&
    N = SizeOfArray(Arr)
    If N = 0 Then
        Exit Function
    End If
    Dim i&
    For i = LBound(Arr) To UBound(Arr)
        If StrComp(Arr(i), Text, CompareMethod) = 0 Then
            CountStringArray = CountStringArray + 1
        End If
    Next i
End Function

Public Function CountObjectArray(Arr, Value, Member$, _
Optional MemberIsMethod As Boolean = False) As Long
    If Not IsArray(Arr) Then
        Err.Raise 5
    End If
    Dim N&
    N = SizeOfArray(Arr)
    If N = 0 Then
        Exit Function
    End If
    Dim L&
    Dim U&
    L = LBound(Arr)
    U = UBound(Arr)
    Dim i&
    For i = L To U
        If Not IsObject(Arr(i)) Then
            Err.Raise 5
        End If
    Next i
    Dim CallType As VbCallType
    If MemberIsMethod Then
        CallType = VbMethod
    Else
        CallType = VbGet
    End If
    For i = L To U
        If CallByName(Arr(i), Member, CallType) = Value Then
            CountObjectArray = CountObjectArray + 1
        End If
    Next i
End Function


'Dimensions============================================================
'======================================================================

Public Function DimensionsOfArray&(Arr)
    If Not IsArray(Arr) Then
        Err.Raise 5
    End If
    Dim i&
    i = 1
    On Error GoTo Fail
    Do Until UBound(Arr, i) < LBound(Arr, i)
        DimensionsOfArray = i
        i = i + 1
    Loop
Fail:
    On Error GoTo 0
End Function


'Generate==============================================================
'======================================================================

Public Function GenerateLongArray(Start&, Length&, StepValue&, _
Optional LowerBound& = 0) As Long()
    If Length < 1 Then
        Err.Raise 5
    End If
    Dim Arr&()
    ReDim Arr(LowerBound To LowerBound + Length - 1)
    Dim j&
    j = Start
    Dim i&
    For i = LBound(Arr) To UBound(Arr)
        Arr(i) = j
        j = j + StepValue
    Next i
    GenerateLongArray = Arr
End Function

Public Function GenerateDoubleArray(Start#, Length&, StepValue#, _
Optional LowerBound& = 0) As Double()
    If Length < 1 Then
        Err.Raise 5
    End If
    Dim Arr#()
    ReDim Arr(LowerBound To LowerBound + Length - 1)
    Dim j#
    j = Start
    Dim i&
    For i = LBound(Arr) To UBound(Arr)
        Arr(i) = j
        j = j + StepValue
    Next i
    GenerateDoubleArray = Arr
End Function
    
Public Function GeneratePrimesArray(Start&, Length&, _
Optional LowerBound& = 0) As Long()
    If Length < 1 Then
        Err.Raise 5
    End If
    Dim Arr&()
    ReDim Arr(LowerBound To LowerBound + Length - 1)
    Dim i&
    Dim U&
    Dim j&
    i = LowerBound
    U = UBound(Arr) + 1
    j = Start
    Do While i < U
        If pIsPrime(j) Then
            Arr(i) = j
            i = i + 1
        End If
        j = j + 1
    Loop
    GeneratePrimesArray = Arr
End Function

Public Function GeneratePatternArray(PatternArr, Times&, _
Optional LowerBound& = 0)
    If Not IsArray(PatternArr) Then
        Err.Raise 5
    End If
    If Times < 1 Then
        Err.Raise 5
    End If
    Dim N&
    N = SizeOfArray(PatternArr)
    If N = 0 Then
        Err.Raise 5
    End If
    Dim Arr()
    ReDim Arr(LowerBound To LowerBound + N * Times - 1)
    Dim i&
    i = LowerBound
    Dim Fin&
    Fin = UBound(Arr) + 1
    Dim L&
    Dim U&
    L = LBound(PatternArr)
    U = UBound(PatternArr)
    Do While i < Fin
        Dim j&
        For j = L To U
            Arr(i) = PatternArr(j)
            i = i + 1
        Next j
    Loop
    GeneratePatternArray = Arr
End Function

Public Function GenerateFibonacciArray(N1&, N2&, Length&, _
Optional LowerBound& = 0) As Long()
    If Length < 2 Then
        Err.Raise 5
    End If
    Dim L&
    Dim U&
    L = LowerBound
    U = L + Length - 1
    Dim Arr&()
    ReDim Arr(L To U)
    Arr(L) = N1
    Arr(L + 1) = N2
    Dim i&
    For i = L + 2 To U
        Arr(i) = Arr(i - 2) + Arr(i - 1)
    Next i
    GenerateFibonacciArray = Arr
End Function

Public Function GenerateRndArray(Length&, _
Optional LowerBound& = 0) As Double()
    If Length < 1 Then
        Err.Raise 5
    End If
    Dim Arr#()
    ReDim Arr(LowerBound To LowerBound + Length - 1)
    Dim i&
    For i = LBound(Arr) To UBound(Arr)
        Randomize
        Arr(i) = Rnd
    Next i
    GenerateRndArray = Arr
End Function

Public Function GenerateRandomLongArray(MinValue&, MaxValue&, _
Length&, Optional LowerBound& = 0) As Long()
    If MinValue > MaxValue Then
        Err.Raise 5
    End If
    If Length < 1 Then
        Err.Raise 5
    End If
    Dim Arr&()
    ReDim Arr(LowerBound To LowerBound + Length - 1)
    Dim i&
    For i = LBound(Arr) To UBound(Arr)
        Arr(i) = pRandomLong(MinValue, MaxValue)
    Next i
    GenerateRandomLongArray = Arr
End Function

Public Function GenerateRandomDoubleArray(MinValue#, MaxValue#, _
Length&, Optional LowerBound& = 0) As Double()
    If MinValue > MaxValue Then
        Err.Raise 5
    End If
    If Length < 1 Then
        Err.Raise 5
    End If
    Dim Arr#()
    ReDim Arr(LowerBound To LowerBound + Length - 1)
    Dim i&
    For i = LBound(Arr) To UBound(Arr)
        Arr(i) = pRandomDouble(MinValue, MaxValue)
    Next i
    GenerateRandomDoubleArray = Arr
End Function


'Reverse===============================================================
'======================================================================

Public Sub ReverseArray(Arr)
    If Not IsArray(Arr) Then
        Err.Raise 5
    End If
    Dim N&
    N = SizeOfArray(Arr)
    If N = 0 Then
        Exit Sub
    End If
    Dim L&
    Dim R&
    L = LBound(Arr)
    R = UBound(Arr)
    Dim i&
    Dim Tmp
    For i = L To N \ 2
        Tmp = Arr(i)
        Arr(i) = Arr(R)
        Arr(R) = Tmp
        R = R - 1
    Next i
End Sub

Public Sub ReverseObjectArray(Arr)
    If Not IsArray(Arr) Then
        Err.Raise 5
    End If
    Dim N&
    N = SizeOfArray(Arr)
    If N = 0 Then
        Exit Sub
    End If
    Dim L&
    Dim R&
    L = LBound(Arr)
    R = UBound(Arr)
    Dim i&
    For i = L To R
        If Not IsObject(Arr(i)) Then
            Err.Raise 5
        End If
    Next i
    Dim Tmp
    For i = L To N \ 2
        Set Tmp = Arr(i)
        Set Arr(i) = Arr(R)
        Set Arr(R) = Tmp
        R = R - 1
    Next i
End Sub


'Search================================================================
'======================================================================

Public Function SearchArray(Arr, Value, _
Optional Direction As Boolean = True) As Long
    SearchArray = -1
    If Not IsArray(Arr) Then
        Err.Raise 5
    End If
    Dim N&
    N = SizeOfArray(Arr)
    If N = 0 Then
        Exit Function
    End If
    Dim i&
    Dim Beg&
    Dim Fin&
    Dim C&
    If Direction Then
        Beg = LBound(Arr)
        Fin = UBound(Arr)
        C = 1
    Else
        Beg = UBound(Arr)
        Fin = LBound(Arr)
        C = -1
    End If
    For i = Beg To Fin Step C
        If Arr(i) = Value Then
            SearchArray = i
            Exit Function
        End If
    Next i
End Function

Public Function SearchStringArray(Arr$(), Text$, _
CompareMethod As VbCompareMethod, _
Optional Direction As Boolean = True) As Long
    SearchStringArray = -1
    Dim N&
    N = SizeOfArray(Arr)
    If N = 0 Then
        Exit Function
    End If
    Dim i&
    Dim Beg&
    Dim Fin&
    Dim C&
    If Direction Then
        Beg = LBound(Arr)
        Fin = UBound(Arr)
        C = 1
    Else
        Beg = UBound(Arr)
        Fin = LBound(Arr)
        C = -1
    End If
    For i = Beg To Fin Step C
        If StrComp(Arr(i), Text, CompareMethod) = 0 Then
            SearchStringArray = i
            Exit Function
        End If
    Next i
End Function

Public Function SearchObjectArray(Arr, Value, Member$, _
Optional MemberIsMethod As Boolean = False, _
Optional Direction As Boolean = True) As Long
    SearchObjectArray = -1
    If Not IsArray(Arr) Then
        Err.Raise 5
    End If
    Dim N&
    N = SizeOfArray(Arr)
    If N = 0 Then
        Exit Function
    End If
    Dim CallType As VbCallType
    If MemberIsMethod Then
        CallType = VbMethod
    Else
        CallType = VbGet
    End If
    Dim i&
    Dim Beg&
    Dim Fin&
    Dim C&
    If Direction Then
        Beg = LBound(Arr)
        Fin = UBound(Arr)
        C = 1
    Else
        Beg = UBound(Arr)
        Fin = LBound(Arr)
        C = -1
    End If
    For i = Beg To Fin Step C
        If Not IsObject(Arr(i)) Then
            Err.Raise 5
        End If
    Next i
    For i = Beg To Fin Step C
        If CallByName(Arr(i), Member, CallType) = Value Then
            SearchObjectArray = i
            Exit Function
        End If
    Next i
End Function


'Shuffle===============================================================
'======================================================================

Public Sub ShuffleArray(Arr)
    If Not IsArray(Arr) Then
        Err.Raise 5
    End If
    Dim N&
    N = SizeOfArray(Arr)
    If N = 0 Then
        Exit Sub
    End If
    Dim L&
    Dim U&
    L = LBound(Arr)
    U = UBound(Arr)
    Dim C As Collection
    Set C = New Collection
    Dim i&
    For i = L To U
        C.Add Arr(i)
    Next i
    For i = L To U
        Dim j&
        j = pRandomLong(1, C.Count)
        Arr(i) = C.Item(j)
        C.Remove j
    Next i
    Set C = Nothing
End Sub

Public Sub ShuffleObjectArray(Arr)
    If Not IsArray(Arr) Then
        Err.Raise 5
    End If
    Dim N&
    N = SizeOfArray(Arr)
    If N = 0 Then
        Exit Sub
    End If
    Dim L&
    Dim U&
    L = LBound(Arr)
    U = UBound(Arr)
    Dim i&
    For i = L To U
        If Not IsObject(Arr(i)) Then
            Err.Raise 5
        End If
    Next i
    Dim C As Collection
    Set C = New Collection
    For i = L To U
        C.Add Arr(i)
    Next i
    For i = L To U
        Dim j&
        j = pRandomLong(1, C.Count)
        Set Arr(i) = C.Item(j)
        C.Remove j
    Next i
    Set C = Nothing
End Sub


'Array Size============================================================
'======================================================================

Public Function SizeOfArray(Arr) As Long
    If Not IsArray(Arr) Then
        Err.Raise 5
    End If
    On Error Resume Next
    SizeOfArray = UBound(Arr) - LBound(Arr) + 1
    On Error GoTo 0
End Function


'Sort==================================================================
'======================================================================

Public Sub SortArray(Arr)
    If Not IsArray(Arr) Then
        Err.Raise 5
    End If
    Dim N&
    N = SizeOfArray(Arr)
    If N = 0 Then
        Exit Sub
    End If
    If N < 51 Then
        pInsertionSortV Arr, N
    Else
        pMergeSortV Arr, N
    End If
End Sub

Public Sub SortStringArray(Arr$(), _
Optional CompareMethod As VbCompareMethod = vbBinaryCompare)
    Dim N As Long
    N = SizeOfArray(Arr)
    If N = 0 Then
        Exit Sub
    End If
    If N < 51 Then
        pInsertionSortS Arr, N, CompareMethod
    Else
        pMergeSortS Arr, N, CompareMethod
    End If
End Sub

Public Sub SortObjectArray(Arr, Member$, _
Optional MemberIsMethod As Boolean = False)
    If Not IsArray(Arr) Then
        Err.Raise 5
    End If
    Dim N&
    N = SizeOfArray(Arr)
    If N = 0 Then
        Exit Sub
    End If
    Dim i&
    For i = LBound(Arr) To UBound(Arr)
        If Not IsObject(Arr(i)) Then
            Err.Raise 5
        End If
    Next i
    Dim CallType As VbCallType
    If MemberIsMethod Then
        CallType = VbMethod
    Else
        CallType = VbGet
    End If
    If N < 51 Then
        pInsertionSortO Arr, N, Member, CallType
    Else
        pMergeSortO Arr, N, Member, CallType
    End If
End Sub
    

'Convert===============================================================
'======================================================================

Public Function ToByteArray(Arr) As Byte()
    If Not IsArray(Arr) Then
        Err.Raise 5
    End If
    Dim N&
    N = SizeOfArray(Arr)
    If N > 0 Then
        Dim L&
        Dim U&
        L = LBound(Arr)
        U = UBound(Arr)
        Dim BArr() As Byte
        ReDim BArr(L To U)
        Dim i&
        For i = L To U
            BArr(i) = CByte(Arr(i))
        Next i
    End If
    ToByteArray = BArr
End Function

Public Function ToIntegerArray(Arr) As Integer()
    If Not IsArray(Arr) Then
        Err.Raise 5
    End If
    Dim N&
    N = SizeOfArray(Arr)
    If N > 0 Then
        Dim L&
        Dim U&
        L = LBound(Arr)
        U = UBound(Arr)
        Dim IArr%()
        ReDim IArr(L To U)
        Dim i&
        For i = L To U
            IArr(i) = CInt(Arr(i))
        Next i
    End If
    ToIntegerArray = IArr
End Function

Public Function ToLongArray(Arr) As Long()
    If Not IsArray(Arr) Then
        Err.Raise 5
    End If
    Dim N&
    N = SizeOfArray(Arr)
    If N > 0 Then
        Dim L&
        Dim U&
        L = LBound(Arr)
        U = UBound(Arr)
        Dim LArr&()
        ReDim LArr(L To U)
        Dim i&
        For i = L To U
            LArr(i) = CLng(Arr(i))
        Next i
    End If
    ToLongArray = LArr
End Function

#If VBA7 = 1 And Win64 = 1 Then
Public Function ToLongLongArray(Arr) As LongLong()
    If Not IsArray(Arr) Then
        Err.Raise 5
    End If
    Dim N&
    N = SizeOfArray(Arr)
    If N > 0 Then
        Dim L&
        Dim U&
        L = LBound(Arr)
        U = UBound(Arr)
        Dim LLArr^()
        ReDim LLArr(L To U)
        Dim i&
        For i = L To U
            LLArr(i) = CLngLng(Arr(i))
        Next i
    End If
    ToLongLongArray = LLArr
End Function
#End If

#If VBA7 = 1 Then
Public Function ToLongPtrArray(Arr) As LongPtr()
    If Not IsArray(Arr) Then
        Err.Raise 5
    End If
    Dim N&
    N = SizeOfArray(Arr)
    If N > 0 Then
        Dim L&
        Dim U&
        L = LBound(Arr)
        U = UBound(Arr)
        Dim LPArr() As LongPtr
        ReDim LPArr(L To U)
        Dim i&
        For i = L To U
            LPArr(i) = CLngPtr(Arr(i))
        Next i
    End If
    ToLongPtrArray = LPArr
End Function
#End If

Public Function ToSingleArray(Arr) As Single()
    If Not IsArray(Arr) Then
        Err.Raise 5
    End If
    Dim N&
    N = SizeOfArray(Arr)
    If N > 0 Then
        Dim L&
        Dim U&
        L = LBound(Arr)
        U = UBound(Arr)
        Dim SArr!()
        ReDim SArr(L To U)
        Dim i&
        For i = L To U
            SArr(i) = CSng(Arr(i))
        Next i
    End If
    ToSingleArray = SArr
End Function

Public Function ToDoubleArray(Arr) As Double()
    If Not IsArray(Arr) Then
        Err.Raise 5
    End If
    Dim N&
    N = SizeOfArray(Arr)
    If N > 0 Then
        Dim L&
        Dim U&
        L = LBound(Arr)
        U = UBound(Arr)
        Dim DArr#()
        ReDim DArr(L To U)
        Dim i&
        For i = L To U
            DArr(i) = CDbl(Arr(i))
        Next i
    End If
    ToDoubleArray = DArr
End Function

Public Function ToDecimalArray(Arr) As Variant()
    If Not IsArray(Arr) Then
        Err.Raise 5
    End If
    Dim N&
    N = SizeOfArray(Arr)
    If N > 0 Then
        Dim L&
        Dim U&
        L = LBound(Arr)
        U = UBound(Arr)
        Dim DArr() As Variant
        ReDim DArr(L To U)
        Dim i&
        For i = L To U
            DArr(i) = CDec(Arr(i))
        Next i
    End If
    ToDecimalArray = DArr
End Function

Public Function ToCurrencyArray(Arr) As Currency()
    If Not IsArray(Arr) Then
        Err.Raise 5
    End If
    Dim N&
    N = SizeOfArray(Arr)
    If N > 0 Then
        Dim L&
        Dim U&
        L = LBound(Arr)
        U = UBound(Arr)
        Dim CArr@()
        ReDim CArr(L To U)
        Dim i&
        For i = L To U
            #If VBA7 = 1 And Win64 = 1 Then
            Dim Element
            Element = Arr(i)
            If VarType(Element) = vbLongLong Then
                Element = CDec(Element)
                CArr(i) = CCur(Element)
            Else
                CArr(i) = CCur(Element)
            End If
            #Else
                CArr(i) = CCur(Arr(i))
            #End If
        Next i
    End If
    ToCurrencyArray = CArr
End Function

Public Function ToDateArray(Arr) As Date()
    If Not IsArray(Arr) Then
        Err.Raise 5
    End If
    Dim N&
    N = SizeOfArray(Arr)
    If N > 0 Then
        Dim L&
        Dim U&
        L = LBound(Arr)
        U = UBound(Arr)
        Dim DArr() As Date
        ReDim DArr(L To U)
        Dim i&
        For i = L To U
            DArr(i) = CDate(Arr(i))
        Next i
    End If
    ToDateArray = DArr
End Function

Public Function ToBooleanArray(Arr) As Boolean()
    If Not IsArray(Arr) Then
        Err.Raise 5
    End If
    Dim N&
    N = SizeOfArray(Arr)
    If N > 0 Then
        Dim L&
        Dim U&
        L = LBound(Arr)
        U = UBound(Arr)
        Dim BArr() As Boolean
        ReDim BArr(L To U)
        Dim i&
        For i = L To U
            BArr(i) = CBool(Arr(i))
        Next i
    End If
    ToBooleanArray = BArr
End Function

Public Function ToStringArray(Arr) As String()
    If Not IsArray(Arr) Then
        Err.Raise 5
    End If
    Dim N&
    N = SizeOfArray(Arr)
    If N > 0 Then
        Dim L&
        Dim U&
        L = LBound(Arr)
        U = UBound(Arr)
        Dim SArr$()
        ReDim SArr(L To U)
        Dim i&
        For i = L To U
            SArr(i) = CStr(Arr(i))
        Next i
    End If
    ToStringArray = SArr
End Function

Public Function ToVariantArray(Arr) As Variant()
    If Not IsArray(Arr) Then
        Err.Raise 5
    End If
    Dim N&
    N = SizeOfArray(Arr)
    If N > 0 Then
        Dim L&
        Dim U&
        L = LBound(Arr)
        U = UBound(Arr)
        Dim VArr()
        ReDim VArr(L To U)
        Dim i&
        For i = L To U
            VArr(i) = CVar(Arr(i))
        Next i
    End If
    ToVariantArray = VArr
End Function


'Tests=================================================================
'======================================================================

Private Function TestmodArray() As Boolean

    TestmodArray = _
        TestTypedArrayFunctions And _
        TestConcatArray And _
        TestConcatObjectArray And _
        TestCountArray And _
        TestCountStringArray And _
        TestCountObjectArray And _
        TestDimensionsOfArray And _
        TestGenerateLongArray And _
        TestGenerateDoubleArray And _
        TestGeneratePrimesArray And _
        TestGeneratePatternArray And _
        TestGenerateFibonacciArray And _
        TestGenerateRndArray And _
        TestGenerateRandomLongArray And _
        TestGenerateRandomDoubleArray And _
        TestReverseArray And _
        TestReverseObjectArray And _
        TestSearchArray And _
        TestSearchStringArray And _
        TestSearchObjectArray And _
        TestShuffleArray And _
        TestShuffleObjectArray And _
        TestSizeOfArray
    TestmodArray = TestmodArray And _
        TestSortArray And _
        TestInsertionSortV And _
        TestMergeSortV And _
        TestSortStringArray And _
        TestInsertionSortS And _
        TestMergeSortS And _
        TestSortObjectArray And _
        TestInsertionSortO And _
        TestMergeSortO And _
        TestIsPrime And _
        TestRandomLong And _
        TestRandomDouble
    TestmodArray = TestmodArray And _
        TestToByteArray And _
        TestToIntegerArray And _
        TestToLongArray And _
        TestToSingleArray And _
        TestToDoubleArray And _
        TestToDecimalArray And _
        TestToCurrencyArray And _
        TestToDateArray And _
        TestToBooleanArray And _
        TestToStringArray And _
        TestToVariantArray
        
    #If VBA7 = 1 And Win64 = 1 Then
        TestmodArray = TestmodArray And TestToLongLongArray
    #End If
    
    #If VBA7 = 1 Then
        TestmodArray = TestmodArray And TestToLongPtrArray
    #End If
        
    Debug.Print "TestmodArray: " & TestmodArray

End Function

Private Function TestTypedArrayFunctions() As Boolean
    
    TestTypedArrayFunctions = True
    
    'Byte
    Dim ByteArr() As Byte
    On Error Resume Next
    ByteArr = ArrayByte(1, 2, 3)
    If Err.Number <> 0 Then
        TestTypedArrayFunctions = False
        Debug.Print "Assign Byte"
    End If
    On Error GoTo 0
    If SizeOfArray(ByteArr) <> 3 Then
        TestTypedArrayFunctions = False
        Debug.Print "Byte Array Size"
    End If
    If ByteArr(0) <> 1 Then
        TestTypedArrayFunctions = False
        Debug.Print "Byte Value"
    End If
    
    
    'Integer
    Dim IntArr() As Integer
    On Error Resume Next
    IntArr = ArrayInteger(1, 2, 3)
    If Err.Number <> 0 Then
        TestTypedArrayFunctions = False
        Debug.Print "Assign Integer"
    End If
    On Error GoTo 0
    If SizeOfArray(IntArr) <> 3 Then
        TestTypedArrayFunctions = False
        Debug.Print "Int Array Size"
    End If
    If IntArr(0) <> 1 Then
        TestTypedArrayFunctions = False
        Debug.Print "Int Value"
    End If
    
    'Long
    Dim LngArr() As Long
    On Error Resume Next
    LngArr = ArrayLong(1, 2, 3)
    If Err.Number <> 0 Then
        TestTypedArrayFunctions = False
        Debug.Print "Assign Long"
    End If
    On Error GoTo 0
    If SizeOfArray(LngArr) <> 3 Then
        TestTypedArrayFunctions = False
        Debug.Print "Long Array Size"
    End If
    If LngArr(0) <> 1 Then
        TestTypedArrayFunctions = False
        Debug.Print "Long Value"
    End If
    
    'LongLong
    #If VBA7 = 1 And Win64 = 1 Then
    Dim LngLngArr() As LongLong
    On Error Resume Next
    LngLngArr = ArrayLongLong(1, 2, 3)
    If Err.Number <> 0 Then
        TestTypedArrayFunctions = False
        Debug.Print "Assign LongLong"
    End If
    On Error GoTo 0
    If SizeOfArray(LngLngArr) <> 3 Then
        TestTypedArrayFunctions = False
        Debug.Print "LongLong Array Size"
    End If
    If LngLngArr(0) <> 1 Then
        TestTypedArrayFunctions = False
        Debug.Print "LongLong Value"
    End If
    #End If
    
    'LongPtr
    #If VBA7 = 1 Then
    Dim LngPtrArr() As LongPtr
    On Error Resume Next
    LngPtrArr = ArrayLongPtr(1, 2, 3)
    If Err.Number <> 0 Then
        TestTypedArrayFunctions = False
        Debug.Print "Assign LongPtr"
    End If
    On Error GoTo 0
    If SizeOfArray(LngPtrArr) <> 3 Then
        TestTypedArrayFunctions = False
        Debug.Print "LongPtr Array Size"
    End If
    If LngPtrArr(0) <> 1 Then
        TestTypedArrayFunctions = False
        Debug.Print "LongPtr Value"
    End If
    #End If
    
    'Single
    Dim SngArr() As Single
    On Error Resume Next
    SngArr = ArraySingle(1.5, 2.5, 3.5)
    If Err.Number <> 0 Then
        TestTypedArrayFunctions = False
        Debug.Print "Assign Single"
    End If
    On Error GoTo 0
    If SizeOfArray(SngArr) <> 3 Then
        TestTypedArrayFunctions = False
        Debug.Print "Single Array Size"
    End If
    If SngArr(0) <> 1.5 Then
        TestTypedArrayFunctions = False
        Debug.Print "Single Value"
    End If
    
    'Double
    Dim DblArr() As Double
    On Error Resume Next
    DblArr = ArrayDouble(1.5, 2.5, 3.5)
    If Err.Number <> 0 Then
        TestTypedArrayFunctions = False
        Debug.Print "Assign Double"
    End If
    On Error GoTo 0
    If SizeOfArray(DblArr) <> 3 Then
        TestTypedArrayFunctions = False
        Debug.Print "Double Array Size"
    End If
    If DblArr(0) <> 1.5 Then
        TestTypedArrayFunctions = False
        Debug.Print "Double Value"
    End If
    
    'Decimal
    Dim DecArr() As Variant
    On Error Resume Next
    DecArr = ArrayDecimal(1.5, 2.5, 3.5)
    If Err.Number <> 0 Then
        TestTypedArrayFunctions = False
        Debug.Print "Assign Decimal"
    End If
    On Error GoTo 0
    If SizeOfArray(DecArr) <> 3 Then
        TestTypedArrayFunctions = False
        Debug.Print "Decimal Array Size"
    End If
    If DecArr(0) <> 1.5 Then
        TestTypedArrayFunctions = False
        Debug.Print "Decimal Value"
    End If
    
    'Currency
    Dim CurArr() As Currency
    On Error Resume Next
    CurArr = ArrayCurrency(1.5, 2.5, 3.5)
    If Err.Number <> 0 Then
        TestTypedArrayFunctions = False
        Debug.Print "Assign Currency"
    End If
    On Error GoTo 0
    If SizeOfArray(CurArr) <> 3 Then
        TestTypedArrayFunctions = False
        Debug.Print "Currency Array Size"
    End If
    If CurArr(0) <> 1.5 Then
        TestTypedArrayFunctions = False
        Debug.Print "Currency Value"
    End If
    
    'Date
    Dim DateArr() As Date
    On Error Resume Next
    DateArr = ArrayDate("01/01/2021", #1/1/2022#, #1/1/2023#)
    If Err.Number <> 0 Then
        TestTypedArrayFunctions = False
        Debug.Print "Assign Date"
    End If
    On Error GoTo 0
    If SizeOfArray(DateArr) <> 3 Then
        TestTypedArrayFunctions = False
        Debug.Print "Date Array Size"
    End If
    If DateArr(0) <> #1/1/2021# Then
        TestTypedArrayFunctions = False
        Debug.Print "Date Value"
    End If
    
    'Boolean
    Dim BoolArr() As Boolean
    On Error Resume Next
    BoolArr = ArrayBoolean(True, False, True)
    If Err.Number <> 0 Then
        TestTypedArrayFunctions = False
        Debug.Print "Assign Boolean"
    End If
    On Error GoTo 0
    If SizeOfArray(BoolArr) <> 3 Then
        TestTypedArrayFunctions = False
        Debug.Print "Boolean Array Size"
    End If
    If BoolArr(0) <> True Then
        TestTypedArrayFunctions = False
        Debug.Print "Bool Value"
    End If
    
    'String
    Dim StrArr() As String
    On Error Resume Next
    StrArr = ArrayString("A", "B", "C")
    If Err.Number <> 0 Then
        TestTypedArrayFunctions = False
        Debug.Print "Assign String"
    End If
    On Error GoTo 0
    If SizeOfArray(StrArr) <> 3 Then
        TestTypedArrayFunctions = False
        Debug.Print "String Array Size"
    End If
    If StrArr(0) <> "A" Then
        TestTypedArrayFunctions = False
        Debug.Print "Str Value"
    End If
    
    Debug.Print "TestTypedArrayFunctions: " & TestTypedArrayFunctions
    
End Function

Private Function TestConcatArray() As Boolean
    
    TestConcatArray = True
    
    Dim Arr1() As Long
    Dim Arr2() As Long
    
    'Empty Empty
    On Error Resume Next
    ConcatArray Arr1, Arr2
    If Err.Number <> 0 Then
        TestConcatArray = False
        Debug.Print "Empty Empty"
    End If
    On Error GoTo 0
    
    'Empty NotEmpty
    ReDim Arr2(0 To 2)
    On Error Resume Next
    ConcatArray Arr1, Arr2
    If Err.Number <> 0 Then
        TestConcatArray = False
        Debug.Print "Empty NotEmpty"
    End If
    On Error GoTo 0
    
    'NotEmpty Empty
    Erase Arr2
    ReDim Arr1(0 To 2)
    On Error Resume Next
    ConcatArray Arr1, Arr2
    If Err.Number <> 0 Then
        TestConcatArray = False
        Debug.Print "NotEmpty Empty"
    End If
    On Error GoTo 0
    
    'One
    ReDim Arr1(0 To 0)
    ReDim Arr2(0 To 0)
    ConcatArray Arr1, Arr2
    If UBound(Arr1) - LBound(Arr1) + 1 <> 2 Or _
    Arr1(0) <> 0 Or _
    Arr1(1) <> 0 Then
        TestConcatArray = False
        Debug.Print "One"
    End If
    
    'Multiple
    ReDim Arr1(0 To 2)
    ReDim Arr2(0 To 2)
    ConcatArray Arr1, Arr2
    If UBound(Arr1) - LBound(Arr1) + 1 <> 6 Then
        TestConcatArray = False
        Debug.Print "One"
    End If
    
    'LBound 1 0
    ReDim Arr1(1 To 2)
    ReDim Arr2(0 To 2)
    ConcatArray Arr1, Arr2
    If UBound(Arr1) - LBound(Arr1) + 1 <> 5 Then
        TestConcatArray = False
        Debug.Print "LBound 1 0 Size"
    End If
    If LBound(Arr1) <> 1 Or UBound(Arr1) <> 5 Then
        TestConcatArray = False
        Debug.Print "LBound 1 0 Bounds"
    End If
    
    'LBound 0 1
    ReDim Arr1(0 To 2)
    ReDim Arr2(1 To 2)
    ConcatArray Arr1, Arr2
    If UBound(Arr1) - LBound(Arr1) + 1 <> 5 Then
        TestConcatArray = False
        Debug.Print "LBound 0 1 Size"
    End If
    If LBound(Arr1) <> 0 Or UBound(Arr1) <> 4 Then
        TestConcatArray = False
        Debug.Print "LBound 0 1 Bounds"
    End If
    
    'LBound 1 1
    ReDim Arr1(1 To 2)
    ReDim Arr2(1 To 2)
    ConcatArray Arr1, Arr2
    If UBound(Arr1) - LBound(Arr1) + 1 <> 4 Then
        TestConcatArray = False
        Debug.Print "LBound 0 0 Size"
    End If
    If LBound(Arr1) <> 1 Or UBound(Arr1) <> 4 Then
        TestConcatArray = False
        Debug.Print "LBound 1 1 Bounds"
    End If
    
    Debug.Print "TestConcatArray: " & TestConcatArray
    
End Function

Private Function TestConcatObjectArray() As Boolean

    TestConcatObjectArray = True
    
    Dim Arr1() As Object
    Dim Arr2() As Object
    
    'Not Object
    Dim NOArr1(0 To 0)
    Dim NOArr2(0 To 0)
    NOArr1(0) = 1
    Set NOArr2(0) = New Collection
    On Error Resume Next
    ConcatObjectArray NOArr1, NOArr2
    If Err.Number <> 5 Then
        TestConcatObjectArray = False
        Debug.Print "Not Object 1"
    End If
    On Error GoTo 0
    Set NOArr1(0) = New Collection
    NOArr2(0) = 1
    On Error Resume Next
    ConcatObjectArray NOArr1, NOArr2
    If Err.Number <> 5 Then
        TestConcatObjectArray = False
        Debug.Print "Not Object 2"
    End If
    On Error GoTo 0
    On Error GoTo 0
    NOArr1(0) = 1
    NOArr2(0) = 1
    On Error Resume Next
    ConcatObjectArray NOArr1, NOArr2
    If Err.Number <> 5 Then
        TestConcatObjectArray = False
        Debug.Print "Not Object both"
    End If
    On Error GoTo 0
    
    'Empty Empty
    On Error Resume Next
    ConcatObjectArray Arr1, Arr2
    If Err.Number <> 0 Then
        TestConcatObjectArray = False
        Debug.Print "Empty Empty"
    End If
    On Error GoTo 0
    
    'Empty NotEmpty
    ReDim Arr2(0 To 2)
    On Error Resume Next
    ConcatObjectArray Arr1, Arr2
    If Err.Number <> 0 Then
        TestConcatObjectArray = False
        Debug.Print "Empty NotEmpty"
    End If
    On Error GoTo 0
    
    'NotEmpty Empty
    Erase Arr2
    ReDim Arr1(0 To 2)
    On Error Resume Next
    ConcatObjectArray Arr1, Arr2
    If Err.Number <> 0 Then
        TestConcatObjectArray = False
        Debug.Print "NotEmpty Empty"
    End If
    On Error GoTo 0
    
    'One
    ReDim Arr1(0 To 0)
    ReDim Arr2(0 To 0)
    ConcatObjectArray Arr1, Arr2
    If UBound(Arr1) - LBound(Arr1) + 1 <> 2 Or _
    Not Arr1(0) Is Nothing Or _
    Not Arr1(1) Is Nothing Then
        TestConcatObjectArray = False
        Debug.Print "One"
    End If
    
    'Multiple
    ReDim Arr1(0 To 2)
    ReDim Arr2(0 To 2)
    ConcatObjectArray Arr1, Arr2
    If UBound(Arr1) - LBound(Arr1) + 1 <> 6 Then
        TestConcatObjectArray = False
        Debug.Print "One"
    End If
    
    'LBound 1 0
    ReDim Arr1(1 To 2)
    ReDim Arr2(0 To 2)
    ConcatObjectArray Arr1, Arr2
    If UBound(Arr1) - LBound(Arr1) + 1 <> 5 Then
        TestConcatObjectArray = False
        Debug.Print "LBound 1 0 Size"
    End If
    If LBound(Arr1) <> 1 Or UBound(Arr1) <> 5 Then
        TestConcatObjectArray = False
        Debug.Print "LBound 1 0 Bounds"
    End If
    
    'LBound 0 1
    ReDim Arr1(0 To 2)
    ReDim Arr2(1 To 2)
    ConcatObjectArray Arr1, Arr2
    If UBound(Arr1) - LBound(Arr1) + 1 <> 5 Then
        TestConcatObjectArray = False
        Debug.Print "LBound 0 1 Size"
    End If
    If LBound(Arr1) <> 0 Or UBound(Arr1) <> 4 Then
        TestConcatObjectArray = False
        Debug.Print "LBound 0 1 Bounds"
    End If
    
    'LBound 1 1
    ReDim Arr1(1 To 2)
    ReDim Arr2(1 To 2)
    ConcatObjectArray Arr1, Arr2
    If UBound(Arr1) - LBound(Arr1) + 1 <> 4 Then
        TestConcatObjectArray = False
        Debug.Print "LBound 1 1 Size"
    End If
    If LBound(Arr1) <> 1 Or UBound(Arr1) <> 4 Then
        TestConcatObjectArray = False
        Debug.Print "LBound 1 1 Bounds"
    End If
    
    Debug.Print "TestConcatObjectArray: " & TestConcatObjectArray
    
End Function

Private Function TestCountArray() As Boolean
    
    TestCountArray = True
    
    Dim Arr() As Long
    
    'Empty
    Dim i As Long
    On Error Resume Next
    i = CountArray(Arr, 0)
    If Err.Number <> 0 Or i <> 0 Then
        TestCountArray = False
        Debug.Print "Empty"
    End If
    On Error GoTo 0
    
    'One
    ReDim Arr(0 To 0)
        'Not there
        If CountArray(Arr, 1) <> 0 Then
            TestCountArray = False
            Debug.Print "One Not there"
        End If
        'There
        If CountArray(Arr, 0) <> 1 Then
            TestCountArray = False
            Debug.Print "One There"
        End If

    'Multiple
    ReDim Arr(0 To 2)
        'Not there
        If CountArray(Arr, 1) <> 0 Then
            TestCountArray = False
            Debug.Print "Multiple Not there"
        End If
        'There
        Arr(2) = 1
        If CountArray(Arr, 0) <> 2 Then
            TestCountArray = False
            Debug.Print "Multiple There"
        End If
    
    Debug.Print "TestCountArray: " & TestCountArray
    
End Function

Private Function TestCountStringArray() As Boolean

    TestCountStringArray = True
    
    Dim Arr() As String
    
    'Empty
    Dim i As Long
    On Error Resume Next
    i = CountStringArray(Arr, "A", vbBinaryCompare)
    If Err.Number <> 0 Or i <> 0 Then
        TestCountStringArray = False
        Debug.Print "Empty"
    End If
    On Error GoTo 0
    
    'One
    ReDim Arr(0 To 0)
        'Not there
        If CountStringArray(Arr, "A", vbBinaryCompare) <> 0 Then
            TestCountStringArray = False
            Debug.Print "One Not there"
        End If
        'There
        Arr(0) = "A"
        If CountStringArray(Arr, "A", vbBinaryCompare) <> 1 Then
            TestCountStringArray = False
            Debug.Print "One There"
        End If

    'Multiple
    ReDim Arr(0 To 2)
        'Not there
        If CountStringArray(Arr, "A", vbBinaryCompare) <> 0 Then
            TestCountStringArray = False
            Debug.Print "Multiple Not there"
        End If
        'There
        Arr(0) = "A"
        Arr(1) = "A"
        Arr(2) = "B"
        If CountStringArray(Arr, "A", vbBinaryCompare) <> 2 Then
            TestCountStringArray = False
            Debug.Print "Multiple There"
        End If
    
    'Compare Text
    ReDim Arr(0 To 2)
        'Not there
        If CountStringArray(Arr, "A", vbTextCompare) <> 0 Then
            TestCountStringArray = False
            Debug.Print "Multiple Not there Compare Text"
        End If
        'There
        Arr(0) = "A"
        Arr(1) = "a"
        Arr(2) = "B"
        If CountStringArray(Arr, "A", vbTextCompare) <> 2 Then
            TestCountStringArray = False
            Debug.Print "Multiple There Compare Text"
        End If
        
    Debug.Print "TestCountStringArray: " & TestCountStringArray
        
End Function

Private Function TestCountObjectArray() As Boolean

    TestCountObjectArray = True
    
    Dim Arr() As Object
    
    Dim i As Long
    
    'Not Object
    Dim NOArr(0 To 0)
    NOArr(0) = 1
    On Error Resume Next
    i = CountObjectArray(NOArr, 0, "Count", True)
    If Err.Number <> 5 Then
        TestCountObjectArray = False
        Debug.Print "Not Object"
    End If
    On Error GoTo 0
    
    'Empty
    On Error Resume Next
    i = CountObjectArray(Arr, 0, "Count", True)
    If Err.Number <> 0 Or i <> 0 Then
        TestCountObjectArray = False
        Debug.Print "Empty"
    End If
    On Error GoTo 0
    
    'One
    ReDim Arr(0 To 0)
    Set Arr(0) = New Collection
        'Not there
        If CountObjectArray(Arr, 1, "Count", True) <> 0 Then
            TestCountObjectArray = False
            Debug.Print "One Not there"
        End If
        'There
        Arr(0).Add 1
        If CountObjectArray(Arr, 1, "Count", True) <> 1 Then
            TestCountObjectArray = False
            Debug.Print "One There"
        End If

    'Multiple
    ReDim Arr(0 To 2)
    Set Arr(0) = New Collection
    Set Arr(1) = New Collection
    Set Arr(2) = New Collection
        'Not there
        If CountObjectArray(Arr, 1, "Count", True) <> 0 Then
            TestCountObjectArray = False
            Debug.Print "Multiple Not there"
        End If
        'There
        Arr(0).Add 1
        Arr(1).Add 1
        If CountObjectArray(Arr, 1, "Count", True) <> 2 Then
            TestCountObjectArray = False
            Debug.Print "Multiple There"
        End If
    
    Debug.Print "TestCountObjectArray: " & TestCountObjectArray
        
End Function

Private Function TestDimensionsOfArray() As Boolean
    
    TestDimensionsOfArray = True
    
    Dim LArr() As Long
    Dim VArr() As Variant
    Dim OArr() As Object
    
    Dim D As Long
    
    'Long Array
    On Error Resume Next
    D = DimensionsOfArray(LArr)
    If Err.Number <> 0 Or D <> 0 Then
        TestDimensionsOfArray = False
        Debug.Print "Uninitialized Long Array"
    End If
    On Error GoTo 0
    ReDim LArr(0 To 0)
    D = DimensionsOfArray(LArr)
    If D <> 1 Then
        TestDimensionsOfArray = False
        Debug.Print "1 Dimension Long Array"
    End If
    ReDim LArr(0 To 0, 0 To 0)
    D = DimensionsOfArray(LArr)
    If D <> 2 Then
        TestDimensionsOfArray = False
        Debug.Print "2 Dimension Long Array"
    End If
    
    'Variant Array
    On Error Resume Next
    D = DimensionsOfArray(VArr)
    If Err.Number <> 0 Or D <> 0 Then
        TestDimensionsOfArray = False
        Debug.Print "Uninitialized Variant Array"
    End If
    On Error GoTo 0
    ReDim VArr(0 To 0)
    D = DimensionsOfArray(VArr)
    If D <> 1 Then
        TestDimensionsOfArray = False
        Debug.Print "1 Dimension Variant Array"
    End If
    ReDim VArr(0 To 0, 0 To 0)
    D = DimensionsOfArray(VArr)
    If D <> 2 Then
        TestDimensionsOfArray = False
        Debug.Print "2 Dimension Variant Array"
    End If
    
    'Object Array
    On Error Resume Next
    D = DimensionsOfArray(OArr)
    If Err.Number <> 0 Or D <> 0 Then
        TestDimensionsOfArray = False
        Debug.Print "Uninitialized Object Array"
    End If
    On Error GoTo 0
    ReDim OArr(0 To 0)
    D = DimensionsOfArray(OArr)
    If D <> 1 Then
        TestDimensionsOfArray = False
        Debug.Print "1 Dimension Object Array"
    End If
    ReDim OArr(0 To 0, 0 To 0)
    D = DimensionsOfArray(OArr)
    If D <> 2 Then
        TestDimensionsOfArray = False
        Debug.Print "2 Dimension Object Array"
    End If
    
    Debug.Print "TestDimensionsOfArray: " & TestDimensionsOfArray

End Function

Private Function TestGenerateLongArray() As Boolean

    TestGenerateLongArray = True

    Dim Arr() As Long

    '-1
    On Error Resume Next
    Arr = GenerateLongArray(1, -1, 1)
    If Err.Number <> 5 Then
        TestGenerateLongArray = False
        Debug.Print "(1, -1, 1)"
    End If
    On Error GoTo 0

    '0
    On Error Resume Next
    Arr = GenerateLongArray(1, 0, 1)
    If Err.Number <> 5 Then
        TestGenerateLongArray = False
        Debug.Print "(1, 0, 1)"
    End If
    On Error GoTo 0
    
    '1 To 1 Step 0
    Arr = GenerateLongArray(1, 10, 0)
    If Arr(LBound(Arr)) <> 1 Then
        TestGenerateLongArray = False
        Debug.Print "LBound Value (1, 10, 0)"
    End If
    If Arr(UBound(Arr)) <> 1 Then
        TestGenerateLongArray = False
        Debug.Print "UBound Value (1, 10, 0)"
    End If
    If LBound(Arr) <> 0 Then
        TestGenerateLongArray = False
        Debug.Print "LBound Index (1, 10, 0)"
    End If
    If UBound(Arr) <> 9 Then
        TestGenerateLongArray = False
        Debug.Print "UBound Index (1, 10, 0)"
    End If

    '1 To 10 Step 1
    Arr = GenerateLongArray(1, 10, 1)
    If Arr(LBound(Arr)) <> 1 Then
        TestGenerateLongArray = False
        Debug.Print "LBound Value (1, 10, 1)"
    End If
    If Arr(UBound(Arr)) <> 10 Then
        TestGenerateLongArray = False
        Debug.Print "UBound Value (1, 10, 1)"
    End If
    If LBound(Arr) <> 0 Then
        TestGenerateLongArray = False
        Debug.Print "LBound Index (1, 10, 1)"
    End If
    If UBound(Arr) <> 9 Then
        TestGenerateLongArray = False
        Debug.Print "UBound Index (1, 10, 1)"
    End If

    '0 To 9 Step 1
    Arr = GenerateLongArray(0, 10, 1)
    If Arr(LBound(Arr)) <> 0 Then
        TestGenerateLongArray = False
        Debug.Print "LBound Value (0, 10, 1)"
    End If
    If Arr(UBound(Arr)) <> 9 Then
        TestGenerateLongArray = False
        Debug.Print "UBound Value (0, 10, 1)"
    End If
    If LBound(Arr) <> 0 Then
        TestGenerateLongArray = False
        Debug.Print "LBound Index (0, 10, 1)"
    End If
    If UBound(Arr) <> 9 Then
        TestGenerateLongArray = False
        Debug.Print "UBound Index (0, 10, 1)"
    End If

    '1 To 19 Step 2
    Arr = GenerateLongArray(1, 10, 2)
    If Arr(LBound(Arr)) <> 1 Then
        TestGenerateLongArray = False
        Debug.Print "LBound Value (1, 10, 2)"
    End If
    If Arr(UBound(Arr)) <> 19 Then
        TestGenerateLongArray = False
        Debug.Print "UBound Value (1, 10, 2)"
    End If
    If LBound(Arr) <> 0 Then
        TestGenerateLongArray = False
        Debug.Print "LBound Index (1, 10, 2)"
    End If
    If UBound(Arr) <> 9 Then
        TestGenerateLongArray = False
        Debug.Print "UBound Index (1, 10, 2)"
    End If

    '10 To 1 Step -1
    Arr = GenerateLongArray(10, 10, -1)
    If Arr(LBound(Arr)) <> 10 Then
        TestGenerateLongArray = False
        Debug.Print "LBound Value (10, 10, -1)"
    End If
    If Arr(UBound(Arr)) <> 1 Then
        TestGenerateLongArray = False
        Debug.Print "UBound Value (10, 10, -1)"
    End If
    If LBound(Arr) <> 0 Then
        TestGenerateLongArray = False
        Debug.Print "LBound Index (10, 10, -1)"
    End If
    If UBound(Arr) <> 9 Then
        TestGenerateLongArray = False
        Debug.Print "UBound Index (10, 10, -1)"
    End If

    '1 To 10 Step 1 LBound = 1
    Arr = GenerateLongArray(1, 10, 1, 1)
    If Arr(LBound(Arr)) <> 1 Then
        TestGenerateLongArray = False
        Debug.Print "LBound (1, 10, 1, 1)"
    End If
    If Arr(UBound(Arr)) <> 10 Then
        TestGenerateLongArray = False
        Debug.Print "UBound (1, 10, 1, 1)"
    End If
    If LBound(Arr) <> 1 Then
        TestGenerateLongArray = False
        Debug.Print "LBound Index (1, 10, 1, 1)"
    End If
    If UBound(Arr) <> 10 Then
        TestGenerateLongArray = False
        Debug.Print "UBound Index (1, 10, 1, 1)"
    End If

    '1 To 10 Step 1 LBound = -10
    Arr = GenerateLongArray(1, 10, 1, -10)
    If Arr(LBound(Arr)) <> 1 Then
        TestGenerateLongArray = False
        Debug.Print "LBound Value (1, 10, 1, -10)"
    End If
    If Arr(UBound(Arr)) <> 10 Then
        TestGenerateLongArray = False
        Debug.Print "UBound Value (1, 10, 1, -10)"
    End If
    If LBound(Arr) <> -10 Then
        TestGenerateLongArray = False
        Debug.Print "LBound Index (1, 10, 1, -10)"
    End If
    If UBound(Arr) <> -1 Then
        TestGenerateLongArray = False
        Debug.Print "UBound Index (1, 10, 1, -10)"
    End If

    Debug.Print "TestGenerateLongArray: " & TestGenerateLongArray

End Function

Private Function TestGenerateDoubleArray() As Boolean

    TestGenerateDoubleArray = True

    Dim Arr() As Double

    '-1
    On Error Resume Next
    Arr = GenerateDoubleArray(1, -1, 1)
    If Err.Number <> 5 Then
        TestGenerateDoubleArray = False
        Debug.Print "(1, -1, 1)"
    End If
    On Error GoTo 0

    '0
    On Error Resume Next
    Arr = GenerateDoubleArray(1, 0, 1)
    If Err.Number <> 5 Then
        TestGenerateDoubleArray = False
        Debug.Print "(1, 0, 1)"
    End If
    On Error GoTo 0
    
    '3.5 To 3.5 Step 0
    Arr = GenerateDoubleArray(3.5, 10, 0)
    If Arr(LBound(Arr)) <> 3.5 Then
        TestGenerateDoubleArray = False
        Debug.Print "LBound Value (3.5, 10, 0)"
    End If
    If Arr(UBound(Arr)) <> 3.5 Then
        TestGenerateDoubleArray = False
        Debug.Print "UBound Value (3.5, 10, 0)"
    End If
    If LBound(Arr) <> 0 Then
        TestGenerateDoubleArray = False
        Debug.Print "LBound Index (3.5, 10, 0)"
    End If
    If UBound(Arr) <> 9 Then
        TestGenerateDoubleArray = False
        Debug.Print "UBound Index (3.5, 10, 0)"
    End If

    '3.5 To 8 Step 0.5
    Arr = GenerateDoubleArray(3.5, 10, 0.5)
    If Arr(LBound(Arr)) <> 3.5 Then
        TestGenerateDoubleArray = False
        Debug.Print "LBound Value (3.5, 10, 0.5)"
    End If
    If Arr(UBound(Arr)) <> 8 Then
        TestGenerateDoubleArray = False
        Debug.Print "UBound Value (3.5, 10, 0.5)"
    End If
    If LBound(Arr) <> 0 Then
        TestGenerateDoubleArray = False
        Debug.Print "LBound Index (3.5, 10, 0.5)"
    End If
    If UBound(Arr) <> 9 Then
        TestGenerateDoubleArray = False
        Debug.Print "UBound Index (3.5, 10, 0.5)"
    End If

    '2.5 To 7 Step 0.5
    Arr = GenerateDoubleArray(2.5, 10, 0.5)
    If Arr(LBound(Arr)) <> 2.5 Then
        TestGenerateDoubleArray = False
        Debug.Print "LBound Value (2.5, 10, 0.5)"
    End If
    If Arr(UBound(Arr)) <> 7 Then
        TestGenerateDoubleArray = False
        Debug.Print "UBound Value (2.5, 10, 0.5)"
    End If
    If LBound(Arr) <> 0 Then
        TestGenerateDoubleArray = False
        Debug.Print "LBound Index (2.5, 10, 0.5)"
    End If
    If UBound(Arr) <> 9 Then
        TestGenerateDoubleArray = False
        Debug.Print "UBound Index (2.5, 10, 0.5)"
    End If

    '0.5 To 9.5 Step 1
    Arr = GenerateDoubleArray(0.5, 10, 1)
    If Arr(LBound(Arr)) <> 0.5 Then
        TestGenerateDoubleArray = False
        Debug.Print "LBound Value (0.5, 10, 1)"
    End If
    If Arr(UBound(Arr)) <> 9.5 Then
        TestGenerateDoubleArray = False
        Debug.Print "UBound Value (0.5, 10, 1)"
    End If
    If LBound(Arr) <> 0 Then
        TestGenerateDoubleArray = False
        Debug.Print "LBound Index (0.5, 10, 1)"
    End If
    If UBound(Arr) <> 9 Then
        TestGenerateDoubleArray = False
        Debug.Print "UBound Index (0.5, 10, 1)"
    End If

    '8 To 3.5 Step -0.5
    Arr = GenerateDoubleArray(8, 10, -0.5)
    If Arr(LBound(Arr)) <> 8 Then
        TestGenerateDoubleArray = False
        Debug.Print "LBound Value (8, 10, -0.5)"
    End If
    If Arr(UBound(Arr)) <> 3.5 Then
        TestGenerateDoubleArray = False
        Debug.Print "UBound Value (8, 10, -0.5)"
    End If
    If LBound(Arr) <> 0 Then
        TestGenerateDoubleArray = False
        Debug.Print "LBound Index (8, 10, -0.5)"
    End If
    If UBound(Arr) <> 9 Then
        TestGenerateDoubleArray = False
        Debug.Print "UBound Index (8, 10, -0.5)"
    End If

    '3.5 To 8 Step 0.5 LBound = 1
    Arr = GenerateDoubleArray(3.5, 10, 0.5, 1)
    If Arr(LBound(Arr)) <> 3.5 Then
        TestGenerateDoubleArray = False
        Debug.Print "LBound Value (3.5, 10, 0.5, 1)"
    End If
    If Arr(UBound(Arr)) <> 8 Then
        TestGenerateDoubleArray = False
        Debug.Print "UBound Value (3.5, 10, 0.5, 1)"
    End If
    If LBound(Arr) <> 1 Then
        TestGenerateDoubleArray = False
        Debug.Print "LBound Index (3.5, 10, 0.5, 1)"
    End If
    If UBound(Arr) <> 10 Then
        TestGenerateDoubleArray = False
        Debug.Print "UBound Index (3.5, 10, 0.5, 1)"
    End If

    '3.5 To 8 Step 0.5 LBound = -10
    Arr = GenerateDoubleArray(3.5, 10, 0.5, -10)
    If Arr(LBound(Arr)) <> 3.5 Then
        TestGenerateDoubleArray = False
        Debug.Print "LBound Value (3.5, 10, 0.5, -10)"
    End If
    If Arr(UBound(Arr)) <> 8 Then
        TestGenerateDoubleArray = False
        Debug.Print "UBound Value (3.5, 10, 0.5, -10)"
    End If
    If LBound(Arr) <> -10 Then
        TestGenerateDoubleArray = False
        Debug.Print "LBound Index (3.5, 10, 0.5, -10)"
    End If
    If UBound(Arr) <> -1 Then
        TestGenerateDoubleArray = False
        Debug.Print "UBound Index (3.5, 10, 0.5, -10)"
    End If

    Debug.Print "TestGenerateDoubleArray: " & TestGenerateDoubleArray
    
End Function

Private Function TestGeneratePrimesArray() As Boolean
    
    TestGeneratePrimesArray = True
    
    Dim Arr() As Long
    
    '-1
    On Error Resume Next
    Arr = GeneratePrimesArray(1, -1)
    If Err.Number <> 5 Then
        TestGeneratePrimesArray = False
        Debug.Print "(1, -1)"
    End If
    On Error GoTo 0

    '0
    On Error Resume Next
    Arr = GeneratePrimesArray(1, 0)
    If Err.Number <> 5 Then
        TestGeneratePrimesArray = False
        Debug.Print "(1, 0)"
    End If
    On Error GoTo 0
    
    '-1
    On Error Resume Next
    Arr = GeneratePrimesArray(-1, 10)
    If Err.Number <> 5 Then
        TestGeneratePrimesArray = False
        Debug.Print "(-1, 10)"
    End If
    On Error GoTo 0
    
    '0
    On Error Resume Next
    Arr = GeneratePrimesArray(0, 10)
    If Err.Number <> 5 Then
        TestGeneratePrimesArray = False
        Debug.Print "(0, 10)"
    End If
    On Error GoTo 0
    
    '1
    Arr = GeneratePrimesArray(1, 10)
    If Arr(LBound(Arr)) <> 2 Then
        TestGeneratePrimesArray = False
        Debug.Print "LBound Value (1, 10)"
    End If
    If Arr(UBound(Arr)) <> 29 Then
        TestGeneratePrimesArray = False
        Debug.Print "UBound Value (1, 10)"
    End If
    
    '2
    Arr = GeneratePrimesArray(2, 10)
    If Arr(LBound(Arr)) <> 2 Then
        TestGeneratePrimesArray = False
        Debug.Print "LBound Value (2, 10)"
    End If
    If Arr(UBound(Arr)) <> 29 Then
        TestGeneratePrimesArray = False
        Debug.Print "UBound Value (2, 10)"
    End If
    
    '4
    Arr = GeneratePrimesArray(4, 10)
    If Arr(LBound(Arr)) <> 5 Then
        TestGeneratePrimesArray = False
        Debug.Print "LBound Value (4, 10)"
    End If
    If Arr(UBound(Arr)) <> 37 Then
        TestGeneratePrimesArray = False
        Debug.Print "UBound Value (4, 10)"
    End If
    
    '7
    Arr = GeneratePrimesArray(7, 10)
    If Arr(LBound(Arr)) <> 7 Then
        TestGeneratePrimesArray = False
        Debug.Print "LBound Value (7, 10)"
    End If
    If Arr(UBound(Arr)) <> 41 Then
        TestGeneratePrimesArray = False
        Debug.Print "UBound Value (7, 10)"
    End If
    
    'Bounds
    Arr = GeneratePrimesArray(1, 10, 1)
    If Arr(LBound(Arr)) <> 2 Then
        TestGeneratePrimesArray = False
        Debug.Print "LBound (1, 10, 1)"
    End If
    If Arr(UBound(Arr)) <> 29 Then
        TestGeneratePrimesArray = False
        Debug.Print "UBound (1, 10, 1)"
    End If
    If LBound(Arr) <> 1 Then
        TestGeneratePrimesArray = False
        Debug.Print "LBound Index (1, 10, 1)"
    End If
    If UBound(Arr) <> 10 Then
        TestGeneratePrimesArray = False
        Debug.Print "UBound Index (1, 10, 1)"
    End If
    
    Debug.Print "TestGeneratePrimesArray: " & TestGeneratePrimesArray
    
End Function

Private Function TestGeneratePatternArray() As Boolean
    
    TestGeneratePatternArray = True
    
    Dim Arr()
    Dim PatternArr()
    
    'Empty
    On Error Resume Next
    Arr = GeneratePatternArray(PatternArr, 0)
    If Err.Number <> 5 Then
        TestGeneratePatternArray = False
        Debug.Print "0"
    End If
    On Error GoTo 0
    
    ReDim PatternArr(0 To 2)
    PatternArr(0) = 2
    PatternArr(1) = 3
    PatternArr(2) = 5
    
    '-1
    On Error Resume Next
    Arr = GeneratePatternArray(PatternArr, -1)
    If Err.Number <> 5 Then
        TestGeneratePatternArray = False
        Debug.Print "-1"
    End If
    On Error GoTo 0
    
    '0
    On Error Resume Next
    Arr = GeneratePatternArray(PatternArr, 0)
    If Err.Number <> 5 Then
        TestGeneratePatternArray = False
        Debug.Print "0"
    End If
    On Error GoTo 0
    
    '1
    Arr = GeneratePatternArray(PatternArr, 1)
    If Arr(LBound(Arr)) <> 2 Or Arr(UBound(Arr)) <> 5 Then
        TestGeneratePatternArray = False
        Debug.Print "1 first and last"
    End If
    If SizeOfArray(Arr) <> 3 Then
        TestGeneratePatternArray = False
        Debug.Print "1 size"
    End If
    If LBound(Arr) <> 0 Then
        TestGeneratePatternArray = False
        Debug.Print "1 LBound"
    End If
    
    '2
    Arr = GeneratePatternArray(PatternArr, 2)
    If Arr(LBound(Arr)) <> 2 Or Arr(UBound(Arr)) <> 5 Then
        TestGeneratePatternArray = False
        Debug.Print "2 first and last"
    End If
    If SizeOfArray(Arr) <> 6 Then
        TestGeneratePatternArray = False
        Debug.Print "2 size"
    End If
    
    'LBound
    Arr = GeneratePatternArray(PatternArr, 1, 1)
    If Arr(LBound(Arr)) <> 2 Or Arr(UBound(Arr)) <> 5 Then
        TestGeneratePatternArray = False
        Debug.Print "1 first and last LBound = 1"
    End If
    If SizeOfArray(Arr) <> 3 Then
        TestGeneratePatternArray = False
        Debug.Print "1 size LBound = 1"
    End If
    If LBound(Arr) <> 1 Then
        TestGeneratePatternArray = False
        Debug.Print "1 LBound LBound = 1"
    End If
    
    Debug.Print "TestGeneratePatternArray: " & TestGeneratePatternArray
    
End Function

Private Function TestGenerateFibonacciArray() As Boolean

    TestGenerateFibonacciArray = True

    Dim Arr() As Long

    '-1
    On Error Resume Next
    Arr = GenerateFibonacciArray(1, 1, -1)
    If Err.Number <> 5 Then
        TestGenerateFibonacciArray = False
        Debug.Print "-1"
    End If
    On Error GoTo 0

    '0
    On Error Resume Next
    Arr = GenerateFibonacciArray(1, 1, 0)
    If Err.Number <> 5 Then
        TestGenerateFibonacciArray = False
        Debug.Print "0"
    End If
    On Error GoTo 0
    
    '(1, 1, 10)
    Arr = GenerateFibonacciArray(1, 1, 10)
    If Arr(LBound(Arr)) <> 1 Then
        TestGenerateFibonacciArray = False
        Debug.Print "LBound Value (1, 1, 10)"
    End If
    If Arr(UBound(Arr)) <> 55 Then
        TestGenerateFibonacciArray = False
        Debug.Print "UBound Value (1, 1, 10)"
    End If
    If LBound(Arr) <> 0 Then
        TestGenerateFibonacciArray = False
        Debug.Print "LBound Index (1, 1, 10)"
    End If

    '(5, 10, 10)
    Arr = GenerateFibonacciArray(5, 10, 10)
    If Arr(LBound(Arr)) <> 5 Then
        TestGenerateFibonacciArray = False
        Debug.Print "LBound Value (5, 10, 10)"
    End If
    If Arr(UBound(Arr)) <> 445 Then
        TestGenerateFibonacciArray = False
        Debug.Print "UBound Value (5, 10, 10)"
    End If

    '(1, 1, 10, 1)
    Arr = GenerateFibonacciArray(1, 1, 10, 1)
    If Arr(LBound(Arr)) <> 1 Then
        TestGenerateFibonacciArray = False
        Debug.Print "LBound Value (1, 1, 10, 1) LBound = 1"
    End If
    If Arr(UBound(Arr)) <> 55 Then
        TestGenerateFibonacciArray = False
        Debug.Print "UBound Value (1, 1, 10, 1) LBound = 1"
    End If
    If LBound(Arr) <> 1 Then
        TestGenerateFibonacciArray = False
        Debug.Print "LBound Index (1, 1, 10, 1) LBound = 1"
    End If

    Debug.Print "TestGenerateFibonacciArray: " & TestGenerateFibonacciArray

End Function

Private Function TestGenerateRndArray() As Boolean

    TestGenerateRndArray = True

    Dim Arr() As Double

    '-1
    On Error Resume Next
    Arr = GenerateRndArray(-1)
    If Err.Number <> 5 Then
        TestGenerateRndArray = False
        Debug.Print "-1"
    End If
    On Error GoTo 0

    '0
    On Error Resume Next
    Arr = GenerateRndArray(0)
    If Err.Number <> 5 Then
        TestGenerateRndArray = False
        Debug.Print "0"
    End If
    On Error GoTo 0

    '1
    Arr = GenerateRndArray(1)
    If LBound(Arr) <> 0 Or UBound(Arr) <> 0 Then
        TestGenerateRndArray = False
        Debug.Print "1"
    End If

    '2
    Arr = GenerateRndArray(2)
    If LBound(Arr) <> 0 Or UBound(Arr) <> 1 Then
        TestGenerateRndArray = False
        Debug.Print "2"
    End If

    'Bounds
    Arr = GenerateRndArray(10000)
    Dim i As Long
    For i = LBound(Arr) To UBound(Arr)
        If Arr(i) >= 1 Or Arr(i) < 0 Then
            TestGenerateRndArray = False
            Debug.Print "Out of bounds"
        End If
    Next i

    Debug.Print "TestGenerateRndArray: " & TestGenerateRndArray

End Function

Private Function TestGenerateRandomLongArray() As Boolean

    TestGenerateRandomLongArray = True
    
    Dim Arr() As Long
    
    '-1
    On Error Resume Next
    Arr = GenerateRandomLongArray(1, 10, -1)
    If Err.Number <> 5 Then
        TestGenerateRandomLongArray = False
        Debug.Print "-1"
    End If
    On Error GoTo 0
    
    '0
    On Error Resume Next
    Arr = GenerateRandomLongArray(1, 10, 0)
    If Err.Number <> 5 Then
        TestGenerateRandomLongArray = False
        Debug.Print "0"
    End If
    On Error GoTo 0
    
    '1
    Arr = GenerateRandomLongArray(1, 10, 1)
    If LBound(Arr) <> 0 Or UBound(Arr) <> 0 Then
        TestGenerateRandomLongArray = False
        Debug.Print "1"
    End If
    
    '10
    Arr = GenerateRandomLongArray(1, 10, 10)
    If LBound(Arr) <> 0 Or UBound(Arr) <> 9 Then
        TestGenerateRandomLongArray = False
        Debug.Print "10"
    End If
    
    '10 LBound = 0
    Arr = GenerateRandomLongArray(1, 10, 10, 1)
    If LBound(Arr) <> 1 Or UBound(Arr) <> 10 Then
        TestGenerateRandomLongArray = False
        Debug.Print "10 LBound = 1"
    End If
    
    Arr = GenerateRandomLongArray(1, 5, 100)
    Dim F1 As Boolean
    Dim F2 As Boolean
    Dim F3 As Boolean
    Dim F4 As Boolean
    Dim F5 As Boolean
    Dim i As Long
    For i = LBound(Arr) To UBound(Arr)
        Select Case Arr(i)
            Case 1
                F1 = True
            Case 2
                F2 = True
            Case 3
                F3 = True
            Case 4
                F4 = True
            Case 5
                F5 = True
            Case Else
                TestGenerateRandomLongArray = False
                Debug.Print "Out of Bounds"
        End Select
    Next i
    If Not (F1 And F2 And F3 And F4 And F5) Then
        TestGenerateRandomLongArray = False
        Debug.Print "Not all there"
    End If
    
    Debug.Print "TestGenerateRandomLongArray: " & TestGenerateRandomLongArray
    
End Function

Private Function TestGenerateRandomDoubleArray() As Boolean

    TestGenerateRandomDoubleArray = True

    Dim Arr() As Double

    '-1
    On Error Resume Next
    Arr = GenerateRandomDoubleArray(1, 10, -1)
    If Err.Number <> 5 Then
        TestGenerateRandomDoubleArray = False
        Debug.Print "-1"
    End If
    On Error GoTo 0

    '0
    On Error Resume Next
    Arr = GenerateRandomDoubleArray(1, 10, 0)
    If Err.Number <> 5 Then
        TestGenerateRandomDoubleArray = False
        Debug.Print "0"
    End If
    On Error GoTo 0

    '1
    Arr = GenerateRandomDoubleArray(1, 10, 1)
    If LBound(Arr) <> 0 Or UBound(Arr) <> 0 Then
        TestGenerateRandomDoubleArray = False
        Debug.Print "1"
    End If

    '10
    Arr = GenerateRandomDoubleArray(1, 10, 10)
    If LBound(Arr) <> 0 Or UBound(Arr) <> 9 Then
        TestGenerateRandomDoubleArray = False
        Debug.Print "10"
    End If

    '10 LBound = 0
    Arr = GenerateRandomDoubleArray(1, 10, 10, 1)
    If LBound(Arr) <> 1 Or UBound(Arr) <> 10 Then
        TestGenerateRandomDoubleArray = False
        Debug.Print "10 LBound = 1"
    End If

    Arr = GenerateRandomDoubleArray(1, 5, 100)
    Dim F1 As Boolean
    Dim F2 As Boolean
    Dim F3 As Boolean
    Dim F4 As Boolean
    Dim F5 As Boolean
    Dim i As Long
    For i = LBound(Arr) To UBound(Arr)
        Select Case CLng(Arr(i))
            Case 1
                F1 = True
            Case 2
                F2 = True
            Case 3
                F3 = True
            Case 4
                F4 = True
            Case 5
                F5 = True
            Case Else
                TestGenerateRandomDoubleArray = False
                Debug.Print "Out of Bounds"
        End Select
    Next i
    If Not (F1 And F2 And F3 And F4 And F5) Then
        TestGenerateRandomDoubleArray = False
        Debug.Print "Not all there"
    End If

    Debug.Print "TestGenerateRandomDoubleArray: " & TestGenerateRandomDoubleArray

End Function

Private Function TestReverseArray() As Boolean
    
    TestReverseArray = True
    
    Dim Arr() As Long
    
    'Empty
    On Error Resume Next
    ReverseArray Arr
    If Err.Number <> 0 Then
        TestReverseArray = False
        Debug.Print "Empty"
    End If
    On Error GoTo 0
    
    'One
    ReDim Arr(0 To 0) As Long
    Arr(0) = 1
    ReverseArray Arr
    If Arr(0) <> 1 Then
        TestReverseArray = False
        Debug.Print "One"
    End If
    
    'Multiple
    ReDim Arr(0 To 4) As Long
    Arr(0) = 1
    Arr(1) = 2
    Arr(2) = 3
    Arr(3) = 4
    Arr(4) = 5
    ReverseArray Arr
    If Arr(0) <> 5 Or _
    Arr(1) <> 4 Or _
    Arr(2) <> 3 Or _
    Arr(3) <> 2 Or _
    Arr(4) <> 1 Then
        TestReverseArray = False
        Debug.Print "Multiple"
    End If
    
    Debug.Print "TestReverseArray: " & TestReverseArray
    
End Function

Private Function TestReverseObjectArray() As Boolean

    TestReverseObjectArray = True
    
    Dim Arr() As Object
    
    'Not Object
    Dim NOArr(0 To 0)
    NOArr(0) = 1
    On Error Resume Next
    ReverseObjectArray NOArr
    If Err.Number <> 5 Then
        TestReverseObjectArray = False
        Debug.Print "Not Object"
    End If
    On Error GoTo 0
    
    'Empty
    On Error Resume Next
    ReverseObjectArray Arr
    If Err.Number <> 0 Then
        TestReverseObjectArray = False
        Debug.Print "Empty"
    End If
    On Error GoTo 0
    
    'One
    ReDim Arr(0 To 0) As Object
    Set Arr(0) = New Collection
    Arr(0).Add 1
    ReverseObjectArray Arr
    If Arr(0).Item(1) <> 1 Then
        TestReverseObjectArray = False
        Debug.Print "One"
    End If
    
    'Multiple
    ReDim Arr(0 To 4) As Object
    Set Arr(0) = New Collection
    Arr(0).Add 1
    Set Arr(1) = New Collection
    Arr(1).Add 2
    Set Arr(2) = New Collection
    Arr(2).Add 3
    Set Arr(3) = New Collection
    Arr(3).Add 4
    Set Arr(4) = New Collection
    Arr(4).Add 5
    ReverseObjectArray Arr
    If Arr(0).Item(1) <> 5 Or _
    Arr(1).Item(1) <> 4 Or _
    Arr(2).Item(1) <> 3 Or _
    Arr(3).Item(1) <> 2 Or _
    Arr(4).Item(1) <> 1 Then
        TestReverseObjectArray = False
        Debug.Print "Multiple"
    End If
    
    Debug.Print "TestReverseObjectArray: " & TestReverseObjectArray
    
End Function

Private Function TestSearchArray() As Boolean

    TestSearchArray = True
    
    Dim Arr() As Long
    
    Dim i&
    
    'Empty
    On Error Resume Next
    i = SearchArray(Arr, 1)
    If Err.Number <> 0 Or i <> -1 Then
        TestSearchArray = False
        Debug.Print "Empty"
    End If
    On Error GoTo 0
    
    'One
    ReDim Arr(0 To 0)
        'Not there
        If SearchArray(Arr, 1) <> -1 Then
            TestSearchArray = False
            Debug.Print "One Not there"
        End If
        'There
        If SearchArray(Arr, 0) <> 0 Then
            TestSearchArray = False
            Debug.Print "One There"
        End If
        
    'Multiple
    ReDim Arr(0 To 2)
        'Not there
        If SearchArray(Arr, 1) <> -1 Then
            TestSearchArray = False
            Debug.Print "Multiple Not there"
        End If
        'There
        Arr(1) = 1
        Arr(2) = 1
        If SearchArray(Arr, 1) <> 1 Then
            TestSearchArray = False
            Debug.Print "Multiple There"
        End If
        
    'Direction
    ReDim Arr(0 To 2)
        'Not there
        If SearchArray(Arr, 1, False) <> -1 Then
            TestSearchArray = False
            Debug.Print "Multiple Not there Direction"
        End If
        'There
        Arr(1) = 1
        Arr(2) = 1
        If SearchArray(Arr, 1, False) <> 2 Then
            TestSearchArray = False
            Debug.Print "Multiple There Direction"
        End If
        
    Debug.Print "TestSearchArray: " & TestSearchArray
    
End Function

Private Function TestSearchStringArray() As Boolean

    TestSearchStringArray = True
    
    Dim Arr() As String
    
    Dim i&
    
    'Empty
    On Error Resume Next
    i = SearchStringArray(Arr, "A", vbBinaryCompare)
    If Err.Number <> 0 Or i <> -1 Then
        TestSearchStringArray = False
        Debug.Print "Empty"
    End If
    On Error GoTo 0
    
    'One
    ReDim Arr(0 To 0)
        'Not there
        If SearchStringArray(Arr, "A", vbBinaryCompare) <> -1 Then
            TestSearchStringArray = False
            Debug.Print "One Not there"
        End If
        'There
        Arr(0) = "A"
        If SearchStringArray(Arr, "A", vbBinaryCompare) <> 0 Then
            TestSearchStringArray = False
            Debug.Print "One There"
        End If
        
    'Multiple
    ReDim Arr(0 To 2)
        'Not there
        If SearchStringArray(Arr, "A", vbBinaryCompare) <> -1 Then
            TestSearchStringArray = False
            Debug.Print "Multiple Not there"
        End If
        'There
        Arr(1) = "A"
        Arr(2) = "A"
        If SearchStringArray(Arr, "A", vbBinaryCompare) <> 1 Then
            TestSearchStringArray = False
            Debug.Print "Multiple There"
        End If
     
    'Compare Text
    ReDim Arr(0 To 2)
        'Not there
        If SearchStringArray(Arr, "A", vbTextCompare) <> -1 Then
            TestSearchStringArray = False
            Debug.Print "Multiple Not there Compare Text"
        End If
        'There
        Arr(1) = "A"
        Arr(2) = "a"
        If SearchStringArray(Arr, "A", vbTextCompare) <> 1 Then
            TestSearchStringArray = False
            Debug.Print "Multiple There Compare Text"
        End If
        
    'Direction
    ReDim Arr(0 To 2)
        'Not there
        If SearchStringArray(Arr, "A", vbBinaryCompare, False) <> -1 Then
            TestSearchStringArray = False
            Debug.Print "Multiple Not there Direction"
        End If
        'There
        Arr(1) = "A"
        Arr(2) = "A"
        If SearchStringArray(Arr, "A", vbBinaryCompare, False) <> 2 Then
            TestSearchStringArray = False
            Debug.Print "Multiple There Direction"
        End If
        
    Debug.Print "TestSearchStringArray: " & TestSearchStringArray
        
End Function

Private Function TestSearchObjectArray() As Boolean

    TestSearchObjectArray = True
    
    Dim Arr() As Object
    
    Dim i&
    
    'Not Object
    Dim NOArr(0 To 0)
    NOArr(0) = 1
    On Error Resume Next
    i = SearchObjectArray(NOArr, 1, "Count", True)
    If Err.Number <> 5 Then
        TestSearchObjectArray = False
        Debug.Print "Not Object"
    End If
    On Error GoTo 0
    
    'Empty
    On Error Resume Next
    i = SearchObjectArray(Arr, 1, "Count", True)
    If Err.Number <> 0 Or i <> -1 Then
        TestSearchObjectArray = False
        Debug.Print "Empty"
    End If
    On Error GoTo 0
    
    'One
    ReDim Arr(0 To 0)
    Set Arr(0) = New Collection
        'Not there
        If SearchObjectArray(Arr, 1, "Count", True) <> -1 Then
            TestSearchObjectArray = False
            Debug.Print "One Not there"
        End If
        'There
        Arr(0).Add 1
        If SearchObjectArray(Arr, 1, "Count", True) <> 0 Then
            TestSearchObjectArray = False
            Debug.Print "One There"
        End If
        
    'Multiple
    ReDim Arr(0 To 2)
    Set Arr(0) = New Collection
    Set Arr(1) = New Collection
    Set Arr(2) = New Collection
        'Not there
        If SearchObjectArray(Arr, 1, "Count", True) <> -1 Then
            TestSearchObjectArray = False
            Debug.Print "Multiple Not there"
        End If
        'There
        Arr(1).Add 1
        Arr(2).Add 1
        If SearchObjectArray(Arr, 1, "Count", True) <> 1 Then
            TestSearchObjectArray = False
            Debug.Print "Multiple There"
        End If
        
    'Direction
    ReDim Arr(0 To 2)
    Set Arr(0) = New Collection
    Set Arr(1) = New Collection
    Set Arr(2) = New Collection
        'Not there
        If SearchObjectArray(Arr, 1, "Count", True, False) <> -1 Then
            TestSearchObjectArray = False
            Debug.Print "Multiple Not there Direction"
        End If
        'There
        Arr(1).Add 1
        Arr(2).Add 1
        If SearchObjectArray(Arr, 1, "Count", True, False) <> 2 Then
            TestSearchObjectArray = False
            Debug.Print "Multiple There Direction"
        End If
        
    Debug.Print "TestSearchObjectArray: " & TestSearchObjectArray
        
End Function

Private Function TestShuffleArray() As Boolean
    
    TestShuffleArray = True
    
    'Empty
    Dim Arr() As Long
    On Error Resume Next
    ShuffleArray Arr
    If Err.Number <> 0 Then
        TestShuffleArray = False
        Debug.Print "Empty"
    End If
    On Error GoTo 0
    
    'One
    ReDim Arr(0 To 0)
    On Error Resume Next
    ShuffleArray Arr
    If Err.Number <> 0 Or Arr(0) <> 0 Then
        TestShuffleArray = False
        Debug.Print "One"
    End If
    On Error GoTo 0
    
    'Multiple
    ReDim Arr(0 To 4)
    Arr(0) = 1
    Arr(1) = 2
    Arr(2) = 3
    Arr(3) = 4
    Arr(4) = 5
    ShuffleArray Arr
    Dim F1 As Boolean
    Dim F2 As Boolean
    Dim F3 As Boolean
    Dim F4 As Boolean
    Dim F5 As Boolean
    Dim i As Long
    For i = LBound(Arr) To UBound(Arr)
        Select Case Arr(i)
            Case 1
                F1 = True
            Case 2
                F2 = True
            Case 3
                F3 = True
            Case 4
                F4 = True
            Case 5
                F5 = True
            Case Else
                TestShuffleArray = False
                Debug.Print "Shouldn't be there"
        End Select
    Next i
    If Not (F1 And F2 And F3 And F4 And F5) Then
        TestShuffleArray = False
        Debug.Print "Not all there"
    End If
    If Arr(0) = 1 And Arr(1) = 2 And Arr(2) = 3 And Arr(3) = 4 And Arr(4) = 5 Then
        ShuffleArray Arr
        If Arr(0) = 1 And Arr(1) = 2 And Arr(2) = 3 And Arr(3) = 4 And Arr(4) = 5 Then
            ShuffleArray Arr
            If Arr(0) = 1 And Arr(1) = 2 And Arr(2) = 3 And Arr(3) = 4 And Arr(4) = 5 Then
                TestShuffleArray = False
                Debug.Print "Did not shuffle"
            End If
        End If
    End If
    
    Debug.Print "TestShuffleArray: " & TestShuffleArray

End Function

Private Function TestShuffleObjectArray() As Boolean

    TestShuffleObjectArray = True
    
    'Not Object
    Dim NOArr(0 To 0)
    NOArr(0) = 1
    On Error Resume Next
    ShuffleObjectArray NOArr
    If Err.Number <> 5 Then
        TestShuffleObjectArray = False
        Debug.Print "Not Object"
    End If
    On Error GoTo 0
    
    'Empty
    Dim Arr() As Object
    On Error Resume Next
    ShuffleObjectArray Arr
    If Err.Number <> 0 Then
        TestShuffleObjectArray = False
        Debug.Print "Empty"
    End If
    On Error GoTo 0
    
    'One
    ReDim Arr(0 To 0)
    Set Arr(0) = New Collection
    On Error Resume Next
    ShuffleObjectArray Arr
    If Err.Number <> 0 Or Arr(0).Count <> 0 Then
        TestShuffleObjectArray = False
        Debug.Print "One"
    End If
    On Error GoTo 0
    
    'Multiple
    ReDim Arr(0 To 4)
    Set Arr(0) = New Collection
    Arr(0).Add 1
    Set Arr(1) = New Collection
    Arr(1).Add 2
    Set Arr(2) = New Collection
    Arr(2).Add 3
    Set Arr(3) = New Collection
    Arr(3).Add 4
    Set Arr(4) = New Collection
    Arr(4).Add 5
    ShuffleObjectArray Arr
    Dim F1 As Boolean
    Dim F2 As Boolean
    Dim F3 As Boolean
    Dim F4 As Boolean
    Dim F5 As Boolean
    Dim i As Long
    For i = LBound(Arr) To UBound(Arr)
        Select Case Arr(i).Item(1)
            Case 1
                F1 = True
            Case 2
                F2 = True
            Case 3
                F3 = True
            Case 4
                F4 = True
            Case 5
                F5 = True
            Case Else
                TestShuffleObjectArray = False
                Debug.Print "Shouldn't be there"
        End Select
    Next i
    If Not (F1 And F2 And F3 And F4 And F5) Then
        TestShuffleObjectArray = False
        Debug.Print "Not all there"
    End If
    If Arr(0).Item(1) = 1 And Arr(1).Item(1) = 2 And Arr(2).Item(1) = 3 And _
    Arr(3).Item(1) = 4 And Arr(4).Item(1) = 5 Then
        ShuffleObjectArray Arr
        If Arr(0).Item(1) = 1 And Arr(1).Item(1) = 2 And Arr(2).Item(1) = 3 And _
        Arr(3).Item(1) = 4 And Arr(4).Item(1) = 5 Then
            ShuffleObjectArray Arr
            If Arr(0).Item(1) = 1 And Arr(1).Item(1) = 2 And Arr(2).Item(1) = 3 And _
            Arr(3).Item(1) = 4 And Arr(4).Item(1) = 5 Then
                TestShuffleObjectArray = False
                Debug.Print "Did not shuffle"
            End If
        End If
    End If
    
    Debug.Print "TestShuffleObjectArray: " & TestShuffleObjectArray
    
End Function

Private Function TestSizeOfArray() As Boolean
    
    TestSizeOfArray = True
    
    Dim LArr() As Long
    Dim SArr() As String
    Dim OArr() As Object
    
    Dim i&
    
    'Empty
    On Error Resume Next
    i = SizeOfArray(LArr)
    If Err.Number <> 0 Or i <> 0 Then
        TestSizeOfArray = False
        Debug.Print "Empty Long"
    End If
    On Error GoTo 0
    On Error Resume Next
    i = SizeOfArray(SArr)
    If Err.Number <> 0 Or i <> 0 Then
        TestSizeOfArray = False
        Debug.Print "Empty String"
    End If
    On Error GoTo 0
        On Error Resume Next
    i = SizeOfArray(OArr)
    If Err.Number <> 0 Or i <> 0 Then
        TestSizeOfArray = False
        Debug.Print "Empty Object"
    End If
    On Error GoTo 0
    
    'One
    ReDim LArr(0 To 0)
    ReDim SArr(0 To 0)
    ReDim OArr(0 To 0)
    If SizeOfArray(LArr) <> 1 Then
        TestSizeOfArray = False
        Debug.Print "One Long"
    End If
    If SizeOfArray(SArr) <> 1 Then
        TestSizeOfArray = False
        Debug.Print "One String"
    End If
    If SizeOfArray(OArr) <> 1 Then
        TestSizeOfArray = False
        Debug.Print "One Object"
    End If
    
    'Multiple
    ReDim LArr(0 To 2)
    ReDim SArr(0 To 2)
    ReDim OArr(0 To 2)
    If SizeOfArray(LArr) <> 3 Then
        TestSizeOfArray = False
        Debug.Print "Multiple Long"
    End If
    If SizeOfArray(SArr) <> 3 Then
        TestSizeOfArray = False
        Debug.Print "Multiple String"
    End If
    If SizeOfArray(OArr) <> 3 Then
        TestSizeOfArray = False
        Debug.Print "Multiple Object"
    End If
    
    Debug.Print "TestSizeOfArray: " & TestSizeOfArray
    
End Function

Private Function TestSortArray() As Boolean

    TestSortArray = True
    
    Dim Arr() As Long

    'Unitialized
    On Error Resume Next
    SortArray Arr
    If Err.Number <> 0 Then
        TestSortArray = False
        Debug.Print "Uninitialized"
        Debug.Print "Error " & Err.Number & ": " & Err.Description
    End If
    On Error GoTo 0

    'One
    ReDim Arr(0 To 0) As Long
    Arr(0) = 1
    SortArray Arr
    If Arr(0) <> 1 Then
        TestSortArray = False
        Debug.Print "One"
    End If
    
    ReDim Arr(0 To 4) As Long

    'Shuffled
    Arr(0) = 3
    Arr(1) = 1
    Arr(2) = 4
    Arr(3) = 5
    Arr(4) = 2
    SortArray Arr
    If Arr(0) <> 1 Or _
    Arr(1) <> 2 Or _
    Arr(2) <> 3 Or _
    Arr(3) <> 4 Or _
    Arr(4) <> 5 Then
        TestSortArray = False
        Debug.Print "Shuffled"
    End If

    'Sorted
    Arr(0) = 1
    Arr(1) = 2
    Arr(2) = 3
    Arr(3) = 4
    Arr(4) = 5
    SortArray Arr
    If Arr(0) <> 1 Or _
    Arr(1) <> 2 Or _
    Arr(2) <> 3 Or _
    Arr(3) <> 4 Or _
    Arr(4) <> 5 Then
        TestSortArray = False
        Debug.Print "Sorted"
    End If

    'Reverse Sorted
    Arr(0) = 5
    Arr(1) = 4
    Arr(2) = 3
    Arr(3) = 2
    Arr(4) = 1
    SortArray Arr
    If Arr(0) <> 1 Or _
    Arr(1) <> 2 Or _
    Arr(2) <> 3 Or _
    Arr(3) <> 4 Or _
    Arr(4) <> 5 Then
        TestSortArray = False
        Debug.Print "Reverse sorted"
    End If

    'Same
    Arr(0) = 1
    Arr(1) = 1
    Arr(2) = 1
    Arr(3) = 1
    Arr(4) = 1
    SortArray Arr
    If Arr(0) <> 1 Or _
    Arr(1) <> 1 Or _
    Arr(2) <> 1 Or _
    Arr(3) <> 1 Or _
    Arr(4) <> 1 Then
        TestSortArray = False
        Debug.Print "Same"
    End If

    Debug.Print "TestSortArray: " & TestSortArray

End Function

Private Function TestInsertionSortV() As Boolean

    TestInsertionSortV = True

    Dim Arr() As Long

    'Unitialized
    On Error Resume Next
    pInsertionSortV Arr, 0
    If Err.Number <> 0 Then
        TestInsertionSortV = False
        Debug.Print "Uninitialized"
        Debug.Print "Error " & Err.Number & ": " & Err.Description
    End If
    On Error GoTo 0

    'One
    ReDim Arr(0 To 0) As Long
    Arr(0) = 1
    pInsertionSortV Arr, 1
    If Arr(0) <> 1 Then
        TestInsertionSortV = False
        Debug.Print "One"
    End If
    
    ReDim Arr(0 To 4) As Long

    'Shuffled
    Arr(0) = 3
    Arr(1) = 1
    Arr(2) = 4
    Arr(3) = 5
    Arr(4) = 2
    pInsertionSortV Arr, 5
    If Arr(0) <> 1 Or _
    Arr(1) <> 2 Or _
    Arr(2) <> 3 Or _
    Arr(3) <> 4 Or _
    Arr(4) <> 5 Then
        TestInsertionSortV = False
        Debug.Print "Shuffled"
    End If

    'Sorted
    Arr(0) = 1
    Arr(1) = 2
    Arr(2) = 3
    Arr(3) = 4
    Arr(4) = 5
    pInsertionSortV Arr, 5
    If Arr(0) <> 1 Or _
    Arr(1) <> 2 Or _
    Arr(2) <> 3 Or _
    Arr(3) <> 4 Or _
    Arr(4) <> 5 Then
        TestInsertionSortV = False
        Debug.Print "Sorted"
    End If

    'Reverse Sorted
    Arr(0) = 5
    Arr(1) = 4
    Arr(2) = 3
    Arr(3) = 2
    Arr(4) = 1
    pInsertionSortV Arr, 5
    If Arr(0) <> 1 Or _
    Arr(1) <> 2 Or _
    Arr(2) <> 3 Or _
    Arr(3) <> 4 Or _
    Arr(4) <> 5 Then
        TestInsertionSortV = False
        Debug.Print "Reverse sorted"
    End If

    'Same
    Arr(0) = 1
    Arr(1) = 1
    Arr(2) = 1
    Arr(3) = 1
    Arr(4) = 1
    pInsertionSortV Arr, 5
    If Arr(0) <> 1 Or _
    Arr(1) <> 1 Or _
    Arr(2) <> 1 Or _
    Arr(3) <> 1 Or _
    Arr(4) <> 1 Then
        TestInsertionSortV = False
        Debug.Print "Same"
    End If

    Debug.Print "TestInsertionSortV: " & TestInsertionSortV

End Function

Private Function TestMergeSortV() As Boolean

    TestMergeSortV = True

    Dim Arr() As Long

    'Unitialized
    On Error Resume Next
    pMergeSortV Arr, 0
    If Err.Number <> 0 Then
        TestMergeSortV = False
        Debug.Print "Uninitialized"
        Debug.Print "Error " & Err.Number & ": " & Err.Description
    End If
    On Error GoTo 0

    'One
    ReDim Arr(0 To 0) As Long
    Arr(0) = 1
    pMergeSortV Arr, 1
    If Arr(0) <> 1 Then
        TestMergeSortV = False
        Debug.Print "One"
    End If
    
    ReDim Arr(0 To 4) As Long

    'Shuffled
    Arr(0) = 3
    Arr(1) = 1
    Arr(2) = 4
    Arr(3) = 5
    Arr(4) = 2
    pMergeSortV Arr, 5
    If Arr(0) <> 1 Or _
    Arr(1) <> 2 Or _
    Arr(2) <> 3 Or _
    Arr(3) <> 4 Or _
    Arr(4) <> 5 Then
        TestMergeSortV = False
        Debug.Print "Shuffled"
    End If

    'Sorted
    Arr(0) = 1
    Arr(1) = 2
    Arr(2) = 3
    Arr(3) = 4
    Arr(4) = 5
    pMergeSortV Arr, 5
    If Arr(0) <> 1 Or _
    Arr(1) <> 2 Or _
    Arr(2) <> 3 Or _
    Arr(3) <> 4 Or _
    Arr(4) <> 5 Then
        TestMergeSortV = False
        Debug.Print "Sorted"
    End If

    'Reverse Sorted
    Arr(0) = 5
    Arr(1) = 4
    Arr(2) = 3
    Arr(3) = 2
    Arr(4) = 1
    pMergeSortV Arr, 5
    If Arr(0) <> 1 Or _
    Arr(1) <> 2 Or _
    Arr(2) <> 3 Or _
    Arr(3) <> 4 Or _
    Arr(4) <> 5 Then
        TestMergeSortV = False
        Debug.Print "Reverse sorted"
    End If

    'Same
    Arr(0) = 1
    Arr(1) = 1
    Arr(2) = 1
    Arr(3) = 1
    Arr(4) = 1
    pMergeSortV Arr, 5
    If Arr(0) <> 1 Or _
    Arr(1) <> 1 Or _
    Arr(2) <> 1 Or _
    Arr(3) <> 1 Or _
    Arr(4) <> 1 Then
        TestMergeSortV = False
        Debug.Print "Same"
    End If

    Debug.Print "TestMergeSortV: " & TestMergeSortV

End Function

Private Function TestSortStringArray() As Boolean

    TestSortStringArray = True

    Dim Arr() As String

    'Unitialized
    On Error Resume Next
    SortStringArray Arr
    If Err.Number <> 0 Then
        TestSortStringArray = False
        Debug.Print "Uninitialized"
        Debug.Print "Error " & Err.Number & ": " & Err.Description
    End If
    On Error GoTo 0

    'One
    ReDim Arr(0 To 0) As String
    Arr(0) = "A"
    SortStringArray Arr, vbBinaryCompare
    If Arr(0) <> "A" Then
        TestSortStringArray = False
        Debug.Print "One"
    End If
    
    ReDim Arr(0 To 4) As String

    'Shuffled
    Arr(0) = "C"
    Arr(1) = "A"
    Arr(2) = "D"
    Arr(3) = "E"
    Arr(4) = "B"
    SortStringArray Arr, vbBinaryCompare
    If Arr(0) <> "A" Or _
    Arr(1) <> "B" Or _
    Arr(2) <> "C" Or _
    Arr(3) <> "D" Or _
    Arr(4) <> "E" Then
        TestSortStringArray = False
        Debug.Print "Shuffled"
    End If

    'Sorted
    Arr(0) = "A"
    Arr(1) = "B"
    Arr(2) = "C"
    Arr(3) = "D"
    Arr(4) = "E"
    SortStringArray Arr, vbBinaryCompare
    If Arr(0) <> "A" Or _
    Arr(1) <> "B" Or _
    Arr(2) <> "C" Or _
    Arr(3) <> "D" Or _
    Arr(4) <> "E" Then
        TestSortStringArray = False
        Debug.Print "Sorted"
    End If

    'Reverse Sorted
    Arr(0) = "E"
    Arr(1) = "D"
    Arr(2) = "C"
    Arr(3) = "B"
    Arr(4) = "A"
    SortStringArray Arr, vbBinaryCompare
    If Arr(0) <> "A" Or _
    Arr(1) <> "B" Or _
    Arr(2) <> "C" Or _
    Arr(3) <> "D" Or _
    Arr(4) <> "E" Then
        TestSortStringArray = False
        Debug.Print "Reverse sorted"
    End If

    'Same
    Arr(0) = "A"
    Arr(1) = "A"
    Arr(2) = "A"
    Arr(3) = "A"
    Arr(4) = "A"
    SortStringArray Arr
    If Arr(0) <> "A" Or _
    Arr(1) <> "A" Or _
    Arr(2) <> "A" Or _
    Arr(3) <> "A" Or _
    Arr(4) <> "A" Then
        TestSortStringArray = False
        Debug.Print "Same"
    End If

    Debug.Print "TestSortStringArray: " & TestSortStringArray

End Function

Private Function TestInsertionSortS() As Boolean

    TestInsertionSortS = True

    Dim Arr() As String

    'Unitialized
    On Error Resume Next
    pInsertionSortS Arr, 0, vbBinaryCompare
    If Err.Number <> 0 Then
        TestInsertionSortS = False
        Debug.Print "Uninitialized"
        Debug.Print "Error " & Err.Number & ": " & Err.Description
    End If
    On Error GoTo 0

    'One
    ReDim Arr(0 To 0) As String
    Arr(0) = "A"
    pInsertionSortS Arr, 1, vbBinaryCompare
    If Arr(0) <> "A" Then
        TestInsertionSortS = False
        Debug.Print "One"
    End If
    
    ReDim Arr(0 To 4) As String

    'Shuffled
    Arr(0) = "C"
    Arr(1) = "A"
    Arr(2) = "D"
    Arr(3) = "E"
    Arr(4) = "B"
    pInsertionSortS Arr, 5, vbBinaryCompare
    If Arr(0) <> "A" Or _
    Arr(1) <> "B" Or _
    Arr(2) <> "C" Or _
    Arr(3) <> "D" Or _
    Arr(4) <> "E" Then
        TestInsertionSortS = False
        Debug.Print "Shuffled"
    End If

    'Sorted
    Arr(0) = "A"
    Arr(1) = "B"
    Arr(2) = "C"
    Arr(3) = "D"
    Arr(4) = "E"
    pInsertionSortS Arr, 5, vbBinaryCompare
    If Arr(0) <> "A" Or _
    Arr(1) <> "B" Or _
    Arr(2) <> "C" Or _
    Arr(3) <> "D" Or _
    Arr(4) <> "E" Then
        TestInsertionSortS = False
        Debug.Print "Sorted"
    End If

    'Reverse Sorted
    Arr(0) = "E"
    Arr(1) = "D"
    Arr(2) = "C"
    Arr(3) = "B"
    Arr(4) = "A"
    pInsertionSortS Arr, 5, vbBinaryCompare
    If Arr(0) <> "A" Or _
    Arr(1) <> "B" Or _
    Arr(2) <> "C" Or _
    Arr(3) <> "D" Or _
    Arr(4) <> "E" Then
        TestInsertionSortS = False
        Debug.Print "Reverse sorted"
    End If

    'Same
    Arr(0) = "A"
    Arr(1) = "A"
    Arr(2) = "A"
    Arr(3) = "A"
    Arr(4) = "A"
    pInsertionSortS Arr, 5, vbBinaryCompare
    If Arr(0) <> "A" Or _
    Arr(1) <> "A" Or _
    Arr(2) <> "A" Or _
    Arr(3) <> "A" Or _
    Arr(4) <> "A" Then
        TestInsertionSortS = False
        Debug.Print "Same"
    End If

    Debug.Print "TestInsertionSortS: " & TestInsertionSortS
    
End Function

Private Function TestMergeSortS() As Boolean

    TestMergeSortS = True

    Dim Arr() As String

    'Unitialized
    On Error Resume Next
    pMergeSortS Arr, 0, vbBinaryCompare
    If Err.Number <> 0 Then
        TestMergeSortS = False
        Debug.Print "Uninitialized"
        Debug.Print "Error " & Err.Number & ": " & Err.Description
    End If
    On Error GoTo 0

    'One
    ReDim Arr(0 To 0) As String
    Arr(0) = "A"
    pMergeSortS Arr, 1, vbBinaryCompare
    If Arr(0) <> "A" Then
        TestMergeSortS = False
        Debug.Print "One"
    End If
    
    ReDim Arr(0 To 4) As String

    'Shuffled
    Arr(0) = "C"
    Arr(1) = "A"
    Arr(2) = "D"
    Arr(3) = "E"
    Arr(4) = "B"
    pMergeSortS Arr, 5, vbBinaryCompare
    If Arr(0) <> "A" Or _
    Arr(1) <> "B" Or _
    Arr(2) <> "C" Or _
    Arr(3) <> "D" Or _
    Arr(4) <> "E" Then
        TestMergeSortS = False
        Debug.Print "Shuffled"
    End If

    'Sorted
    Arr(0) = "A"
    Arr(1) = "B"
    Arr(2) = "C"
    Arr(3) = "D"
    Arr(4) = "E"
    pMergeSortS Arr, 5, vbBinaryCompare
    If Arr(0) <> "A" Or _
    Arr(1) <> "B" Or _
    Arr(2) <> "C" Or _
    Arr(3) <> "D" Or _
    Arr(4) <> "E" Then
        TestMergeSortS = False
        Debug.Print "Sorted"
    End If

    'Reverse Sorted
    Arr(0) = "E"
    Arr(1) = "D"
    Arr(2) = "C"
    Arr(3) = "B"
    Arr(4) = "A"
    pMergeSortS Arr, 5, vbBinaryCompare
    If Arr(0) <> "A" Or _
    Arr(1) <> "B" Or _
    Arr(2) <> "C" Or _
    Arr(3) <> "D" Or _
    Arr(4) <> "E" Then
        TestMergeSortS = False
        Debug.Print "Reverse sorted"
    End If

    'Same
    Arr(0) = "A"
    Arr(1) = "A"
    Arr(2) = "A"
    Arr(3) = "A"
    Arr(4) = "A"
    pMergeSortS Arr, 5, vbBinaryCompare
    If Arr(0) <> "A" Or _
    Arr(1) <> "A" Or _
    Arr(2) <> "A" Or _
    Arr(3) <> "A" Or _
    Arr(4) <> "A" Then
        TestMergeSortS = False
        Debug.Print "Same"
    End If

    Debug.Print "TestMergeSortS: " & TestMergeSortS
    
End Function

Private Function TestSortObjectArray() As Boolean

    TestSortObjectArray = True

    'Not Object
    Dim NOArr(0 To 0)
    NOArr(0) = 1
    On Error Resume Next
    SortObjectArray NOArr, "Count", True
    If Err.Number <> 5 Then
        TestSortObjectArray = False
        Debug.Print "Not Object"
    End If
    On Error GoTo 0
    
    Dim Arr() As Object

    'Unitialized
    On Error Resume Next
    SortObjectArray Arr, "Count", True
    If Err.Number <> 0 Then
        TestSortObjectArray = False
        Debug.Print "Uninitialized"
        Debug.Print "Error " & Err.Number & ": " & Err.Description
    End If
    On Error GoTo 0

    'One
    ReDim Arr(0 To 0) As Object
    Set Arr(0) = New Collection
    Arr(0).Add 1
    SortObjectArray Arr, "Count", True
    If Arr(0).Item(1) <> 1 Then
        TestSortObjectArray = False
        Debug.Print "One"
    End If
    
    ReDim Arr(0 To 4) As Object

    'Shuffled
    Set Arr(0) = New Collection
    Arr(0).Add 1
    Arr(0).Add 1
    Arr(0).Add 1
    Set Arr(1) = New Collection
    Arr(1).Add 1
    Set Arr(2) = New Collection
    Arr(2).Add 1
    Arr(2).Add 1
    Arr(2).Add 1
    Arr(2).Add 1
    Set Arr(3) = New Collection
    Arr(3).Add 1
    Arr(3).Add 1
    Arr(3).Add 1
    Arr(3).Add 1
    Arr(3).Add 1
    Set Arr(4) = New Collection
    Arr(4).Add 1
    Arr(4).Add 1
    SortObjectArray Arr, "Count", True
    If Arr(0).Count <> 1 Or _
    Arr(1).Count <> 2 Or _
    Arr(2).Count <> 3 Or _
    Arr(3).Count <> 4 Or _
    Arr(4).Count <> 5 Then
        TestSortObjectArray = False
        Debug.Print "Shuffled"
    End If

    'Sorted
    Set Arr(0) = New Collection
    Arr(0).Add 1
    Set Arr(1) = New Collection
    Arr(1).Add 1
    Arr(1).Add 1
    Set Arr(2) = New Collection
    Arr(2).Add 1
    Arr(2).Add 1
    Arr(2).Add 1
    Set Arr(3) = New Collection
    Arr(3).Add 1
    Arr(3).Add 1
    Arr(3).Add 1
    Arr(3).Add 1
    Set Arr(4) = New Collection
    Arr(4).Add 1
    Arr(4).Add 1
    Arr(4).Add 1
    Arr(4).Add 1
    Arr(4).Add 1
    SortObjectArray Arr, "Count", True
    If Arr(0).Count <> 1 Or _
    Arr(1).Count <> 2 Or _
    Arr(2).Count <> 3 Or _
    Arr(3).Count <> 4 Or _
    Arr(4).Count <> 5 Then
        TestSortObjectArray = False
        Debug.Print "Sorted"
    End If

    'Reverse Sorted
    Set Arr(0) = New Collection
    Arr(0).Add 1
    Arr(0).Add 1
    Arr(0).Add 1
    Arr(0).Add 1
    Arr(0).Add 1
    Set Arr(1) = New Collection
    Arr(1).Add 1
    Arr(1).Add 1
    Arr(1).Add 1
    Arr(1).Add 1
    Set Arr(2) = New Collection
    Arr(2).Add 1
    Arr(2).Add 1
    Arr(2).Add 1
    Set Arr(3) = New Collection
    Arr(3).Add 1
    Arr(3).Add 1
    Set Arr(4) = New Collection
    Arr(4).Add 1
    SortObjectArray Arr, "Count", True
    If Arr(0).Count <> 1 Or _
    Arr(1).Count <> 2 Or _
    Arr(2).Count <> 3 Or _
    Arr(3).Count <> 4 Or _
    Arr(4).Count <> 5 Then
        TestSortObjectArray = False
        Debug.Print "Reverse sorted"
    End If

    'Same
    Set Arr(0) = New Collection
    Arr(0).Add 1
    Set Arr(1) = New Collection
    Arr(1).Add 1
    Set Arr(2) = New Collection
    Arr(2).Add 1
    Set Arr(3) = New Collection
    Arr(3).Add 1
    Set Arr(4) = New Collection
    Arr(4).Add 1
    SortObjectArray Arr, "Count", True
    If Arr(0).Count <> 1 Or _
    Arr(1).Count <> 1 Or _
    Arr(2).Count <> 1 Or _
    Arr(3).Count <> 1 Or _
    Arr(4).Count <> 1 Then
        TestSortObjectArray = False
        Debug.Print "Same"
    End If

    Debug.Print "TestSortObjectArray: " & TestSortObjectArray
    
End Function

Private Function TestInsertionSortO() As Boolean

    TestInsertionSortO = True

    Dim Arr() As Object

    'Unitialized
    On Error Resume Next
    pInsertionSortO Arr, 0, "Count", VbMethod
    If Err.Number <> 0 Then
        TestInsertionSortO = False
        Debug.Print "Uninitialized"
        Debug.Print "Error " & Err.Number & ": " & Err.Description
    End If
    On Error GoTo 0

    'One
    ReDim Arr(0 To 0) As Object
    Set Arr(0) = New Collection
    Arr(0).Add 1
    pInsertionSortO Arr, 1, "Count", VbMethod
    If Arr(0).Item(1) <> 1 Then
        TestInsertionSortO = False
        Debug.Print "One"
    End If
    
    ReDim Arr(0 To 4) As Object

    'Shuffled
    Set Arr(0) = New Collection
    Arr(0).Add 1
    Arr(0).Add 1
    Arr(0).Add 1
    Set Arr(1) = New Collection
    Arr(1).Add 1
    Set Arr(2) = New Collection
    Arr(2).Add 1
    Arr(2).Add 1
    Arr(2).Add 1
    Arr(2).Add 1
    Set Arr(3) = New Collection
    Arr(3).Add 1
    Arr(3).Add 1
    Arr(3).Add 1
    Arr(3).Add 1
    Arr(3).Add 1
    Set Arr(4) = New Collection
    Arr(4).Add 1
    Arr(4).Add 1
    pInsertionSortO Arr, 5, "Count", VbMethod
    If Arr(0).Count <> 1 Or _
    Arr(1).Count <> 2 Or _
    Arr(2).Count <> 3 Or _
    Arr(3).Count <> 4 Or _
    Arr(4).Count <> 5 Then
        TestInsertionSortO = False
        Debug.Print "Shuffled"
    End If

    'Sorted
    Set Arr(0) = New Collection
    Arr(0).Add 1
    Set Arr(1) = New Collection
    Arr(1).Add 1
    Arr(1).Add 1
    Set Arr(2) = New Collection
    Arr(2).Add 1
    Arr(2).Add 1
    Arr(2).Add 1
    Set Arr(3) = New Collection
    Arr(3).Add 1
    Arr(3).Add 1
    Arr(3).Add 1
    Arr(3).Add 1
    Set Arr(4) = New Collection
    Arr(4).Add 1
    Arr(4).Add 1
    Arr(4).Add 1
    Arr(4).Add 1
    Arr(4).Add 1
    pInsertionSortO Arr, 5, "Count", VbMethod
    If Arr(0).Count <> 1 Or _
    Arr(1).Count <> 2 Or _
    Arr(2).Count <> 3 Or _
    Arr(3).Count <> 4 Or _
    Arr(4).Count <> 5 Then
        TestInsertionSortO = False
        Debug.Print "Sorted"
    End If

    'Reverse Sorted
    Set Arr(0) = New Collection
    Arr(0).Add 1
    Arr(0).Add 1
    Arr(0).Add 1
    Arr(0).Add 1
    Arr(0).Add 1
    Set Arr(1) = New Collection
    Arr(1).Add 1
    Arr(1).Add 1
    Arr(1).Add 1
    Arr(1).Add 1
    Set Arr(2) = New Collection
    Arr(2).Add 1
    Arr(2).Add 1
    Arr(2).Add 1
    Set Arr(3) = New Collection
    Arr(3).Add 1
    Arr(3).Add 1
    Set Arr(4) = New Collection
    Arr(4).Add 1
    pInsertionSortO Arr, 5, "Count", VbMethod
    If Arr(0).Count <> 1 Or _
    Arr(1).Count <> 2 Or _
    Arr(2).Count <> 3 Or _
    Arr(3).Count <> 4 Or _
    Arr(4).Count <> 5 Then
        TestInsertionSortO = False
        Debug.Print "Reverse sorted"
    End If

    'Same
    Set Arr(0) = New Collection
    Arr(0).Add 1
    Set Arr(1) = New Collection
    Arr(1).Add 1
    Set Arr(2) = New Collection
    Arr(2).Add 1
    Set Arr(3) = New Collection
    Arr(3).Add 1
    Set Arr(4) = New Collection
    Arr(4).Add 1
    pInsertionSortO Arr, 5, "Count", VbMethod
    If Arr(0).Count <> 1 Or _
    Arr(1).Count <> 1 Or _
    Arr(2).Count <> 1 Or _
    Arr(3).Count <> 1 Or _
    Arr(4).Count <> 1 Then
        TestInsertionSortO = False
        Debug.Print "Same"
    End If

    Debug.Print "TestInsertionSortO: " & TestInsertionSortO
    
End Function

Private Function TestMergeSortO() As Boolean

    TestMergeSortO = True

    Dim Arr() As Object

    'Unitialized
    On Error Resume Next
    pMergeSortO Arr, 0, "Count", VbMethod
    If Err.Number <> 0 Then
        TestMergeSortO = False
        Debug.Print "Uninitialized"
        Debug.Print "Error " & Err.Number & ": " & Err.Description
    End If
    On Error GoTo 0

    'One
    ReDim Arr(0 To 0) As Object
    Set Arr(0) = New Collection
    Arr(0).Add 1
    pMergeSortO Arr, 1, "Count", VbMethod
    If Arr(0).Item(1) <> 1 Then
        TestMergeSortO = False
        Debug.Print "One"
    End If
    
    ReDim Arr(0 To 4) As Object

    'Shuffled
    Set Arr(0) = New Collection
    Arr(0).Add 1
    Arr(0).Add 1
    Arr(0).Add 1
    Set Arr(1) = New Collection
    Arr(1).Add 1
    Set Arr(2) = New Collection
    Arr(2).Add 1
    Arr(2).Add 1
    Arr(2).Add 1
    Arr(2).Add 1
    Set Arr(3) = New Collection
    Arr(3).Add 1
    Arr(3).Add 1
    Arr(3).Add 1
    Arr(3).Add 1
    Arr(3).Add 1
    Set Arr(4) = New Collection
    Arr(4).Add 1
    Arr(4).Add 1
    pMergeSortO Arr, 5, "Count", VbMethod
    If Arr(0).Count <> 1 Or _
    Arr(1).Count <> 2 Or _
    Arr(2).Count <> 3 Or _
    Arr(3).Count <> 4 Or _
    Arr(4).Count <> 5 Then
        TestMergeSortO = False
        Debug.Print "Shuffled"
    End If

    'Sorted
    Set Arr(0) = New Collection
    Arr(0).Add 1
    Set Arr(1) = New Collection
    Arr(1).Add 1
    Arr(1).Add 1
    Set Arr(2) = New Collection
    Arr(2).Add 1
    Arr(2).Add 1
    Arr(2).Add 1
    Set Arr(3) = New Collection
    Arr(3).Add 1
    Arr(3).Add 1
    Arr(3).Add 1
    Arr(3).Add 1
    Set Arr(4) = New Collection
    Arr(4).Add 1
    Arr(4).Add 1
    Arr(4).Add 1
    Arr(4).Add 1
    Arr(4).Add 1
    pMergeSortO Arr, 5, "Count", VbMethod
    If Arr(0).Count <> 1 Or _
    Arr(1).Count <> 2 Or _
    Arr(2).Count <> 3 Or _
    Arr(3).Count <> 4 Or _
    Arr(4).Count <> 5 Then
        TestMergeSortO = False
        Debug.Print "Sorted"
    End If

    'Reverse Sorted
    Set Arr(0) = New Collection
    Arr(0).Add 1
    Arr(0).Add 1
    Arr(0).Add 1
    Arr(0).Add 1
    Arr(0).Add 1
    Set Arr(1) = New Collection
    Arr(1).Add 1
    Arr(1).Add 1
    Arr(1).Add 1
    Arr(1).Add 1
    Set Arr(2) = New Collection
    Arr(2).Add 1
    Arr(2).Add 1
    Arr(2).Add 1
    Set Arr(3) = New Collection
    Arr(3).Add 1
    Arr(3).Add 1
    Set Arr(4) = New Collection
    Arr(4).Add 1
    pMergeSortO Arr, 5, "Count", VbMethod
    If Arr(0).Count <> 1 Or _
    Arr(1).Count <> 2 Or _
    Arr(2).Count <> 3 Or _
    Arr(3).Count <> 4 Or _
    Arr(4).Count <> 5 Then
        TestMergeSortO = False
        Debug.Print "Reverse sorted"
    End If

    'Same
    Set Arr(0) = New Collection
    Arr(0).Add 1
    Set Arr(1) = New Collection
    Arr(1).Add 1
    Set Arr(2) = New Collection
    Arr(2).Add 1
    Set Arr(3) = New Collection
    Arr(3).Add 1
    Set Arr(4) = New Collection
    Arr(4).Add 1
    pMergeSortO Arr, 5, "Count", VbMethod
    If Arr(0).Count <> 1 Or _
    Arr(1).Count <> 1 Or _
    Arr(2).Count <> 1 Or _
    Arr(3).Count <> 1 Or _
    Arr(4).Count <> 1 Then
        TestMergeSortO = False
        Debug.Print "Same"
    End If

    Debug.Print "TestMergeSortO: " & TestMergeSortO
    
End Function

Private Function TestIsPrime() As Boolean
    
    TestIsPrime = True
    
    Dim b As Boolean
    
    '-1
    On Error Resume Next
    b = pIsPrime(-1)
    If Err.Number <> 5 Then
        TestIsPrime = False
        Debug.Print "-1"
    End If
    On Error GoTo 0
    
    '0
    On Error Resume Next
    b = pIsPrime(0)
    If Err.Number <> 5 Then
        TestIsPrime = False
        Debug.Print "0"
    End If
    On Error GoTo 0
    
    '1
    If pIsPrime(1) Then
        TestIsPrime = False
        Debug.Print "1"
    End If
    
    '2
    If Not pIsPrime(2) Then
        TestIsPrime = False
        Debug.Print "2"
    End If
    
    '3
    If Not pIsPrime(3) Then
        TestIsPrime = False
        Debug.Print "3"
    End If
    
    '4
    If pIsPrime(4) Then
        TestIsPrime = False
        Debug.Print "4"
    End If
    
    '5
    If Not pIsPrime(5) Then
        TestIsPrime = False
        Debug.Print "5"
    End If
    
    '17
    If Not pIsPrime(17) Then
        TestIsPrime = False
        Debug.Print "17"
    End If
    
    Debug.Print "TestIsPrime: " & TestIsPrime
    
End Function

Private Function TestRandomLong() As Boolean

    TestRandomLong = True
    
    Dim i&
    Dim j&
    
    'Min > Max
    On Error Resume Next
    i = RandomLong(10, 1)
    If Err.Number <> 5 Then
        TestRandomLong = False
        Debug.Print "Min > Max"
    End If
    On Error GoTo 0
    
    Dim F1 As Boolean
    Dim F2 As Boolean
    Dim F3 As Boolean
    Dim F4 As Boolean
    Dim F5 As Boolean
    For i = 1 To 100
        j = RandomLong(1, 5)
        Select Case j
            Case 1
                F1 = True
            Case 2
                F2 = True
            Case 3
                F3 = True
            Case 4
                F4 = True
            Case 5
                F5 = True
            Case Else
                TestRandomLong = False
                Debug.Print "Bounds"
        End Select
    Next i
    If Not (F1 And F2 And F3 And F4 And F5) Then
        TestRandomLong = False
        Debug.Print "Not all there"
    End If
    
    Debug.Print "TestRandomLong: " & TestRandomLong
    
End Function

Private Function TestRandomDouble() As Boolean

    TestRandomDouble = True
    
    Dim i&
    Dim j&
    
    'Min > Max
    On Error Resume Next
    i = RandomLong(10, 1)
    If Err.Number <> 5 Then
        TestRandomDouble = False
        Debug.Print "Min > Max"
    End If
    On Error GoTo 0
    
    Dim F1 As Boolean
    Dim F2 As Boolean
    Dim F3 As Boolean
    Dim F4 As Boolean
    Dim F5 As Boolean
    For i = 1 To 100
        j = RandomLong(1, 5)
        Select Case CLng(j)
            Case 1
                F1 = True
            Case 2
                F2 = True
            Case 3
                F3 = True
            Case 4
                F4 = True
            Case 5
                F5 = True
            Case Else
                TestRandomDouble = False
                Debug.Print "Bounds"
        End Select
    Next i
    If Not (F1 And F2 And F3 And F4 And F5) Then
        TestRandomDouble = False
        Debug.Print "Not all there"
    End If
    
    Debug.Print "TestRandomDouble: " & TestRandomDouble
    
End Function

Private Function TestToByteArray() As Boolean

    TestToByteArray = True
    
    Dim Arr()
    Dim BArr() As Byte
    
    'Empty
    On Error Resume Next
    BArr = ToByteArray(Arr)
    If Err.Number <> 0 Then
        TestToByteArray = False
        Debug.Print "Empty"
    End If
    On Error GoTo 0
    
    'Assign
    Arr = Array(1, 2, 3)
    On Error Resume Next
    BArr = ToByteArray(Arr)
    If Err.Number <> 0 Then
        TestToByteArray = False
        Debug.Print "Assign"
    End If
    On Error GoTo 0
    
    'Value
    If BArr(0) <> 1 Then
        TestToByteArray = False
        Debug.Print "Value"
    End If
    
    'Overflow
    Arr = Array(256, 257, 258)
    On Error Resume Next
    BArr = ToByteArray(Arr)
    If Err.Number <> 6 Then
        TestToByteArray = False
        Debug.Print "Overflow"
    End If
    On Error GoTo 0
    
    Debug.Print "TestToByteArray: " & TestToByteArray
    
End Function

Private Function TestToIntegerArray() As Boolean

    TestToIntegerArray = True
    
    Dim Arr()
    Dim IArr() As Integer
    
    'Empty
    On Error Resume Next
    IArr = ToIntegerArray(Arr)
    If Err.Number <> 0 Then
        TestToIntegerArray = False
        Debug.Print "Empty"
    End If
    On Error GoTo 0
    
    'Assign
    Arr = Array(1, 2, 3)
    On Error Resume Next
    IArr = ToIntegerArray(Arr)
    If Err.Number <> 0 Then
        TestToIntegerArray = False
        Debug.Print "Assign"
    End If
    On Error GoTo 0
    
    'Value
    If IArr(0) <> 1 Then
        TestToIntegerArray = False
        Debug.Print "Value"
    End If
    
    'Overflow
    Arr = Array(32767, 32768, 32769)
    On Error Resume Next
    IArr = ToIntegerArray(Arr)
    If Err.Number <> 6 Then
        TestToIntegerArray = False
        Debug.Print "Overflow"
    End If
    On Error GoTo 0
    
    Debug.Print "TestToIntegerArray: " & TestToIntegerArray
    
End Function

Private Function TestToLongArray() As Boolean

    TestToLongArray = True
    
    Dim Arr()
    Dim LArr() As Long
    
    'Empty
    On Error Resume Next
    LArr = ToLongArray(Arr)
    If Err.Number <> 0 Then
        TestToLongArray = False
        Debug.Print "Empty"
    End If
    On Error GoTo 0
    
    'Assign
    Arr = Array(1, 2, 3)
    On Error Resume Next
    LArr = ToLongArray(Arr)
    If Err.Number <> 0 Then
        TestToLongArray = False
        Debug.Print "Assign"
    End If
    On Error GoTo 0
    
    'Value
    If LArr(0) <> 1 Then
        TestToLongArray = False
        Debug.Print "Value"
    End If
    
    'Overflow
    Arr = Array(2147483648#, 2147483649#, 2147483650#)
    On Error Resume Next
    LArr = ToLongArray(Arr)
    If Err.Number <> 6 Then
        TestToLongArray = False
        Debug.Print "Overflow"
    End If
    On Error GoTo 0
    
    Debug.Print "TestToLongArray: " & TestToLongArray
    
End Function

#If VBA7 = 1 And Win64 = 1 Then
Private Function TestToLongLongArray() As Boolean

    TestToLongLongArray = True
    
    Dim Arr()
    Dim LLArr() As LongLong
    
    'Empty
    On Error Resume Next
    LLArr = ToLongLongArray(Arr)
    If Err.Number <> 0 Then
        TestToLongLongArray = False
        Debug.Print "Empty"
    End If
    On Error GoTo 0
    
    'Assign
    Arr = Array(1, 2, 3)
    On Error Resume Next
    LLArr = ToLongLongArray(Arr)
    If Err.Number <> 0 Then
        TestToLongLongArray = False
        Debug.Print "Assign"
    End If
    On Error GoTo 0
    
    'Value
    If LLArr(0) <> 1 Then
        TestToLongLongArray = False
        Debug.Print "Value"
    End If
    
    'Overflow
    Arr = Array(CDec("9223372036854775808"), _
    CDec("9223372036854775809"), _
    CDec("9223372036854775810"))
    On Error Resume Next
    LLArr = ToLongLongArray(Arr)
    If Err.Number <> 6 Then
        TestToLongLongArray = False
        Debug.Print "Overflow"
    End If
    On Error GoTo 0
    
    Debug.Print "TestToLongLongArray: " & TestToLongLongArray
    
End Function
#End If

#If VBA7 = 1 Then
Private Function TestToLongPtrArray() As Boolean

    TestToLongPtrArray = True
    
    Dim Arr()
    Dim LPArr() As LongPtr
    
#If Win64 = 1 Then

    'Empty
    On Error Resume Next
    LPArr = ToLongPtrArray(Arr)
    If Err.Number <> 0 Then
        TestToLongPtrArray = False
        Debug.Print "Empty"
    End If
    On Error GoTo 0
    
    'Assign
    Arr = Array(1, 2, 3)
    On Error Resume Next
    LPArr = ToLongPtrArray(Arr)
    If Err.Number <> 0 Then
        TestToLongPtrArray = False
        Debug.Print "Assign"
    End If
    On Error GoTo 0
    
    'Value
    If LPArr(0) <> 1 Then
        TestToLongPtrArray = False
        Debug.Print "Value"
    End If
    
    'Overflow
    Arr = Array(CDec("9223372036854775808"), _
    CDec("9223372036854775809"), _
    CDec("9223372036854775810"))
    On Error Resume Next
    LPArr = ToLongPtrArray(Arr)
    If Err.Number <> 6 Then
        TestToLongPtrArray = False
        Debug.Print "Overflow"
    End If
    On Error GoTo 0
    
#Else

    'Empty
    On Error Resume Next
    LPArr = ToLongPtrArray(Arr)
    If Err.Number <> 0 Then
        TestToLongPtrArray = False
        Debug.Print "Empty"
    End If
    On Error GoTo 0
    
    'Assign
    Arr = Array(1, 2, 3)
    On Error Resume Next
    LPArr = ToLongPtrArray(Arr)
    If Err.Number <> 0 Then
        TestToLongPtrArray = False
        Debug.Print "Assign"
    End If
    On Error GoTo 0
    
    'Value
    If LPArr(0) <> 1 Then
        TestToLongPtrArray = False
        Debug.Print "Value"
    End If
    
    'Overflow
    Arr = Array(2147483648#, 2147483649#, 2147483650#)
    On Error Resume Next
    LPArr = ToLongPtrArray(Arr)
    If Err.Number <> 6 Then
        TestToLongPtrArray = False
        Debug.Print "Overflow"
    End If
    On Error GoTo 0
    
#End If
    
    Debug.Print "TestToLongPtrArray: " & TestToLongPtrArray
    
End Function
#End If

Private Function TestToSingleArray() As Boolean

    TestToSingleArray = True
    
    Dim Arr()
    Dim SArr() As Single
    
    'Empty
    On Error Resume Next
    SArr = ToSingleArray(Arr)
    If Err.Number <> 0 Then
        TestToSingleArray = False
        Debug.Print "Empty"
    End If
    On Error GoTo 0
    
    'Assign
    Arr = Array(1.5, 2.5, 3.5)
    On Error Resume Next
    SArr = ToSingleArray(Arr)
    If Err.Number <> 0 Then
        TestToSingleArray = False
        Debug.Print "Assign"
    End If
    On Error GoTo 0
    
    'Value
    If SArr(0) <> 1.5 Then
        TestToSingleArray = False
        Debug.Print "Value"
    End If
    
    'Overflow
    Arr = Array(3.402823E+39, 3.402823E+40, 3.402823E+41)
    On Error Resume Next
    SArr = ToSingleArray(Arr)
    If Err.Number <> 6 Then
        TestToSingleArray = False
        Debug.Print "Overflow"
    End If
    On Error GoTo 0
    
    Debug.Print "TestToSingleArray: " & TestToSingleArray
    
End Function

Private Function TestToDoubleArray() As Boolean

    TestToDoubleArray = True
    
    Dim Arr()
    Dim DArr() As Double
    
    'Empty
    On Error Resume Next
    DArr = ToDoubleArray(Arr)
    If Err.Number <> 0 Then
        TestToDoubleArray = False
        Debug.Print "Empty"
    End If
    On Error GoTo 0
    
    'Assign
    Arr = Array(1.5, 2.5, 3.5)
    On Error Resume Next
    DArr = ToDoubleArray(Arr)
    If Err.Number <> 0 Then
        TestToDoubleArray = False
        Debug.Print "Assign"
    End If
    On Error GoTo 0
    
    'Value
    If DArr(0) <> 1.5 Then
        TestToDoubleArray = False
        Debug.Print "Value"
    End If
    
    'Overflow
    
    Debug.Print "TestToDoubleArray: " & TestToDoubleArray
    
End Function

Private Function TestToDecimalArray() As Boolean

    TestToDecimalArray = True
    
    Dim Arr()
    Dim DArr() As Variant
    
    'Empty
    On Error Resume Next
    DArr = ToDecimalArray(Arr)
    If Err.Number <> 0 Then
        TestToDecimalArray = False
        Debug.Print "Empty"
    End If
    On Error GoTo 0
    
    'Assign
    Arr = Array(1.5, 2.5, 3.5)
    On Error Resume Next
    DArr = ToDecimalArray(Arr)
    If Err.Number <> 0 Then
        TestToDecimalArray = False
        Debug.Print "Assign"
    End If
    On Error GoTo 0
    
    'Value
    If DArr(0) <> 1.5 Then
        TestToDecimalArray = False
        Debug.Print "Value"
    End If
    
    'Overflow
    Arr = Array(CDbl("1.79769313486232E+305"), _
    CDbl("1.79769313486232E+306"), _
    CDbl("1.79769313486232E+307"))
    On Error Resume Next
    DArr = ToDecimalArray(Arr)
    If Err.Number <> 6 Then
        TestToDecimalArray = False
        Debug.Print "Overflow"
    End If
    On Error GoTo 0
    
    Debug.Print "TestToDecimalArray: " & TestToDecimalArray
    
End Function

Private Function TestToCurrencyArray() As Boolean

    TestToCurrencyArray = True
    
    Dim Arr()
    Dim CArr() As Currency
    
    'Empty
    On Error Resume Next
    CArr = ToCurrencyArray(Arr)
    If Err.Number <> 0 Then
        TestToCurrencyArray = False
        Debug.Print "Empty"
    End If
    On Error GoTo 0
    
    'Assign
    Arr = Array(1.5, 2.5, 3.5)
    On Error Resume Next
    CArr = ToCurrencyArray(Arr)
    If Err.Number <> 0 Then
        TestToCurrencyArray = False
        Debug.Print "Assign"
    End If
    On Error GoTo 0
    
    'Value
    If CArr(0) <> 1.5 Then
        TestToCurrencyArray = False
        Debug.Print "Value"
    End If
    
    'Overflow
    Arr = Array(CDbl("1.79769313486232E+305"), _
    CDbl("1.79769313486232E+306"), _
    CDbl("1.79769313486232E+307"))
    On Error Resume Next
    CArr = ToCurrencyArray(Arr)
    If Err.Number <> 6 Then
        TestToCurrencyArray = False
        Debug.Print "Overflow"
    End If
    On Error GoTo 0
    
    'Overflow LongLong
    Arr = Array(922337203685478^, _
    922337203685479^, _
    922337203685480^)
    On Error Resume Next
    CArr = ToCurrencyArray(Arr)
    If Err.Number <> 6 Then
        TestToCurrencyArray = False
        Debug.Print "Overflow"
    End If
    On Error GoTo 0
    Debug.Print "TestToCurrencyArray: " & TestToCurrencyArray
    
End Function

Private Function TestToDateArray() As Boolean

    TestToDateArray = True

    Dim Arr()
    Dim DArr() As Date

    'Empty
    On Error Resume Next
    DArr = ToDateArray(Arr)
    If Err.Number <> 0 Then
        TestToDateArray = False
        Debug.Print "Empty"
    End If
    On Error GoTo 0

    'Assign
    Arr = Array("1/01/2020", "1/01/2021", "1/01/2022")
    On Error Resume Next
    DArr = ToDateArray(Arr)
    If Err.Number <> 0 Then
        TestToDateArray = False
        Debug.Print "Assign"
    End If
    On Error GoTo 0
    
    'Value
    If DArr(0) <> #1/1/2020# Then
        TestToDateArray = False
        Debug.Print "Value"
    End If
    
    Debug.Print "TestToDateArray: " & TestToDateArray

End Function

Private Function TestToBooleanArray() As Boolean

    TestToBooleanArray = True

    Dim Arr()
    Dim BArr() As Boolean

    'Empty
    On Error Resume Next
    BArr = ToBooleanArray(Arr)
    If Err.Number <> 0 Then
        TestToBooleanArray = False
        Debug.Print "Empty"
    End If
    On Error GoTo 0

    'Assign
    Arr = Array(1, 2, 3)
    On Error Resume Next
    BArr = ToBooleanArray(Arr)
    If Err.Number <> 0 Then
        TestToBooleanArray = False
        Debug.Print "Assign"
    End If
    On Error GoTo 0

    'Value
    If BArr(0) <> True Then
        TestToBooleanArray = False
        Debug.Print "Value"
    End If
    
    Debug.Print "TestToBooleanArray: " & TestToBooleanArray
    
End Function

Private Function TestToStringArray() As Boolean

    TestToStringArray = True

    Dim Arr()
    Dim SArr() As String

    'Empty
    On Error Resume Next
    SArr = ToStringArray(Arr)
    If Err.Number <> 0 Then
        TestToStringArray = False
        Debug.Print "Empty"
    End If
    On Error GoTo 0

    'Assign
    Arr = Array(1, 2, 3)
    On Error Resume Next
    SArr = ToStringArray(Arr)
    If Err.Number <> 0 Then
        TestToStringArray = False
        Debug.Print "Assign"
    End If
    On Error GoTo 0

    'Value
    If SArr(0) <> "1" Then
        TestToStringArray = False
        Debug.Print "Value"
    End If
    
    Debug.Print "TestToStringArray: " & TestToStringArray
    
End Function

Private Function TestToVariantArray() As Boolean

    TestToVariantArray = True

    Dim Arr()
    Dim VArr() As Variant

    'Empty
    On Error Resume Next
    VArr = ToVariantArray(Arr)
    If Err.Number <> 0 Then
        TestToVariantArray = False
        Debug.Print "Empty"
    End If
    On Error GoTo 0

    'Assign
    Arr = Array(1, 2, 3)
    On Error Resume Next
    VArr = ToVariantArray(Arr)
    If Err.Number <> 0 Then
        TestToVariantArray = False
        Debug.Print "Assign"
    End If
    On Error GoTo 0

    'Value
    If VArr(0) <> 1 Then
        TestToVariantArray = False
        Debug.Print "Value"
    End If
    
    Debug.Print "TestToVariantArray: " & TestToVariantArray
    
End Function
