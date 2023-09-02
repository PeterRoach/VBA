Attribute VB_Name = "modCollection"
Option Explicit

'Meta Data=============================================================
'======================================================================

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'  Copyright © 2023 Peter D Roach. All Rights Reserved.
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
'  Module Name: modCollection
'  Module Description: Contains functions for working with Collections.
'  Module Version: 1.0
'  Module License: MIT
'  Module Author: Peter Roach; PeterRoach.Code@gmail.com
'  Module Contents:
'  ----------------------------------------
'    Public Procedures:
'        CollectionToStringArray
'        JoinCollection
'        ValidateCollectionType
'   Test Procedures:
'        TestmodCollection
'        TestCollectionToStringArray
'        TestJoinCollection
'        TestValidateCollectionType
'  ----------------------------------------

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'  Example Usage:

Private Sub Example()

    Dim Coll As Collection
    Set Coll = New Collection
    Coll.Add "A"
    Coll.Add "B"
    Coll.Add "C"

    Dim Arr$()
    Arr = CollectionToStringArray(Coll)
    Debug.Print Join(Arr, ",")

    Debug.Print JoinCollection(Coll, ",")

    Coll.Add New Collection

    Debug.Print ValidateCollectionType(Coll, "Collection")

    Set Coll = New Collection
    Coll.Add New Collection
    Coll.Add New Collection
    Coll.Add New Collection

    Debug.Print ValidateCollectionType(Coll, "Collection")

End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


'Functions=============================================================
'======================================================================

Public Function CollectionToStringArray(Coll As Collection) As String()
    Dim Arr$()
    Dim Item
    Dim i&
    If Coll.Count > 0 Then
        ReDim Arr(0 To Coll.Count - 1)
        For Each Item In Coll
            Arr(i) = CStr(Item)
            i = i + 1
        Next Item
    End If
    CollectionToStringArray = Arr
End Function

Public Function JoinCollection$(Coll As Collection, Optional Delimiter$ = " ")
    Dim Arr$()
    Arr = CollectionToStringArray(Coll)
    JoinCollection = Join(Arr, Delimiter)
End Function

Public Function ValidateCollectionType(Coll As Collection, DataType$) As Boolean
    Dim i&
    Dim Item
    For Each Item In Coll
        If TypeName(Item) <> DataType Then
            Exit Function
        End If
    Next Item
    ValidateCollectionType = True
End Function


'Tests=================================================================
'======================================================================

Private Function TestmodCollection() As Boolean

    TestmodCollection = _
        TestCollectionToStringArray And _
        TestJoinCollection And _
        TestValidateCollectionType

    Debug.Print "TestmodCollection: " & TestmodCollection

End Function

Private Function TestCollectionToStringArray() As Boolean

    TestCollectionToStringArray = True

    Dim Coll As Collection
    Dim Arr$()

    Set Coll = New Collection
    Coll.Add "A"
    Coll.Add "B"
    Coll.Add "C"
    Arr = CollectionToStringArray(Coll)

    If Arr(0) <> "A" Or Arr(1) <> "B" Or Arr(2) <> "C" Then
        TestCollectionToStringArray = False
        Debug.Print "Values"
    End If

    If LBound(Arr) <> 0 Or UBound(Arr) <> 2 Then
        TestCollectionToStringArray = False
        Debug.Print "Bounds"
    End If

    Set Coll = New Collection
    Coll.Add "A"
    Coll.Add 1
    Coll.Add True
    Coll.Add 3.5
    Arr = CollectionToStringArray(Coll)

    If Arr(0) <> "A" Or Arr(1) <> "1" Or Arr(2) <> "True" Or Arr(3) <> "3.5" Then
        TestCollectionToStringArray = False
        Debug.Print "Values Different Types"
    End If

    If LBound(Arr) <> 0 Or UBound(Arr) <> 3 Then
        TestCollectionToStringArray = False
        Debug.Print "Bounds Different Types"
    End If

    Set Coll = New Collection
    On Error Resume Next
    Arr = CollectionToStringArray(Coll)
    If Err.Number <> 0 Then
        TestCollectionToStringArray = False
        Debug.Print "Empty Collection"
    End If
    On Error GoTo 0

    Set Coll = New Collection
    Coll.Add New Collection
    On Error Resume Next
    Arr = CollectionToStringArray(Coll)
    If Err.Number <> 450 Then
        TestCollectionToStringArray = False
        Debug.Print "Object"
    End If
    On Error GoTo 0

    Debug.Print "TestCollectionToStringArray: " & TestCollectionToStringArray

End Function

Private Function TestJoinCollection() As Boolean

    TestJoinCollection = True

    Dim Coll As Collection

    Set Coll = New Collection
    If JoinCollection(Coll, ",") <> "" Then
        TestJoinCollection = False
        Debug.Print "Empty"
    End If

    Set Coll = New Collection
    Coll.Add "A"
    Coll.Add "B"
    Coll.Add "C"
    If JoinCollection(Coll, ",") <> "A,B,C" Then
        TestJoinCollection = False
        Debug.Print "Normal"
    End If

    Set Coll = New Collection
    Coll.Add "A"
    Coll.Add "1"
    Coll.Add "True"
    Coll.Add "3.5"
    If JoinCollection(Coll, ",") <> "A,1,True,3.5" Then
        TestJoinCollection = False
        Debug.Print "Different Types"
    End If

    Set Coll = New Collection
    Coll.Add "A"
    Coll.Add New Collection
    On Error Resume Next
    JoinCollection Coll, ","
    If Err.Number <> 450 Then
        TestJoinCollection = False
        Debug.Print "Object"
    End If
    On Error GoTo 0

    Debug.Print "TestJoinCollection: " & TestJoinCollection

End Function

Private Function TestValidateCollectionType() As Boolean

    TestValidateCollectionType = True
    
    Dim Coll As Collection

    'True

    Set Coll = New Collection
    Coll.Add 1
    Coll.Add 2
    If Not ValidateCollectionType(Coll, "Integer") Then
        TestValidateCollectionType = False
        Debug.Print "Integer"
    End If

    Set Coll = New Collection
    Coll.Add 1&
    Coll.Add 2&
    If Not ValidateCollectionType(Coll, "Long") Then
        TestValidateCollectionType = False
        Debug.Print "Long"
    End If

    Set Coll = New Collection
    Coll.Add 1#
    Coll.Add 2#
    If Not ValidateCollectionType(Coll, "Double") Then
        TestValidateCollectionType = False
        Debug.Print "Double"
    End If

    Set Coll = New Collection
    Coll.Add True
    Coll.Add False
    If Not ValidateCollectionType(Coll, "Boolean") Then
        TestValidateCollectionType = False
        Debug.Print "Boolean"
    End If

    Set Coll = New Collection
    Coll.Add "1"
    Coll.Add "2"
    If Not ValidateCollectionType(Coll, "String") Then
        TestValidateCollectionType = False
        Debug.Print "String"
    End If

    Set Coll = New Collection
    Coll.Add New Collection
    Coll.Add New Collection
    If Not ValidateCollectionType(Coll, "Collection") Then
        TestValidateCollectionType = False
        Debug.Print "Collection"
    End If
    
    'False

    Set Coll = New Collection
    Coll.Add 1
    Coll.Add 2
    Coll.Add 3&
    If ValidateCollectionType(Coll, "Integer") Then
        TestValidateCollectionType = False
        Debug.Print "Integer False"
    End If

    Set Coll = New Collection
    Coll.Add 1&
    Coll.Add 2&
    Coll.Add 3
    If ValidateCollectionType(Coll, "Long") Then
        TestValidateCollectionType = False
        Debug.Print "Long False"
    End If

    Set Coll = New Collection
    Coll.Add 1#
    Coll.Add 2#
    Coll.Add 3!
    If ValidateCollectionType(Coll, "Double") Then
        TestValidateCollectionType = False
        Debug.Print "Double False"
    End If

    Set Coll = New Collection
    Coll.Add True
    Coll.Add False
    Coll.Add 1
    If ValidateCollectionType(Coll, "Boolean") Then
        TestValidateCollectionType = False
        Debug.Print "Boolean False 1"
    End If
    
    Set Coll = New Collection
    Coll.Add True
    Coll.Add False
    Coll.Add -1
    If ValidateCollectionType(Coll, "Boolean") Then
        TestValidateCollectionType = False
        Debug.Print "Boolean False -1"
    End If
    
    Set Coll = New Collection
    Coll.Add True
    Coll.Add False
    Coll.Add 0
    If ValidateCollectionType(Coll, "Boolean") Then
        TestValidateCollectionType = False
        Debug.Print "Boolean False 0"
    End If

    Set Coll = New Collection
    Coll.Add "1"
    Coll.Add "2"
    Coll.Add 3
    If ValidateCollectionType(Coll, "String") Then
        TestValidateCollectionType = False
        Debug.Print "String False"
    End If

    Set Coll = New Collection
    Coll.Add New Collection
    Coll.Add New Collection
    Coll.Add 1
    If ValidateCollectionType(Coll, "Collection") Then
        TestValidateCollectionType = False
        Debug.Print "Collection False"
    End If

    Debug.Print "TestValidateCollectionType: " & TestValidateCollectionType

End Function
