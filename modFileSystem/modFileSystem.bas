Attribute VB_Name = "modFileSystem"
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
'  Module Name: modFileSystem
'  Module Description: Contains functions for working with the file system.
'  Module Version: 1.0
'  Module License: MIT
'  Module Author: Peter Roach; PeterRoach.Code@gmail.com
'  Module Contents:
'  ----------------------------------------
'    Public Procedures:
'        FolderExists
'        FileExists
'        JoinPath
'    Test Procedures:
'        TestmodFileSystem
'        TestFolderExists
'        TestFileExists
'        TestJoinPath
'  ----------------------------------------

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'  Example Usage:

Private Sub Example()

    Dim FolderPath$
    FolderPath = Environ$("USERPROFILE") & "\Desktop\MyFolder"

    Dim FilePath$
    FilePath = JoinPath(FolderPath, "MyFile.txt")

    If FolderExists(FolderPath) Then
        MsgBox "Folder Exists!"
    Else
        MsgBox "Folder Does Not Exist!"
    End If

    If FileExists(FilePath) Then
        MsgBox "File Exists!"
    Else
        MsgBox "File Does Not Exist!"
    End If

End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


'Functions=============================================================
'======================================================================

Public Function FolderExists(Path$) As Boolean

    If Len(Path) = 0 Then Err.Raise 5

    If Len(Dir(Path, vbDirectory)) > 0 Then

        If (GetAttr(Path) And vbDirectory) <> vbDirectory Then Err.Raise 5

        FolderExists = True

    End If

End Function

Public Function FileExists(Path$) As Boolean

    If Len(Path) = 0 Then Err.Raise 5

    If Len(Dir(Path, vbHidden + vbReadOnly + vbSystem)) > 0 Then

        If (GetAttr(Path) And vbDirectory) = vbDirectory Then Err.Raise 5

        FileExists = True

    End If

End Function

Public Function JoinPath$(ParamArray Parts())

    Dim LB&
    LB = LBound(Parts)

    Dim UB&
    UB = UBound(Parts)

    If UB - LB + 1 = 1 Then

        JoinPath = Parts(LB)

        Exit Function

    End If

    Dim Arr$()
    ReDim Arr(LB To UB)

    Dim i&

    For i = LB To UB

        If Right$(Parts(i), 1) = "\" Then

            Arr(i) = Left$(Parts(i), Len(Parts(i)) - 1)

        Else

            Arr(i) = Parts(i)

        End If

    Next i

    JoinPath = Join(Arr, "\")

End Function


'Tests=================================================================
'======================================================================

Private Function TestmodFileSystem() As Boolean

    TestmodFileSystem = _
        TestFolderExists And _
        TestFileExists And _
        TestJoinPath

    Debug.Print "TestmodFileSystem: " & TestmodFileSystem

End Function

Private Function TestFolderExists() As Boolean

    TestFolderExists = True

    If Not FolderExists("C:") Then
        TestFolderExists = False
        Debug.Print "Drive"
    End If

    If Not FolderExists("C:\") Then
        TestFolderExists = False
        Debug.Print "Drive"
    End If

    If Not FolderExists(Environ$("USERPROFILE") & "\Desktop") Then
        TestFolderExists = False
        Debug.Print "Folder exists"
    End If

    If FolderExists(Environ$("USERPROFILE") & "\Desktop\FakeFolderPathThatDoesNotExist") Then
        TestFolderExists = False
        Debug.Print "Folder does not exist"
    End If

    On Error Resume Next
    FolderExists Environ$("USERPROFILE") & "\Desktop\CreateThisFile.txt"
    If Err.Number <> 5 Then
        TestFolderExists = False
        Debug.Print "File exists"
    End If
    On Error GoTo 0

    On Error Resume Next
    FolderExists Environ$("USERPROFILE") & "\Desktop\FakeFilePathThatDoesNotExist.txt"
    If Err.Number <> 0 Then
        TestFolderExists = False
        Debug.Print "File does not exist"
    End If
    On Error GoTo 0

    On Error Resume Next
    FolderExists ""
    If Err.Number <> 5 Then
        TestFolderExists = False
        Debug.Print "Blank"
    End If
    On Error GoTo 0

    Debug.Print "TestFolderExists: " & TestFolderExists

End Function

Private Function TestFileExists() As Boolean

    TestFileExists = True

    If Not FileExists(Environ$("USERPROFILE") & "\Desktop\CreateThisFile.txt") Then
        TestFileExists = False
        Debug.Print "File exists"
    End If

    If FileExists(Environ$("USERPROFILE") & "\Desktop\FakeFilePathThatDoesNotExist.txt") Then
        TestFileExists = False
        Debug.Print "File does not exist"
    End If

    On Error Resume Next
    FileExists "C:"
    If Err.Number <> 5 Then
        TestFileExists = False
        Debug.Print "Drive"
    End If
    On Error GoTo 0

    On Error Resume Next
    FileExists "C:\"
    If Err.Number <> 5 Then
        TestFileExists = False
        Debug.Print "Drive"
    End If
    On Error GoTo 0

    On Error Resume Next
    FileExists Environ$("USERPROFILE") & "\Desktop"
    If Err.Number <> 0 Then
        TestFileExists = False
        Debug.Print "Folder exists"
    End If
    On Error GoTo 0

    On Error Resume Next
    FileExists Environ$("USERPROFILE") & "\Desktop\FakeFolderPathThatDoesNotExist"
    If Err.Number <> 0 Then
        TestFileExists = False
        Debug.Print "Folder does not exist"
    End If
    On Error GoTo 0

    On Error Resume Next
    FileExists ""
    If Err.Number <> 5 Then
        TestFileExists = False
        Debug.Print "Blank"
    End If
    On Error GoTo 0

    Debug.Print "TestFileExists: " & TestFileExists

End Function

Private Function TestJoinPath() As Boolean

    TestJoinPath = True

    If JoinPath("C:") <> "C:" Then
        TestJoinPath = False
        Debug.Print "One Arg"
    End If

    If JoinPath("C:", "Users", "username", "Desktop") <> "C:\Users\username\Desktop" Then
        TestJoinPath = False
        Debug.Print "Folder"
    End If

    If JoinPath("C:", "Test.txt") <> "C:\Test.txt" Then
        TestJoinPath = False
        Debug.Print "File"
    End If

    Debug.Print "TestJoinPath: " & TestJoinPath

End Function
