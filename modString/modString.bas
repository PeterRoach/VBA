Attribute VB_Name = "modString"
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
'  Module Name: modString
'  Module Description: Contains functions for working with strings.
'  Module Version: 1.1
'  Module License: MIT
'  Module Author: Peter Roach; PeterRoach.Code@gmail.com
'  Module Contents:
'  ----------------------------------------
'    Public Procedures:
'        AscW2
'        CountSubstring
'        SI (String Interpolation)
'        ReplaceNBSP
'        TrimAll
'        RemoveNonPrintableCharacters 'Does not remove new line characters
'        ReplaceNewLineCharacters
'        CleanText
'   Test Procedures:
'        TestmodString
'        TestAscW2
'        TestCountSubstring
'        TestSI
'        TestReplaceNBSP
'        TestTrimAll
'        TestRemoveNonPrintableCharacters
'        TestReplaceNewLineCharacters
'        TestCleanText
'  ----------------------------------------

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'  Example Usage:

Private Sub Example()
    
    Debug.Print AscW2(ChrW(65535))
    
    Debug.Print CountSubstring("Hello, World", "world", vbTextCompare)
    
    Debug.Print SI("Hello {}. My name is {}.", "World", "Peter")
    
    Debug.Print ReplaceNBSP(Chr(160)) = Chr(32)
    
    Debug.Print TrimAll("   Hello,   World    ")
    
    Debug.Print RemoveNonPrintableCharacters("Hello" & Chr(0) & vbNewLine & "World")
    
    Debug.Print ReplaceNewLineCharacters("Hello" & vbNewLine & "World")
    
    Debug.Print CleanText("  Hello   " & Chr(0) & vbNewLine & Chr(160) & "World")

End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


'AscW Fix==============================================================
'======================================================================

'The AscW function in the built-in VBA.Strings module has a problem
'where it returns the correct bit pattern for an unsigned 16-bit
'integer which is incorrect in VBA because VBA uses signed 16-bit
'integer. Thus, after reaching 32767 AscW will start returning negative
'numbers. To work around this issue use one of the functions below.

Public Function AscW2&(Char$)

    AscW2 = AscW(Char) And &HFFFF&

End Function

'Public Function AscW2&(Char$)
'
'    AscW2 = AscW(Char)
'
'    If AscW2 < 0 Then
'        AscW2 = AscW2 + 65536
'    End If
'
'End Function


'General Functions=====================================================
'======================================================================

Public Function CountSubstring&(Text$, Substring$, _
Optional CompareMethod As VbCompareMethod = vbBinaryCompare)
    Dim L&
    Dim i&
    L = Len(Substring)
    If L < 1 Then
        Exit Function
    End If
    For i = 1 To Len(Text) - L + 1
        If StrComp(Mid$(Text, i, L), Substring, CompareMethod) = 0 Then
            CountSubstring = CountSubstring + 1
        End If
    Next i
End Function

Public Function SI$(Text$, ParamArray Args())
    Dim Arr1$()
    Arr1 = Split(Text, "{}")
    Dim C1&
    C1 = UBound(Arr1) - LBound(Arr1) + 1
    Dim C2&
    C2 = UBound(Args) - LBound(Args) + 1
    If C2 <> C1 - 1 Then
        Err.Raise 5
    End If
    Dim Arr2$()
    ReDim Arr2(0 To C1 * 2 - 1)
    Dim i&
    Dim j&
    Dim k&
    For i = LBound(Arr1) To UBound(Arr1)
        Arr2(k) = Arr1(i)
        k = k + 1
        If j < C2 Then
            Arr2(k) = Args(j)
            k = k + 1
            j = j + 1
        End If
    Next i
    SI = Join(Arr2, "")
End Function

Public Function StartsWith(WholeText$, SearchText$, _
Optional CompareMethod As VbCompareMethod = vbBinaryCompare) As Boolean
    If Len(SearchText) = 0 Then
        StartsWith = True
    Else
        StartsWith = InStr(1, WholeText, SearchText, CompareMethod) = 1
    End If
End Function

Public Function EndsWith(WholeText$, SearchText$, Optional _
CompareMethod As VbCompareMethod = vbBinaryCompare) As Boolean
    If Len(SearchText) = 0 Then
        EndsWith = True
    ElseIf Len(SearchText) > Len(WholeText) Then
        EndsWith = False
    Else
        EndsWith = InStrRev(WholeText, SearchText, -1, CompareMethod) = _
        Len(WholeText) - Len(SearchText) + 1
    End If
End Function

'Text Cleaning Functions===============================================
'======================================================================

Public Function ReplaceNBSP$(Text$)
    ReplaceNBSP = _
    Replace(Text, Chr(160), " ")
End Function

Public Function TrimAll$(Text$)
    TrimAll = Trim$(Text)
    Do While InStr(TrimAll, "  ") > 0
        TrimAll = _
        Replace(TrimAll, "  ", " ")
    Loop
End Function

Public Function RemoveNonPrintableCharacters$(Text$)
    'Does not remove new line characters
    Dim i&
    Dim c&
    Dim L&
    L = Len(Text)
    RemoveNonPrintableCharacters = String$(L, Chr(0))
    For i = 1 To L
        Dim CurrentCharCode&
        CurrentCharCode = AscW2(Mid$(Text, i, 1))
        If CurrentCharCode > 31 Or CurrentCharCode = 13 Or CurrentCharCode = 10 Then
            c = c + 1
            Mid$(RemoveNonPrintableCharacters, c, 1) = Mid$(Text, i, 1)
        End If
    Next i
    RemoveNonPrintableCharacters = Left$(RemoveNonPrintableCharacters, c)
End Function

Public Function ReplaceNewLineCharacters$(Text$, Optional Replacement$ = " ")
    ReplaceNewLineCharacters = Replace(Text, vbCrLf, Replacement)
    ReplaceNewLineCharacters = Replace(ReplaceNewLineCharacters, vbCr, Replacement)
    ReplaceNewLineCharacters = Replace(ReplaceNewLineCharacters, vbLf, Replacement)
End Function

Public Function CleanText$(Text$, _
Optional NonPrintable As Boolean = True, _
Optional NewLines As Boolean = True, _
Optional NonBreaking As Boolean = True, _
Optional TrimSpaces As Boolean = True, _
Optional NLReplacement$ = " ")
    CleanText = Text
    If NonPrintable Then CleanText = RemoveNonPrintableCharacters(CleanText)
    If NewLines Then CleanText = ReplaceNewLineCharacters(CleanText, NLReplacement)
    If NonBreaking Then CleanText = ReplaceNBSP(CleanText)
    If TrimSpaces Then CleanText = TrimAll(CleanText)
End Function


'Unit Tests============================================================
'======================================================================

Private Function TestmodString() As Boolean
    
    TestmodString = _
        TestAscW2 And _
        TestCountSubstring And _
        TestSI And _
        TestStartsWith And _
        TestEndsWith And _
        TestReplaceNBSP And _
        TestTrimAll And _
        TestRemoveNonPrintableCharacters And _
        TestReplaceNewLineCharacters And _
        TestCleanText
    
    Debug.Print "TestmodString: " & TestmodString
    
End Function

Private Function TestAscW2() As Boolean
    
    TestAscW2 = True
    
    'Empty
    On Error Resume Next
    AscW2 ""
    If Err.Number <> 5 Then
        TestAscW2 = False
        Debug.Print "Empty"
    End If
    On Error GoTo 0
    
    Dim i&
    
    '0 - 127
    For i = 0 To 127
        If AscW2(Chr(i)) <> i Then
            TestAscW2 = False
            Debug.Print "0 - 127"
            Exit For
        End If
    Next i
    
    '0 - 65535
    For i = 0 To 65535
        If AscW2(ChrW(i)) <> i Then
            TestAscW2 = False
            Debug.Print "0 - 65535"
            Exit For
        End If
    Next i
    
    'Multiple characters
    For i = 0 To 65535
        If AscW2(ChrW(i) & "ello, World!") <> i Then
            TestAscW2 = False
            Debug.Print "Multiple characters"
            Exit For
        End If
    Next i
    
    Debug.Print "TestAscW2: " & TestAscW2
    
End Function

Private Function TestCountSubstring() As Boolean

    TestCountSubstring = True

    'Empty
    If CountSubstring("", "A") <> 0 Then
        TestCountSubstring = False
        Debug.Print "Empty1"
    End If
    If CountSubstring("A", "") <> 0 Then
        TestCountSubstring = False
        Debug.Print "Empty2"
    End If
    If CountSubstring("", "") <> 0 Then
        TestCountSubstring = False
        Debug.Print "Empty3"
    End If
    
    'Single char

        'Not there
        If CountSubstring("Hello", "A") <> 0 Then
            TestCountSubstring = False
            Debug.Print "Single char Not there"
        End If

        'Single there
            'Beginning
            If CountSubstring("Hello", "H") <> 1 Then
                TestCountSubstring = False
                Debug.Print "Single char Single there Beginning"
            End If

            'Middle
            If CountSubstring("Hello", "e") <> 1 Then
                TestCountSubstring = False
                Debug.Print "Single char Single there Middle"
            End If

            'End
            If CountSubstring("Hello", "o") <> 1 Then
                TestCountSubstring = False
                Debug.Print "Single char Single there End"
            End If

        'Multiple there
            'Beginning
            If CountSubstring("AABBCC", "A") <> 2 Then
                TestCountSubstring = False
                Debug.Print "Single char Multiple there Beginning"
            End If

            'Middle
            If CountSubstring("AABBCC", "B") <> 2 Then
                TestCountSubstring = False
                Debug.Print "Single char Multiple there Middle"
            End If

            'End
            If CountSubstring("AABBCC", "C") <> 2 Then
                TestCountSubstring = False
                Debug.Print "Single char Multiple there End"
            End If
        
        'Compare method
            If CountSubstring("AABBCC", "c") <> 0 Then
                TestCountSubstring = False
                Debug.Print "Single char Compare method binary"
            End If
            If CountSubstring("AABBCC", "c", vbTextCompare) <> 2 Then
                TestCountSubstring = False
                Debug.Print "Single char Compare method text"
            End If
        
    'Multiple char

        'Not there
        If CountSubstring("Hello", "ZZ") <> 0 Then
            TestCountSubstring = False
            Debug.Print "Multiple char Not there"
        End If

        'Single there
            'Beginning
            If CountSubstring("Hello", "He") <> 1 Then
                TestCountSubstring = False
                Debug.Print "Multiple char Single there Beginning"
            End If

            'Middle
            If CountSubstring("Hello", "el") <> 1 Then
                TestCountSubstring = False
                Debug.Print "Multiple char Single there Middle"
            End If

            'End
            If CountSubstring("Hello", "lo") <> 1 Then
                TestCountSubstring = False
                Debug.Print "Multiple char Single there End"
            End If

        'Multiple there
            'Beginning
            If CountSubstring("AAAABBBBCCCC", "AA") <> 3 Then
                TestCountSubstring = False
                Debug.Print "Multiple char Multiple there Beginning"
            End If

            'Middle
            If CountSubstring("AAAABBBBCCCC", "BB") <> 3 Then
                TestCountSubstring = False
                Debug.Print "Multiple char Multiple there Middle"
            End If
            
            'End
            If CountSubstring("AAAABBBBCCCC", "CC") <> 3 Then
                TestCountSubstring = False
                Debug.Print "Multiple char Multiple there End"
            End If
            
        'Compare method
        If CountSubstring("AAAABBBBCCCC", "cc") <> 0 Then
            TestCountSubstring = False
            Debug.Print "Multiple char Compare method binary"
        End If
        If CountSubstring("AAAABBBBCCCC", "cc", vbTextCompare) <> 3 Then
            TestCountSubstring = False
            Debug.Print "Multiple char Compare method text"
        End If
            
    Debug.Print "TestCountSubstring: " & TestCountSubstring

End Function

Private Function TestSI() As Boolean
    
    TestSI = True
    
    'Empty
    On Error Resume Next
    SI ""
    If Err.Number <> 5 Then
        TestSI = False
        Debug.Print "Empty"
    End If
    On Error GoTo 0
    
    'None
    If SI("Hello") <> "Hello" Then
        TestSI = False
        Debug.Print "None"
    End If
        
    'Single
        'Beginning
        If SI("{} World", "Hello") <> "Hello World" Then
            TestSI = False
            Debug.Print "Single Beginning"
        End If
        'Middle
        If SI("He{}World", "llo ") <> "Hello World" Then
            TestSI = False
            Debug.Print "Single Middle"
        End If
        'End
        If SI("Hello {}", "World") <> "Hello World" Then
            TestSI = False
            Debug.Print "Single End"
        End If

    'Multiple
        'Beginning
        If SI("{}{} World", "Hello", "Hello") <> "HelloHello World" Then
            TestSI = False
            Debug.Print "Multiple Beginning"
        End If
        'Middle
        If SI("Hello{}{} World", "Hello", "Hello") <> "HelloHelloHello World" Then
            TestSI = False
            Debug.Print "Multiple Middle"
        End If
        'End
        If SI("Hello {}{}", "World", "World") <> "Hello WorldWorld" Then
            TestSI = False
            Debug.Print "Multiple End"
        End If
        
    'Wrong number of args
        'Less
        On Error Resume Next
        SI "Hello {} {}", "World"
        If Err.Number <> 5 Then
            TestSI = False
            Debug.Print "Wrong number of args Less"
        End If
        On Error GoTo 0
        'More
        On Error Resume Next
        SI "Hello {} {}", "World", "World", "World"
        If Err.Number <> 5 Then
            TestSI = False
            Debug.Print "Wrong number of args More"
        End If
        On Error GoTo 0
        'None1
        On Error Resume Next
        SI "Hello", "World"
        If Err.Number <> 5 Then
            TestSI = False
            Debug.Print "Wrong number of args None1"
        End If
        On Error GoTo 0
        'None2
        On Error Resume Next
        SI "Hello {}"
        If Err.Number <> 5 Then
            TestSI = False
            Debug.Print "Wrong number of args None2"
        End If
        On Error GoTo 0
        
    Debug.Print "TestSI: " & TestSI

End Function

Public Function TestStartsWith() As Boolean

    TestStartsWith = True

    If StartsWith("", "Test") <> False Then
        TestStartsWith = False
        Debug.Print "Blank Whole"
    End If

    If StartsWith("Test", "") <> True Then
        TestStartsWith = False
        Debug.Print "Blank Search"
    End If

    If StartsWith("", "") <> True Then
        TestStartsWith = False
        Debug.Print "Blank Both"
    End If

    If StartsWith("A", "A") <> True Then
        TestStartsWith = False
        Debug.Print "Single letter True"
    End If

    If StartsWith("A", "B") <> False Then
        TestStartsWith = False
        Debug.Print "Single letter False"
    End If

    If StartsWith("ABC", "A") <> True Then
        TestStartsWith = False
        Debug.Print "Multiple letter A"
    End If

    If StartsWith("ABC", "AB") <> True Then
        TestStartsWith = False
        Debug.Print "Multiple letter AB"
    End If

    If StartsWith("ABC", "ABC") <> True Then
        TestStartsWith = False
        Debug.Print "Multiple letter ABC"
    End If

    If StartsWith("ABC", "ABCD") <> False Then
        TestStartsWith = False
        Debug.Print "Search longer than Whole"
    End If

    If StartsWith("ABC", "abc") <> False Then
        TestStartsWith = False
        Debug.Print "Case Sensitivity False"
    End If

    If StartsWith("ABC", "abc", vbTextCompare) <> True Then
        TestStartsWith = False
        Debug.Print "Case Sensitivity True"
    End If

    If StartsWith("ABCDAB", "AB") <> True Then
        TestStartsWith = False
        Debug.Print "Repeated"
    End If

    Debug.Print "TestStartsWith: " & TestStartsWith

End Function

Public Function TestEndsWith() As Boolean

    TestEndsWith = True

    If EndsWith("", "Test") <> False Then
        TestEndsWith = False
        Debug.Print "Blank Whole"
    End If

    If EndsWith("Test", "") <> True Then
        TestEndsWith = False
        Debug.Print "Blank Search"
    End If

    If EndsWith("", "") <> True Then
        TestEndsWith = False
        Debug.Print "Blank Both"
    End If

    If EndsWith("A", "A") <> True Then
        TestEndsWith = False
        Debug.Print "Single letter True"
    End If

    If EndsWith("A", "B") <> False Then
        TestEndsWith = False
        Debug.Print "Single letter False"
    End If

    If EndsWith("ABC", "C") <> True Then
        TestEndsWith = False
        Debug.Print "Multiple letter A"
    End If

    If EndsWith("ABC", "BC") <> True Then
        TestEndsWith = False
        Debug.Print "Multiple letter AB"
    End If

    If EndsWith("ABC", "ABC") <> True Then
        TestEndsWith = False
        Debug.Print "Multiple letter ABC"
    End If

    If EndsWith("ABC", "ABCD") <> False Then
        TestEndsWith = False
        Debug.Print "Search longer than Whole"
    End If

    If EndsWith("ABC", "abc") <> False Then
        TestEndsWith = False
        Debug.Print "Case Sensitivity False"
    End If

    If EndsWith("ABC", "abc", vbTextCompare) <> True Then
        TestEndsWith = False
        Debug.Print "Case Sensitivity True"
    End If

    If EndsWith("ABCDAB", "AB") <> True Then
        TestEndsWith = False
        Debug.Print "Repeated"
    End If

    Debug.Print "TestEndsWith: " & TestEndsWith

End Function

Public Function TestReplaceNBSP() As Boolean
    
    TestReplaceNBSP = True
    
    'Empty
    If ReplaceNBSP("") <> "" Then
        TestReplaceNBSP = False
        Debug.Print "Empty"
    End If
    
    'Not There
    If ReplaceNBSP("A") <> "A" Then
        TestReplaceNBSP = False
        Debug.Print "Not There"
    End If
        
    'Single
        'Beginning
        If ReplaceNBSP(Chr(160) & "A") <> " A" Then
            TestReplaceNBSP = False
            Debug.Print "Single Beginning"
        End If
        'Middle
        If ReplaceNBSP("A" & Chr(160) & "A") <> "A A" Then
            TestReplaceNBSP = False
            Debug.Print "Single Middle"
        End If
        'End
        If ReplaceNBSP("A" & Chr(160)) <> "A " Then
            TestReplaceNBSP = False
            Debug.Print "Single End"
        End If
        'Beginning Middle End
        If ReplaceNBSP(Chr(160) & "A" & Chr(160) & "A" & Chr(160)) <> " A A " Then
            TestReplaceNBSP = False
            Debug.Print "Single Beginning Middle End"
        End If
        
    'Multiple
        'Beginning
        If ReplaceNBSP(Chr(160) & Chr(160) & "A") <> "  A" Then
            TestReplaceNBSP = False
            Debug.Print "Multiple Beginning"
        End If
        'Middle
        If ReplaceNBSP("A" & Chr(160) & Chr(160) & "A") <> "A  A" Then
            TestReplaceNBSP = False
            Debug.Print "Multiple Middle"
        End If
        'End
        If ReplaceNBSP("A" & Chr(160) & Chr(160)) <> "A  " Then
            TestReplaceNBSP = False
            Debug.Print "Multiple End"
        End If
        'Beginning Middle End
        If ReplaceNBSP(Chr(160) & Chr(160) & "A" & Chr(160) & Chr(160) & "A" & _
        Chr(160) & Chr(160)) <> "  A  A  " Then
            TestReplaceNBSP = False
            Debug.Print "Multiple Beginning Middle End"
        End If
        
    Debug.Print "TestReplaceNBSP: " & TestReplaceNBSP
    
End Function

Private Function TestTrimAll() As Boolean
    
    TestTrimAll = True
    
    'Empty
    If TrimAll("") <> "" Then
        TestTrimAll = False
        Debug.Print "Empty"
    End If
    
    'One only
    If TrimAll(" ") <> "" Then
        TestTrimAll = False
        Debug.Print "One only"
    End If
    
    'Multiple only
    If TrimAll("  ") <> "" Then
        TestTrimAll = False
        Debug.Print "Multiple only"
    End If
    
    'No spaces
    If TrimAll("HelloWorld") <> "HelloWorld" Then
        TestTrimAll = False
        Debug.Print "No spaces"
    End If
    
    'Single
        'Leading
        If TrimAll(" HelloWorld") <> "HelloWorld" Then
            TestTrimAll = False
            Debug.Print "Single Leading"
        End If
        'Middle
        If TrimAll("Hello World") <> "Hello World" Then
            TestTrimAll = False
            Debug.Print "Single Middle"
        End If
        'Trailing
        If TrimAll("HelloWorld ") <> "HelloWorld" Then
            TestTrimAll = False
            Debug.Print "Single Trailing"
        End If
        'Leading Middle Trailing
        If TrimAll(" Hello World ") <> "Hello World" Then
            TestTrimAll = False
            Debug.Print "Single Leading Middle Trailing"
        End If
    
    'Multiple
        'Leading
        If TrimAll("  HelloWorld") <> "HelloWorld" Then
            TestTrimAll = False
            Debug.Print "Multiple Leading"
        End If
        'Middle
        If TrimAll("Hello  World") <> "Hello World" Then
            TestTrimAll = False
            Debug.Print "Multiple Middle"
        End If
        'Trailing
        If TrimAll("HelloWorld  ") <> "HelloWorld" Then
            TestTrimAll = False
            Debug.Print "Multiple Trailing"
        End If
        'Leading Middle Trailing
        If TrimAll(" Hello  World ") <> "Hello World" Then
            TestTrimAll = False
            Debug.Print "Multiple Leading Middle Trailing"
        End If
    
    Debug.Print "TestTrimAll: " & TestTrimAll
    
End Function

Public Function TestRemoveNonPrintableCharacters() As Boolean
    
    TestRemoveNonPrintableCharacters = True
    
    'Empty
    If RemoveNonPrintableCharacters("") <> "" Then
        TestRemoveNonPrintableCharacters = False
        Debug.Print "Empty"
    End If
    
    'None
    If RemoveNonPrintableCharacters("A") <> "A" Then
        TestRemoveNonPrintableCharacters = False
        Debug.Print "None"
    End If
    
    'Only
    If RemoveNonPrintableCharacters(Chr(0)) <> "" Then
        TestRemoveNonPrintableCharacters = False
        Debug.Print "Only"
    End If
    
    'Multiple none
    If RemoveNonPrintableCharacters("AAA") <> "AAA" Then
        TestRemoveNonPrintableCharacters = False
        Debug.Print "Multiple none"
    End If
    
    'Multiple only
    If RemoveNonPrintableCharacters(Chr(0) & Chr(0) & Chr(0)) <> "" Then
        TestRemoveNonPrintableCharacters = False
        Debug.Print "Multiple only"
    End If
    
    'Beginning
    If RemoveNonPrintableCharacters(Chr(0) & "AA") <> "AA" Then
        TestRemoveNonPrintableCharacters = False
        Debug.Print "Beginning"
    End If
    
    'Middle
    If RemoveNonPrintableCharacters("A" & Chr(0) & "A") <> "AA" Then
        TestRemoveNonPrintableCharacters = False
        Debug.Print "Middle"
    End If
    
    'End
    If RemoveNonPrintableCharacters("AA" & Chr(0)) <> "AA" Then
        TestRemoveNonPrintableCharacters = False
        Debug.Print "End"
    End If
    
    'Beginning Middle End
    If RemoveNonPrintableCharacters(Chr(0) & "A" & Chr(0) & "A" & Chr(0)) <> "AA" Then
        TestRemoveNonPrintableCharacters = False
        Debug.Print "Beginning Middle End"
    End If
    
    'Does not remove new lines
    If RemoveNonPrintableCharacters(vbCrLf & " " & vbCr & " " & vbLf) <> _
    vbCrLf & " " & vbCr & " " & vbLf Then
        TestRemoveNonPrintableCharacters = False
        Debug.Print "Does not remove new lines"
    End If
    
    '0 - 31
    Dim i&
    For i = 0 To 31
        If i <> 10 And i <> 13 Then
            If RemoveNonPrintableCharacters(Chr(i)) <> "" Then
                TestRemoveNonPrintableCharacters = False
                Debug.Print "0 - 31"
            End If
        End If
    Next i
    
    '32 - 65535
    For i = 32 To 65535
        If RemoveNonPrintableCharacters(ChrW(i)) <> ChrW(i) Then
            TestRemoveNonPrintableCharacters = False
            Debug.Print "32 - 65535"
        End If
    Next i
    
    Debug.Print "TestRemoveNonPrintableCharacters: " & TestRemoveNonPrintableCharacters
    
End Function

Public Function TestReplaceNewLineCharacters() As Boolean

    TestReplaceNewLineCharacters = True
    
    'Empty
    If ReplaceNewLineCharacters("") <> "" Then
        TestReplaceNewLineCharacters = False
        Debug.Print "Empty"
    End If
    
    'CRLF
        'Single
            'Only
            If ReplaceNewLineCharacters(vbCrLf) <> " " Then
                TestReplaceNewLineCharacters = False
                Debug.Print "CRLF Single Only"
            End If
            'Leading
            If ReplaceNewLineCharacters(vbCrLf & "A") <> " A" Then
                TestReplaceNewLineCharacters = False
                Debug.Print "CRLF Single Leading"
            End If
            'Trailing
            If ReplaceNewLineCharacters("A" & vbCrLf) <> "A " Then
                TestReplaceNewLineCharacters = False
                Debug.Print "CRLF Single Trailing"
            End If
            'Middle
            If ReplaceNewLineCharacters("A" & vbCrLf & "A") <> "A A" Then
                TestReplaceNewLineCharacters = False
                Debug.Print "CRLF Single Middle"
            End If
        'Multiple
            'Only
            If ReplaceNewLineCharacters(vbCrLf & vbCrLf) <> "  " Then
                TestReplaceNewLineCharacters = False
                Debug.Print "CRLF Multiple Only"
            End If
            'Leading
            If ReplaceNewLineCharacters(vbCrLf & vbCrLf & "A") <> "  A" Then
                TestReplaceNewLineCharacters = False
                Debug.Print "CRLF Multiple Leading"
            End If
            'Trailing
            If ReplaceNewLineCharacters("A" & vbCrLf & vbCrLf) <> "A  " Then
                TestReplaceNewLineCharacters = False
                Debug.Print "CRLF Multiple Trailing"
            End If
            'Middle
            If ReplaceNewLineCharacters("A" & vbCrLf & vbCrLf & "A") <> "A  A" Then
                TestReplaceNewLineCharacters = False
                Debug.Print "CRLF Multiple Middle"
            End If
    'CR
        'Single
            'Only
            If ReplaceNewLineCharacters(vbCr) <> " " Then
                TestReplaceNewLineCharacters = False
                Debug.Print "CR Single Only"
            End If
            'Leading
            If ReplaceNewLineCharacters(vbCr & "A") <> " A" Then
                TestReplaceNewLineCharacters = False
                Debug.Print "CR Single Leading"
            End If
            'Trailing
            If ReplaceNewLineCharacters("A" & vbCr) <> "A " Then
                TestReplaceNewLineCharacters = False
                Debug.Print "CR Single Trailing"
            End If
            'Middle
            If ReplaceNewLineCharacters("A" & vbCr & "A") <> "A A" Then
                TestReplaceNewLineCharacters = False
                Debug.Print "CR Single Middle"
            End If
        'Multiple
            'Only
            If ReplaceNewLineCharacters(vbCr & vbCr) <> "  " Then
                TestReplaceNewLineCharacters = False
                Debug.Print "CR Multiple Only"
            End If
            'Leading
            If ReplaceNewLineCharacters(vbCr & vbCr & "A") <> "  A" Then
                TestReplaceNewLineCharacters = False
                Debug.Print "CR Multiple Leading"
            End If
            'Trailing
            If ReplaceNewLineCharacters("A" & vbCr & vbCr) <> "A  " Then
                TestReplaceNewLineCharacters = False
                Debug.Print "CR Multiple Trailing"
            End If
            'Middle
            If ReplaceNewLineCharacters("A" & vbCr & vbCr & "A") <> "A  A" Then
                TestReplaceNewLineCharacters = False
                Debug.Print "CR Multiple Middle"
            End If
    'LF
        'Single
            'Only
            If ReplaceNewLineCharacters(vbLf) <> " " Then
                TestReplaceNewLineCharacters = False
                Debug.Print "LF Single Only"
            End If
            'Leading
            If ReplaceNewLineCharacters(vbLf & "A") <> " A" Then
                TestReplaceNewLineCharacters = False
                Debug.Print "CR Single Leading"
            End If
            'Trailing
            If ReplaceNewLineCharacters("A" & vbLf) <> "A " Then
                TestReplaceNewLineCharacters = False
                Debug.Print "LF Single Trailing"
            End If
            'Middle
            If ReplaceNewLineCharacters("A" & vbLf & "A") <> "A A" Then
                TestReplaceNewLineCharacters = False
                Debug.Print "LF Single Middle"
            End If
        'Multiple
            'Only
            If ReplaceNewLineCharacters(vbLf & vbLf) <> "  " Then
                TestReplaceNewLineCharacters = False
                Debug.Print "LF Multiple Only"
            End If
            'Leading
            If ReplaceNewLineCharacters(vbLf & vbLf & "A") <> "  A" Then
                TestReplaceNewLineCharacters = False
                Debug.Print "LF Multiple Leading"
            End If
            'Trailing
            If ReplaceNewLineCharacters("A" & vbLf & vbLf) <> "A  " Then
                TestReplaceNewLineCharacters = False
                Debug.Print "LF Multiple Trailing"
            End If
            'Middle
            If ReplaceNewLineCharacters("A" & vbLf & vbLf & "A") <> "A  A" Then
                TestReplaceNewLineCharacters = False
                Debug.Print "LF Multiple Middle"
            End If
        
    'CRLF CR LF
    If ReplaceNewLineCharacters("A" & vbCrLf & "A" & vbCr & "A" & vbLf) <> "A A A " Then
        TestReplaceNewLineCharacters = False
        Debug.Print "CRLF CR LF"
    End If
    
    'Replacement single char
    If ReplaceNewLineCharacters("A" & vbCrLf & "A" & vbCr & "A" & vbLf, "B") <> "ABABAB" Then
        TestReplaceNewLineCharacters = False
        Debug.Print "Replacement single char"
    End If
    
    'Replacement multiple chars
    If ReplaceNewLineCharacters("A" & vbCrLf & "A" & vbCr & "A" & vbLf, "CC") <> "ACCACCACC" Then
        TestReplaceNewLineCharacters = False
        Debug.Print "Replacement multiple chars"
    End If
    
    'Replacement empty
    If ReplaceNewLineCharacters("A" & vbCrLf & "A" & vbCr & "A" & vbLf, "") <> "AAA" Then
        TestReplaceNewLineCharacters = False
        Debug.Print "Replacement empty"
    End If
    
    Debug.Print "TestReplaceNewLineCharacters: " & TestReplaceNewLineCharacters
    
End Function

Public Function TestCleanText() As Boolean

    TestCleanText = True

    'Empty
    If CleanText("") <> "" Then
        TestCleanText = False
        Debug.Print "Empty"
    End If

    'All
    If CleanText(" Hello  " & vbCrLf & "World" & Chr(160) & Chr(0)) <> "Hello World" Then
        TestCleanText = False
        Debug.Print "All"
    End If

    'Disable Non Printable
    If CleanText(" Hello  " & vbCrLf & "World" & Chr(160) & Chr(0), False) <> "Hello World " & Chr(0) Then
        TestCleanText = False
        Debug.Print "Disable Non Printable"
    End If

    'Disable New Lines
    If CleanText(" Hello  " & vbCrLf & "World" & Chr(160) & Chr(0), , False) <> "Hello " & vbCrLf & "World" Then
        TestCleanText = False
        Debug.Print "Disable New Lines"
    End If

    'Disable Non Breaking
    If CleanText(" Hello  " & vbCrLf & "World" & Chr(160) & Chr(0), , , False) <> "Hello World" & Chr(160) Then
        TestCleanText = False
        Debug.Print "Disable Non Breaking"
    End If

    'Disable Trim Spaces
    If CleanText(" Hello  " & vbCrLf & "World" & Chr(160) & Chr(0), , , , False) <> " Hello   World " Then
        TestCleanText = False
        Debug.Print "Disable Trim Spaces"
    End If
    
    'NL Replacement
    If CleanText(" Hello  " & vbCrLf & "World" & Chr(160) & Chr(0), , , , , "|") <> "Hello |World" Then
        TestCleanText = False
        Debug.Print "NL Replacement"
    End If

    Debug.Print "TestCleanText: " & TestCleanText

End Function
