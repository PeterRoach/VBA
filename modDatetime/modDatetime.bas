Attribute VB_Name = "modDateTime"
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
'  Module Name: modDatetime
'  Module Description: Contains functions for working with dates and times.
'  Module Version: 1.0
'  Module License: MIT
'  Module Author: Peter Roach; PeterRoach.Code@gmail.com
'  Module Contents:
'  ----------------------------------------
'    Public Procedures:
'        WeekdayCount
'        WorkdayCount
'        GetUSBankHolidays
'    Test Procedures:
'        TestmodDatetime
'        TestWeekdayCount
'        TestWorkdayCount
'        TestGetUSBankHolidays
'  ----------------------------------------

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'  Example Usage:

Private Sub Example()

    Debug.Print WeekdayCount(#1/1/2023#, #12/31/2023#)

    Dim Holidays() As Date
    Holidays = GetUSBankHolidays(2023)

    Debug.Print WorkdayCount(#1/1/2023#, #12/31/2023#, Holidays)

End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


'Functions=============================================================
'======================================================================

Public Function WeekdayCount&(StartDate As Date, EndDate As Date)

    If StartDate > EndDate Then
        WeekdayCount = -WeekdayCount(EndDate, StartDate)
        Exit Function
    End If

    Dim TotalDays&: TotalDays = DateDiff("d", StartDate, EndDate) + 1

    Dim WeekendDays&: WeekendDays = DateDiff("ww", StartDate, EndDate) * 2

    WeekdayCount = TotalDays - WeekendDays

    Dim StartDay&: StartDay = Weekday(StartDate)

    Dim EndDay&: EndDay = Weekday(EndDate)

    If StartDay = vbSunday Then WeekdayCount = WeekdayCount - 1

    If EndDay = vbSaturday Then WeekdayCount = WeekdayCount - 1

End Function

Public Function WorkdayCount&(StartDate As Date, EndDate As Date, Holidays() As Date)

    If StartDate > EndDate Then
        WorkdayCount = -WorkdayCount(EndDate, StartDate, Holidays)
        Exit Function
    End If

    Dim HolidayCount&

    Dim HolidayWeekday&

    Dim DuplicateFound As Boolean

    Dim i&

    Dim j&

    WorkdayCount = WeekdayCount(StartDate, EndDate)

    For i = LBound(Holidays) To UBound(Holidays)

        If Holidays(i) >= StartDate And Holidays(i) <= EndDate Then

            HolidayWeekday = Weekday(Holidays(i))

            If HolidayWeekday > 1 And HolidayWeekday < 7 Then

                DuplicateFound = False

                For j = i + 1 To UBound(Holidays)

                    If Holidays(j) = Holidays(i) Then

                        DuplicateFound = True

                        Exit For

                    End If

                Next j

                If Not DuplicateFound Then

                    HolidayCount = HolidayCount + 1

                End If

            End If

        End If

    Next i

    WorkdayCount = WorkdayCount - HolidayCount

End Function

Public Function GetUSBankHolidays(Year&) As Date()

    Dim Holidays(0 To 10) As Date

    Dim i&

    'New Year's Day (January 1st)
    Holidays(0) = DateSerial(Year, 1, 1)

    'Martin Luther King Jr. Day (Third Monday in January)
    For i = 15 To 21
        If Weekday(DateSerial(Year, 1, i)) = 2 Then
            Holidays(1) = DateSerial(Year, 1, i)
            Exit For
        End If
    Next i

    'Washington's Birthday / Presidents Day (Third Monday in February)
    For i = 15 To 21
        If Weekday(DateSerial(Year, 2, i)) = 2 Then
            Holidays(2) = DateSerial(Year, 2, i)
            Exit For
        End If
    Next i

    'Memorial Day (Last Monday in May)
    For i = 30 To 24 Step -1
        If Weekday(DateSerial(Year, 5, i)) = 2 Then
            Holidays(3) = DateSerial(Year, 5, i)
            Exit For
        End If
    Next i

    'Juneteenth National Independence Day (June 19th)
    Holidays(4) = DateSerial(Year, 6, 19)

    'Independence Day (July 4th)
    Holidays(5) = DateSerial(Year, 7, 4)

    'Labor Day (First Monday in September)
    For i = 1 To 7
        If Weekday(DateSerial(Year, 9, i)) = 2 Then
            Holidays(6) = DateSerial(Year, 9, i)
            Exit For
        End If
    Next i

    'Columbus Day (Second Monday in October)
    For i = 8 To 14
        If Weekday(DateSerial(Year, 10, i)) = 2 Then
            Holidays(7) = DateSerial(Year, 10, i)
            Exit For
        End If
    Next i

    'Veterans Day (November 11th)
    Holidays(8) = DateSerial(Year, 11, 11)

    'Thanksgiving Day (Fourth Thursday in November)
    For i = 22 To 28
        If Weekday(DateSerial(Year, 11, i)) = 5 Then
            Holidays(9) = DateSerial(Year, 11, i)
            Exit For
        End If
    Next i

    'Christmas Day
    Holidays(10) = DateSerial(Year, 12, 25)

    For i = LBound(Holidays) To UBound(Holidays)

        Select Case Weekday(Holidays(i))

            Case 1: Holidays(i) = Holidays(i) + 1

            Case 7: Holidays(i) = Holidays(i) - 1

        End Select

    Next i

    GetUSBankHolidays = Holidays

End Function

'Test==================================================================
'======================================================================

Private Function TestmodDatetime() As Boolean

    TestmodDatetime = _
        TestWeekdayCount And _
        TestWorkdayCount And _
        TestGetUSBankHolidays

    Debug.Print "TestmodDatetime: " & TestmodDatetime

End Function

Public Function TestWeekdayCount()

    TestWeekdayCount = True

    If WeekdayCount(#9/2/2023#, #9/2/2023#) <> 0 Then
        TestWeekdayCount = False
        Debug.Print "Start and End on same day (weekend)"
    End If
    If WeekdayCount(#9/4/2023#, #9/4/2023#) <> 1 Then
        TestWeekdayCount = False
        Debug.Print "Start and End on same day (weekday)"
    End If

    If WeekdayCount(#9/2/2023#, #9/3/2023#) <> 0 Then
        TestWeekdayCount = False
        Debug.Print "Start and End on Weekend"
    End If

    If WeekdayCount(#9/4/2023#, #9/8/2023#) <> 5 Then
        TestWeekdayCount = False
        Debug.Print "Start and End during week"
    End If

    If WeekdayCount(#9/1/2023#, #9/30/2023#) <> 21 Then
        TestWeekdayCount = False
        Debug.Print "Several Weeks"
    End If

    If WeekdayCount(#9/3/2023#, #9/2/2023#) <> 0 Then
        TestWeekdayCount = False
        Debug.Print "Start and End on Weekend (Start date > End date)"
    End If

    If WeekdayCount(#9/8/2023#, #9/4/2023#) <> -5 Then
        TestWeekdayCount = False
        Debug.Print "Start and End during Weekday (Start date > End date)"
    End If

    If WeekdayCount(#9/30/2023#, #9/1/2023#) <> -21 Then
        TestWeekdayCount = False
        Debug.Print "Several Weeks (Start date > End date)"
    End If

    Debug.Print "TestWeekdayCount: " & TestWeekdayCount

End Function

Public Function TestWorkdayCount()

    TestWorkdayCount = True

    Dim Holidays() As Date

    ReDim Holidays(0 To 1)
    Holidays(0) = #9/2/2023#
    Holidays(1) = #9/3/2023#
    If WorkdayCount(#9/1/2023#, #9/4/2023#, Holidays) <> 2 Then
        TestWorkdayCount = False
        Debug.Print "Weekend holidays"
    End If

    ReDim Holidays(0 To 1)
    Holidays(0) = #9/4/2023#
    Holidays(1) = #9/6/2023#
    If WorkdayCount(#9/2/2023#, #9/10/2023#, Holidays) <> 3 Then
        TestWorkdayCount = False
        Debug.Print "Weekday holidays"
    End If

    ReDim Holidays(0 To 1)
    Holidays(0) = #9/4/2023#
    Holidays(1) = #9/5/2023#
    If WorkdayCount(#9/2/2023#, #9/10/2023#, Holidays) <> 3 Then
        TestWorkdayCount = False
        Debug.Print "Consecutive holidays"
    End If

    'Multipe Holidays on same day should only be removed once
    ReDim Holidays(0 To 1)
    Holidays(0) = #9/4/2023#
    Holidays(1) = #9/4/2023#
    If WorkdayCount(#9/2/2023#, #9/10/2023#, Holidays) <> 4 Then
        TestWorkdayCount = False
        Debug.Print "Same day multiple"
    End If

    Dim USHolidays() As Date
    USHolidays = GetUSBankHolidays(2023)
    If WorkdayCount(#1/1/2023#, #12/31/2023#, USHolidays) <> 249 Then
        TestWorkdayCount = False
        Debug.Print "Whole year 249"
    End If

    USHolidays = GetUSBankHolidays(2022)
    If WorkdayCount(#1/1/2022#, #12/31/2022#, USHolidays) <> 250 Then
        TestWorkdayCount = False
        Debug.Print "Whole year 250"
    End If

    Debug.Print "TestWorkdayCount: " & TestWorkdayCount

End Function

Private Function TestGetUSBankHolidays()

    TestGetUSBankHolidays = True

    Dim Holidays() As Date

    '2023

    Holidays = GetUSBankHolidays(2023)

    If Holidays(0) <> #1/2/2023# Then
        TestGetUSBankHolidays = False
        Debug.Print "New Year's Day (January 1st)"
    End If

    If Holidays(1) <> #1/16/2023# Then
        TestGetUSBankHolidays = False
        Debug.Print "Martin Luther King Jr. Day (Third Monday in January)"
    End If

    If Holidays(2) <> #2/20/2023# Then
        TestGetUSBankHolidays = False
        Debug.Print "Washington's Birthday / Presidents Day (Third Monday in February)"
    End If

    If Holidays(3) <> #5/29/2023# Then
        TestGetUSBankHolidays = False
        Debug.Print "Memorial Day (Last Monday in May)"
    End If

    If Holidays(4) <> #6/19/2023# Then
        TestGetUSBankHolidays = False
        Debug.Print "Juneteenth National Independence Day (June 19th)"
    End If

    If Holidays(5) <> #7/4/2023# Then
        TestGetUSBankHolidays = False
        Debug.Print "Independence Day (July 4th)"
    End If

    If Holidays(6) <> #9/4/2023# Then
        TestGetUSBankHolidays = False
        Debug.Print "Labor Day (First Monday in September)"
    End If

    If Holidays(7) <> #10/9/2023# Then
        TestGetUSBankHolidays = False
        Debug.Print "Columbus Day (Second Monday in October)"
    End If

    If Holidays(8) <> #11/10/2023# Then
        TestGetUSBankHolidays = False
        Debug.Print "Veterans Day (November 11th)"
    End If

    If Holidays(9) <> #11/23/2023# Then
        TestGetUSBankHolidays = False
        Debug.Print "Thanksgiving Day (Fourth Thursday in November)"
    End If

    If Holidays(10) <> #12/25/2023# Then
        TestGetUSBankHolidays = False
        Debug.Print "Christmas Day"
    End If

    '2024

    Holidays = GetUSBankHolidays(2024)

    If Holidays(0) <> #1/1/2024# Then
        TestGetUSBankHolidays = False
        Debug.Print "New Year's Day (January 1st)"
    End If

    If Holidays(1) <> #1/15/2024# Then
        TestGetUSBankHolidays = False
        Debug.Print "Martin Luther King Jr. Day (Third Monday in January)"
    End If

    If Holidays(2) <> #2/19/2024# Then
        TestGetUSBankHolidays = False
        Debug.Print "Washington's Birthday / Presidents Day (Third Monday in February)"
    End If

    If Holidays(3) <> #5/27/2024# Then
        TestGetUSBankHolidays = False
        Debug.Print "Memorial Day (Last Monday in May)"
    End If

    If Holidays(4) <> #6/19/2024# Then
        TestGetUSBankHolidays = False
        Debug.Print "Juneteenth National Independence Day (June 19th)"
    End If

    If Holidays(5) <> #7/4/2024# Then
        TestGetUSBankHolidays = False
        Debug.Print "Independence Day (July 4th)"
    End If

    If Holidays(6) <> #9/2/2024# Then
        TestGetUSBankHolidays = False
        Debug.Print "Labor Day (First Monday in September)"
    End If

    If Holidays(7) <> #10/14/2024# Then
        TestGetUSBankHolidays = False
        Debug.Print "Columbus Day (Second Monday in October)"
    End If

    If Holidays(8) <> #11/11/2024# Then
        TestGetUSBankHolidays = False
        Debug.Print "Veterans Day (November 11th)"
    End If

    If Holidays(9) <> #11/28/2024# Then
        TestGetUSBankHolidays = False
        Debug.Print "Thanksgiving Day (Fourth Thursday in November)"
    End If

    If Holidays(10) <> #12/25/2024# Then
        TestGetUSBankHolidays = False
        Debug.Print "Christmas Day"
    End If

    '2025

    Holidays = GetUSBankHolidays(2025)

    If Holidays(0) <> #1/1/2025# Then
        TestGetUSBankHolidays = False
        Debug.Print "New Year's Day (January 1st)"
    End If

    If Holidays(1) <> #1/20/2025# Then
        TestGetUSBankHolidays = False
        Debug.Print "Martin Luther King Jr. Day (Third Monday in January)"
    End If

    If Holidays(2) <> #2/17/2025# Then
        TestGetUSBankHolidays = False
        Debug.Print "Washington's Birthday / Presidents Day (Third Monday in February)"
    End If

    If Holidays(3) <> #5/26/2025# Then
        TestGetUSBankHolidays = False
        Debug.Print "Memorial Day (Last Monday in May)"
    End If

    If Holidays(4) <> #6/19/2025# Then
        TestGetUSBankHolidays = False
        Debug.Print "Juneteenth National Independence Day (June 19th)"
    End If

    If Holidays(5) <> #7/4/2025# Then
        TestGetUSBankHolidays = False
        Debug.Print "Independence Day (July 4th)"
    End If

    If Holidays(6) <> #9/1/2025# Then
        TestGetUSBankHolidays = False
        Debug.Print "Labor Day (First Monday in September)"
    End If

    If Holidays(7) <> #10/13/2025# Then
        TestGetUSBankHolidays = False
        Debug.Print "Columbus Day (Second Monday in October)"
    End If

    If Holidays(8) <> #11/11/2025# Then
        TestGetUSBankHolidays = False
        Debug.Print "Veterans Day (November 11th)"
    End If

    If Holidays(9) <> #11/27/2025# Then
        TestGetUSBankHolidays = False
        Debug.Print "Thanksgiving Day (Fourth Thursday in November)"
    End If

    If Holidays(10) <> #12/25/2025# Then
        TestGetUSBankHolidays = False
        Debug.Print "Christmas Day"
    End If

    '2026

    Holidays = GetUSBankHolidays(2026)

    If Holidays(0) <> #1/1/2026# Then
        TestGetUSBankHolidays = False
        Debug.Print "New Year's Day (January 1st)"
    End If

    If Holidays(1) <> #1/19/2026# Then
        TestGetUSBankHolidays = False
        Debug.Print "Martin Luther King Jr. Day (Third Monday in January)"
    End If

    If Holidays(2) <> #2/16/2026# Then
        TestGetUSBankHolidays = False
        Debug.Print "Washington's Birthday / Presidents Day (Third Monday in February)"
    End If

    If Holidays(3) <> #5/25/2026# Then
        TestGetUSBankHolidays = False
        Debug.Print "Memorial Day (Last Monday in May)"
    End If

    If Holidays(4) <> #6/19/2026# Then
        TestGetUSBankHolidays = False
        Debug.Print "Juneteenth National Independence Day (June 19th)"
    End If

    If Holidays(5) <> #7/3/2026# Then
        TestGetUSBankHolidays = False
        Debug.Print "Independence Day (July 4th)"
    End If

    If Holidays(6) <> #9/7/2026# Then
        TestGetUSBankHolidays = False
        Debug.Print "Labor Day (First Monday in September)"
    End If

    If Holidays(7) <> #10/12/2026# Then
        TestGetUSBankHolidays = False
        Debug.Print "Columbus Day (Second Monday in October)"
    End If

    If Holidays(8) <> #11/11/2026# Then
        TestGetUSBankHolidays = False
        Debug.Print "Veterans Day (November 11th)"
    End If

    If Holidays(9) <> #11/26/2026# Then
        TestGetUSBankHolidays = False
        Debug.Print "Thanksgiving Day (Fourth Thursday in November)"
    End If

    If Holidays(10) <> #12/25/2026# Then
        TestGetUSBankHolidays = False
        Debug.Print "Christmas Day"
    End If

    Debug.Print "TestGetUSBankHolidays: " & TestGetUSBankHolidays

End Function
