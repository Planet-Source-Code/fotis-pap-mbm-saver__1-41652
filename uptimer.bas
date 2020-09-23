Attribute VB_Name = "uptimer"
'from planet source code
Public Declare Function GetTickCount Lib "kernel32.dll" () As Long
Public Enum TimeFormatType
    DaysHoursMinutesSecondsMilliseconds = 0
    DaysHoursMinutesSeconds = 1
    DaysHoursMinutes = 2
    HoursMinutesSecondsMilliseconds = 3
    HoursMinutesSeconds = 4
    HoursMinutes = 5
    HMSColonSeparated = 6
End Enum
Public Function FormatCount(Count As Long, Optional FormatType As TimeFormatType = 0) As String
Dim Days As Long, Hours As Long, Minutes As Long, Seconds As Long, Miliseconds As Long
    
    Miliseconds = Count Mod 1000
    Count = Count \ 1000
    Days = Count \ (24& * 3600&)
    If Days > 0 Then Count = Count - (24& * 3600& * Days)
    Hours = Count \ 3600&
    If Hours > 0 Then Count = Count - (3600& * Hours)
    Minutes = Count \ 60
    Seconds = Count Mod 60

    Select Case FormatType
        Case 0
            FormatCount = Days & " Days, " & Hours & " Hours, " & _
            Minutes & " Minutes, " & Seconds & " Seconds, " & Miliseconds & _
            " Milliseconds"
        Case 1
            FormatCount = Days & " Days, " & Hours & " Hours, " & _
            Minutes & " Minutes, " & Seconds & " Seconds"
        Case 2
            FormatCount = Days & " Days, " & Hours & " Hours, " & _
            Minutes & " Minutes"
        Case 3
            FormatCount = Hours & " Hours, " & Minutes & " Minutes, " & _
            Seconds & " Seconds, " & Miliseconds & " Milliseconds"
        Case 4
            FormatCount = Hours & " Hours, " & Minutes & " Minutes, " & _
            Seconds & " Seconds"
        Case 5
            FormatCount = Hours & " Hours, " & Minutes & " Minutes"
        Case 6
            FormatCount = Hours & ":" & Minutes & ":" & Seconds
    End Select
End Function

