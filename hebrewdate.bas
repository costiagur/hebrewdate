Attribute VB_Name = "hebrewdate"
Function hebrewdate(thedate)

Dim hebdate As String
Dim hebday As Integer, hebdaylet As String
Dim hebyear As Integer, hebyearlet As String

hebdate = Excel.WorksheetFunction.Text(thedate, "[$-8040D]ddmmmyyyy")

hebday = Mid(hebdate, 1, 2) * 1

hebdaylet = hebdayconv(hebday)

hebyear = Mid(hebdate, Len(hebdate) - 3, 4) * 1

hebyearlet = hebyearconv(hebyear)

hebrewdate = hebdaylet & " " & Chr(225) & Mid(hebdate, 3, Len(hebdate) - 6) & " " & hebyearlet

End Function


Private Function hebdayconv(hebday As Integer)
Dim un As Integer, dec As Integer
Dim unletter As String, decletter As String

decletter = ""

un = (hebday / 10 - Excel.WorksheetFunction.RoundDown(hebday / 10, 0)) * 10

If hebday < 10 Then
    dec = 0
Else
    dec = (hebday - un) / 10
End If

Select Case hebday
    
    Case 15
        unletter = Chr(229)
    
    Case 16
        unletter = Chr(230)
    
    Case Else
        If un = 0 Then
            unletter = ""
        Else
            unletter = Chr(223 + un)
        End If

End Select

Select Case hebday
    Case 15
        decletter = Chr(232)
        
    Case 16
        decletter = Chr(232)
                   
    Case Is > 19
        decletter = Chr(233 + dec)
        
    Case Is > 9
        decletter = Chr(233)
End Select

hebdayconv = decletter & unletter & "'"

End Function


Private Function hebyearconv(hebyear)
Dim un As Integer, dec As Integer
Dim unletter As String, decletter As String

un = (hebyear / 10 - Excel.WorksheetFunction.RoundDown(hebyear / 10, 0)) * 10

If un = 0 Then
    unletter = ""
Else
    unletter = Chr(223 + un)
End If

dec = (hebyear - 5700 - un) / 10

Select Case dec
    Case 1
        decletter = Chr(233)
    Case 2
        decletter = Chr(235)
    Case 3
        decletter = Chr(236)
    Case 4
        decletter = Chr(238)
    Case 5
        decletter = Chr(240)
    Case 6
        decletter = Chr(241)
    Case 7
        decletter = Chr(242)
    Case 8
        decletter = Chr(244)
    Case 9
        decletter = Chr(246)
End Select
        
hebyearconv = Chr(250) & Chr(249) & IIf(un = 0, Chr(34), "") & decletter & IIf(un = 0, "", Chr(34)) & unletter

End Function
