Attribute VB_Name = "DocumentOperations"
Option Explicit

Public Sub subr_StartCustomization()
    
    Dim loc_Cnt_Month As Integer
    Dim loc_Cnt_Row  As Integer
    Dim loc_Cnt_Col  As Integer
    
    Dim loc_Cnt_Day  As Integer
    Dim loc_Max_Day  As Integer
    
    Dim loc_Cnt_WDay As Integer
    
    Dim loc_WKEnd_Day1 As Integer
    Dim loc_WKEnd_Day2 As Integer
    
    Dim loc_Month_Name As String
    
    Dim loc_legend_rowid As Integer
    
    
    On Error Resume Next
    
    loc_legend_rowid = 1
    
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++'
    ' > Font
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++'
    glo_FontName = glo_CCSFileCommand.Item(con_CCSC_FONTNAME)
    
    If func_ValidateFont(glo_FontName) = False Then
    
        glo_FontName = con_DEFAULT_FONT
    
    End If
    
    ThisDocument.Sections(1).Headers(wdHeaderFooterPrimary).Range.Font.Name = glo_FontName
    ThisDocument.Sections(1).Footers(wdHeaderFooterPrimary).Range.Font.Name = glo_FontName
    ThisDocument.Sections(1).Range.Font.Name = glo_FontName
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++'
    
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++'
    ' > Title And Subtitle
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++'
    glo_Title = con_DEFAULT_TITLE
    glo_Title = glo_CCSFileCommand.Item(con_CCSC_TITLE)
    ThisDocument.Sections(1).Headers(wdHeaderFooterPrimary).Range.Tables(1).Cell(1, 2).Range = glo_Title
    
    glo_SubTitle = con_DEFAULT_SUBTITLE
    glo_SubTitle = glo_CCSFileCommand.Item(con_CCSC_SUBTITLE)
    ThisDocument.Sections(1).Headers(wdHeaderFooterPrimary).Range.Tables(1).Cell(2, 2).Range = glo_SubTitle
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++'
    
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++'
    ' > Year (Default - Current Year)
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++'
    glo_Year = CInt(glo_CCSFileCommand.Item(con_CCSC_YEAR))
    
    If glo_Year = 0 Then
    
        glo_Year = Year(Date)
    
    End If
    
    ThisDocument.Sections(1).Headers(wdHeaderFooterPrimary).Range.Tables(1).Cell(1, 3).Range = glo_Year
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++'
        
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++'
    ' > Clear Legend
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++'
    ThisDocument.Sections(1).Footers(wdHeaderFooterPrimary).Range.Tables(1).Shading.ForegroundPatternColor = RGB(255, 255, 255)
    ThisDocument.Sections(1).Footers(wdHeaderFooterPrimary).Range.Tables(1).Range.Delete
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++'
        
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++'
    ' > Weekend Text
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++'
    glo_WKEDay = con_DEFAULT_WEND
    glo_WKEDay = glo_CCSFileCommand.Item(con_CCSC_WKEND)

    If glo_WKEDay <> "none" Then

        ThisDocument.Sections(1).Footers(wdHeaderFooterPrimary).Range.Tables(1).Cell(loc_legend_rowid, 1).Shading.ForegroundPatternColor = RGB(func_ReadColorComponent(con_SHADE_WEEKEND, "R"), func_ReadColorComponent(con_SHADE_WEEKEND, "G"), func_ReadColorComponent(con_SHADE_WEEKEND, "B"))
        ThisDocument.Sections(1).Footers(wdHeaderFooterPrimary).Range.Tables(1).Cell(loc_legend_rowid, 2).Range = "  " & "Weekend"

        loc_legend_rowid = loc_legend_rowid + 1

    End If
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++'

    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++'
    ' > Holiday Text
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++'
    glo_Holiday = glo_CCSFileCommand.Item(con_CCSC_HOLIDAY)

    glo_HLText = con_DEFAULT_HLTEXT
    glo_HLText = glo_CCSFileCommand.Item(con_CCSC_HLTEXT)

    If glo_Holiday <> "" Then

        glo_HLText = ""
        Erase glo_ArrHoliday

    End If

    If glo_HLText <> "" Then

        ThisDocument.Sections(1).Footers(wdHeaderFooterPrimary).Range.Tables(1).Cell(loc_legend_rowid, 1).Shading.ForegroundPatternColor = RGB(func_ReadColorComponent(con_SHADE_HOLIDAY, "R"), func_ReadColorComponent(con_SHADE_HOLIDAY, "G"), func_ReadColorComponent(con_SHADE_HOLIDAY, "B"))
        ThisDocument.Sections(1).Footers(wdHeaderFooterPrimary).Range.Tables(1).Cell(loc_legend_rowid, 2).Range = "  " & glo_HLText

        loc_legend_rowid = loc_legend_rowid + 1

    End If
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++'

    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++'
    ' > Custom Text
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++'
    glo_Custom = glo_CCSFileCommand.Item(con_CCSC_CUSTOM)

    glo_CSText = con_DEFAULT_CSTEXT
    glo_CSText = glo_CCSFileCommand.Item(con_CCSC_CSTEXT)

    If glo_Custom <> "" Then

        glo_CSText = ""
        Erase glo_ArrCustom

    End If

    If glo_CSText <> "" Then

        ThisDocument.Sections(1).Footers(wdHeaderFooterPrimary).Range.Tables(1).Cell(loc_legend_rowid, 1).Shading.ForegroundPatternColor = RGB(func_ReadColorComponent(con_SHADE_CUSTOM, "R"), func_ReadColorComponent(con_SHADE_CUSTOM, "G"), func_ReadColorComponent(con_SHADE_CUSTOM, "B"))
        ThisDocument.Sections(1).Footers(wdHeaderFooterPrimary).Range.Tables(1).Cell(loc_legend_rowid, 2).Range = "  " & glo_CSText

    End If
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++'
    
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++'
    ' > Week Start Day (Default - Monday)
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++'
    glo_WKSDay = con_DEFAULT_WSDAY
    glo_WKSDay = glo_CCSFileCommand.Item(con_CCSC_WKSTART)
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++'
    
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++'
    ' > Selecting Week Template Based On Week Start Day
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++'
    glo_ArrWeekTempl = func_GenerateWeekTemplate(con_CCSC_WS_VALUES, glo_WKSDay)
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++'
    
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++'
    ' > Determine Positions of Weekend Days To Highlight
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++'
    If glo_WKEDay = "none" Then
    
        loc_WKEnd_Day1 = -1
        loc_WKEnd_Day2 = -1
    
    Else
        
        For loc_Cnt_WDay = 0 To 6
        
            If Left(glo_WKEDay, 3) = glo_ArrWeekTempl(loc_Cnt_WDay) Then
            
                Exit For
                
            End If
        
        Next
        
        loc_WKEnd_Day1 = loc_Cnt_WDay + 1
        
        If Len(glo_WKEDay) > 3 Then
        
            loc_WKEnd_Day2 = loc_WKEnd_Day1 + 1
            
            If loc_WKEnd_Day2 > 7 Then
                
                loc_WKEnd_Day2 = loc_WKEnd_Day2 - 7
                
            End If
            
            loc_WKEnd_Day1 = loc_WKEnd_Day1 + 1
            loc_WKEnd_Day2 = loc_WKEnd_Day2 + 1
            
        Else
            loc_WKEnd_Day2 = -1
            loc_WKEnd_Day1 = loc_WKEnd_Day1 + 1
        End If
                
    End If
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++'
    
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++'
    ' > Initializing Calendar (Month by Month)
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++'
    For loc_Cnt_Month = 1 To 12
        
        ThisDocument.Tables(loc_Cnt_Month).Shading.ForegroundPatternColor = RGB(255, 255, 255)
        ThisDocument.Tables(loc_Cnt_Month).Range.Delete
        
    Next
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++'

    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++'
    ' > Populating Month Names And Week Days (Header)
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++'
    For loc_Cnt_Month = 1 To 12
        
        For loc_Cnt_Row = 1 To 1
        
            For loc_Cnt_Col = 2 To 2
                
                Select Case loc_Cnt_Month
                
                    Case 1
                        loc_Month_Name = con_MONTH_JAN
                    
                    Case 2
                        loc_Month_Name = con_MONTH_FEB
                        
                    Case 3
                        loc_Month_Name = con_MONTH_MAR
                    
                    Case 4
                        loc_Month_Name = con_MONTH_APR
                        
                    Case 5
                        loc_Month_Name = con_MONTH_MAY
                    
                    Case 6
                        loc_Month_Name = con_MONTH_JUN
                    
                    Case 7
                        loc_Month_Name = con_MONTH_JUL
                    
                    Case 8
                        loc_Month_Name = con_MONTH_AUG
                    
                    Case 9
                        loc_Month_Name = con_MONTH_SEP
                        
                    Case 10
                        loc_Month_Name = con_MONTH_OCT
                    
                    Case 11
                        loc_Month_Name = con_MONTH_NOV
                    
                    Case 12
                        loc_Month_Name = con_MONTH_DEC
                
                End Select
                
                ThisDocument.Tables(loc_Cnt_Month).Cell(loc_Cnt_Row, loc_Cnt_Col).Range = loc_Month_Name
                
            Next
        
        Next
        
        For loc_Cnt_Row = 2 To 2
        
            For loc_Cnt_Col = 2 To 8
            
                ThisDocument.Tables(loc_Cnt_Month).Cell(loc_Cnt_Row, loc_Cnt_Col).Range = func_FormatWeekDay(glo_ArrWeekTempl(loc_Cnt_Col - 2))
            
            Next
        
        Next
    
    Next
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++'
    
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++'
    ' > Highlighting Weekends
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++'
    
    For loc_Cnt_Month = 1 To 12
        
        For loc_Cnt_Row = 2 To 8
            
            If loc_WKEnd_Day1 <> -1 Then
            
                ThisDocument.Tables(loc_Cnt_Month).Cell(loc_Cnt_Row, loc_WKEnd_Day1).Shading.ForegroundPatternColor = RGB(func_ReadColorComponent(con_SHADE_WEEKEND, "R"), func_ReadColorComponent(con_SHADE_WEEKEND, "G"), func_ReadColorComponent(con_SHADE_WEEKEND, "B"))
            
            End If
            
            If loc_WKEnd_Day2 <> -1 Then
            
                ThisDocument.Tables(loc_Cnt_Month).Cell(loc_Cnt_Row, loc_WKEnd_Day2).Shading.ForegroundPatternColor = RGB(func_ReadColorComponent(con_SHADE_WEEKEND, "R"), func_ReadColorComponent(con_SHADE_WEEKEND, "G"), func_ReadColorComponent(con_SHADE_WEEKEND, "B"))
            
            End If
        
        Next
    
    Next
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++'
    
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++'
    ' > Populating Calendar Days
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++'
    For loc_Cnt_Month = 1 To 12
        
        loc_Max_Day = Day(DateSerial(glo_Year, loc_Cnt_Month + 1, 0))
        
        For loc_Cnt_Day = 1 To loc_Max_Day
        
            Call subr_PlaceDayInTableCell(DateSerial(glo_Year, loc_Cnt_Month, loc_Cnt_Day))
        
        Next
            
    Next
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++'

End Sub

Private Sub subr_PlaceDayInTableCell(ByVal inp_Date As Date)

    Static loc_MoWkNo   As Integer

    Dim loc_Row_No      As Integer
    Dim loc_Col_No      As Integer
    
    Dim loc_Week_No     As Integer
    Dim loc_WSDay       As Integer
    Dim loc_WDay        As String
    
    Dim loc_WDayTx      As String
    
    Dim loc_Counter     As Integer
    
    Select Case glo_WKSDay

        Case "sat"

            loc_WSDay = vbSaturday

        Case "sun"

            loc_WSDay = vbSunday

        Case "mon"

            loc_WSDay = vbMonday
            
        Case "tue"

            loc_WSDay = vbTuesday
            
        Case "wed"

            loc_WSDay = vbWednesday
        
        Case "thu"

            loc_WSDay = vbThursday
            
        Case "fri"

            loc_WSDay = vbFriday

    End Select
    
    ' Determine week day for input date
    loc_WDay = Weekday(inp_Date)
    
    ' Determine week number
    loc_Week_No = Format(inp_Date, "ww", loc_WSDay)
    
    Select Case loc_WDay
    
        Case 1
            
            loc_WDayTx = "sun"
            
        Case 2
        
            loc_WDayTx = "mon"
        
        Case 3
        
            loc_WDayTx = "tue"
        
        Case 4
        
            loc_WDayTx = "wed"
        
        Case 5
        
            loc_WDayTx = "thu"
        
        Case 6
        
            loc_WDayTx = "fri"
        
        Case 7
        
            loc_WDayTx = "sat"
    
    End Select
    
    For loc_Counter = 0 To UBound(glo_ArrWeekTempl)
    
        If loc_WDayTx = glo_ArrWeekTempl(loc_Counter) Then
        
            Exit For
        
        End If
    
    Next
    
    loc_Col_No = loc_Counter + 2
    
    ' Determine week number for current date of current month
    If Day(inp_Date) = 1 Then
        
        loc_MoWkNo = 1
    
    ElseIf Weekday(inp_Date) = loc_WSDay Then
        
        loc_MoWkNo = loc_MoWkNo + 1
        
    End If
    
    ' Finding table cell to place current week number
    loc_Row_No = loc_MoWkNo + 2
    
    ' Populating week number
    If (Day(inp_Date) = 1) Or (Weekday(inp_Date) = loc_WSDay) Then
    
        ThisDocument.Tables(Month(inp_Date)).Cell(loc_Row_No, 1).Range = loc_Week_No
    
    End If
    
    ' Populating day
    ThisDocument.Tables(Month(inp_Date)).Cell(loc_Row_No, loc_Col_No).Range = Day(inp_Date)
    
    ' Added new to continue if holiday none
    On Error GoTo CHECK_CUSTOM
    
    ' Highlighting day if flagged as holiday
    If InStr(1, glo_ArrHoliday(Month(inp_Date)), "(" & Trim(Str(Day(inp_Date))) & ")", vbTextCompare) > 0 Then
        
        ThisDocument.Tables(Month(inp_Date)).Cell(loc_Row_No, loc_Col_No).Shading.ForegroundPatternColor = RGB(func_ReadColorComponent(con_SHADE_HOLIDAY, "R"), func_ReadColorComponent(con_SHADE_HOLIDAY, "G"), func_ReadColorComponent(con_SHADE_HOLIDAY, "B"))
        
    End If
    
CHECK_CUSTOM:
    
    ' Highlighting day if flagged as custom
    If InStr(1, glo_ArrCustom(Month(inp_Date)), "(" & Trim(Str(Day(inp_Date))) & ")", vbTextCompare) > 0 Then
        
        ThisDocument.Tables(Month(inp_Date)).Cell(loc_Row_No, loc_Col_No).Shading.ForegroundPatternColor = RGB(func_ReadColorComponent(con_SHADE_CUSTOM, "R"), func_ReadColorComponent(con_SHADE_CUSTOM, "G"), func_ReadColorComponent(con_SHADE_CUSTOM, "B"))
        
    End If

End Sub

Private Function func_FormatWeekDay(ByVal inp_WKDay As String) As String

    Dim loc_FWKDay As String
    
    Select Case inp_WKDay
    
        Case "sat"
        
            loc_FWKDay = "SA"
        
        Case "sun"
        
            loc_FWKDay = "SU"
        
        Case "mon"
        
            loc_FWKDay = "MO"
        
        Case "tue"
        
            loc_FWKDay = "TU"
        
        Case "wed"
        
            loc_FWKDay = "WE"
        
        Case "thu"
        
            loc_FWKDay = "TH"
        
        Case "fri"
        
            loc_FWKDay = "FR"
    
    End Select
    
        func_FormatWeekDay = loc_FWKDay

End Function

Private Function func_ReadColorComponent(ByVal inp_RGBColorText As String, ByVal inp_CCompID As String) As Integer

    Dim loc_ArrRGBColorComp() As String
    Dim loc_ColorComp         As Integer
    
    loc_ArrRGBColorComp = Split(inp_RGBColorText, ",", , vbTextCompare)
    
    Select Case inp_CCompID
    
        Case "R"
            
            loc_ColorComp = CInt(loc_ArrRGBColorComp(0))
            
        Case "G"
        
            loc_ColorComp = CInt(loc_ArrRGBColorComp(1))
        
        Case "B"
        
            loc_ColorComp = CInt(loc_ArrRGBColorComp(2))
    
    End Select
    
    func_ReadColorComponent = loc_ColorComp

End Function

Private Function func_GenerateWeekTemplate(ByVal inp_WeekDays As String, ByVal inp_WKSDay As String) As String()

    Dim loc_WeekTempl As String
    Dim loc_WSDP As Integer
    
    On Error Resume Next
    loc_WSDP = InStr(1, inp_WeekDays, inp_WKSDay, vbTextCompare)
    loc_WeekTempl = Left(inp_WeekDays, loc_WSDP - 2)
    loc_WeekTempl = Mid(inp_WeekDays, loc_WSDP, Len(inp_WeekDays) - loc_WSDP + 1) & "," & loc_WeekTempl
    
    func_GenerateWeekTemplate = Split(loc_WeekTempl, ",", , vbTextCompare)

End Function

Private Function func_ValidateFont(ByVal inpt_FontName As String) As Boolean

    Dim loc_Cnt_Font As Long
    
    For loc_Cnt_Font = 1 To Application.FontNames.Count
    
        If UCase(Application.FontNames(loc_Cnt_Font)) = UCase(inpt_FontName) Then
        
            func_ValidateFont = True
            Exit Function
            
        End If
    
    Next
    
    func_ValidateFont = False

End Function
