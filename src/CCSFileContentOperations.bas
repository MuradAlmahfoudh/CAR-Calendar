Attribute VB_Name = "CCSFileContentOperations"
Option Explicit

Public Sub subr_AnalyzeCCSFileContent()
    
    Dim loc_ArrCCSFileEntry() As String
    Dim loc_Counter As Integer
    
    Set glo_CCSFileCommand = New Collection
    ReDim glo_ArrHoliday(1 To 12)
    ReDim glo_ArrCustom(1 To 12)
    
    loc_ArrCCSFileEntry = Split(glo_CCSFileContent, vbCrLf)
    
    For loc_Counter = 0 To UBound(loc_ArrCCSFileEntry)
        
        Call subr_RegisterCCSCommand(loc_ArrCCSFileEntry(loc_Counter))
        
    Next

End Sub

Private Sub subr_RegisterCCSCommand(ByVal inp_Command As String)

    Dim loc_ArrCommandComp()    As String
    Dim loc_CommandName         As String
    Dim loc_CommandParam        As String
    
    Dim loc_CollectionValue     As Variant
    
    
    loc_ArrCommandComp = Split(inp_Command, con_CCSCmdInpSwitch, , vbTextCompare)
    
    
    On Error GoTo SKIP_COMMAND_REGISTRATION

    loc_CommandName = LCase(loc_ArrCommandComp(0))
    loc_CommandParam = LCase(loc_ArrCommandComp(1))
    
    
    Select Case loc_CommandName
        
        '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++'
        ' COMMAND > REGENERATE
        '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++'
        Case con_CCSC_REGENERATE
        
        On Error GoTo REGENERATE_NO_ENTRY_FOUND
        loc_CollectionValue = glo_CCSFileCommand.Item(con_CCSC_REGENERATE)
        
        GoTo SKIP_COMMAND_REGISTRATION
        
REGENERATE_NO_ENTRY_FOUND:
        
        If Not (loc_CommandParam = con_CCSC_REGEN_ON Or loc_CommandParam = con_CCSC_REGEN_OFF) Then
        
            GoTo SKIP_COMMAND_REGISTRATION
        
        End If
        '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++'
        
        
        '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++'
        ' COMMAND > YEAR
        '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++'
        Case con_CCSC_YEAR
        
            On Error GoTo YEAR_NO_ENTRY_FOUND
            loc_CollectionValue = CInt(glo_CCSFileCommand.Item(con_CCSC_YEAR))
            
            GoTo SKIP_COMMAND_REGISTRATION

YEAR_NO_ENTRY_FOUND:

            If IsNumeric(loc_CommandParam) = False Then
                
                GoTo SKIP_COMMAND_REGISTRATION
                
            End If
            
            If Not (CInt(loc_CommandParam) >= CInt(con_CCSC_YEAR_MIN) And CInt(loc_CommandParam) <= CInt(con_CCSC_YEAR_MAX)) Then
            
                GoTo SKIP_COMMAND_REGISTRATION
            
            End If
        '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++'
        
        
        '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++'
        ' COMMAND > WEEKSTART
        '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++'
        Case con_CCSC_WKSTART
        
            On Error GoTo WKSTART_NO_ENTRY_FOUND
            loc_CollectionValue = glo_CCSFileCommand.Item(con_CCSC_WKSTART)
            
            GoTo SKIP_COMMAND_REGISTRATION

WKSTART_NO_ENTRY_FOUND:
            
            If Len(loc_CommandParam) = 0 Then
                
                GoTo SKIP_COMMAND_REGISTRATION
                
            End If
            
            If InStr(1, con_CCSC_WS_VALUES, loc_CommandParam, vbTextCompare) = 0 Then
                
                GoTo SKIP_COMMAND_REGISTRATION
                
            End If
        '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++'
        
        
        '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++'
        ' COMMAND > WEEKEND
        '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++'
        Case con_CCSC_WKEND
        
            On Error GoTo WKEND_NO_ENTRY_FOUND
            loc_CollectionValue = glo_CCSFileCommand.Item(con_CCSC_WKEND)
            
            GoTo SKIP_COMMAND_REGISTRATION
                
WKEND_NO_ENTRY_FOUND:

            If Len(loc_CommandParam) = 0 Then
                
                GoTo SKIP_COMMAND_REGISTRATION
                
            End If

            If InStr(1, con_CCSC_WE_VALUES, loc_CommandParam, vbTextCompare) = 0 Then
                
                GoTo SKIP_COMMAND_REGISTRATION
                
            End If
        '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++'
        
        
        '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++'
        ' COMMAND > HOLIDAY
        '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++'
        Case con_CCSC_HOLIDAY
        
            If IsDate(loc_CommandParam) = False And loc_CommandParam = "none" Then
            
                On Error GoTo HOLIDAY_NO_ENTRY_FOUND
                loc_CollectionValue = glo_CCSFileCommand.Item(con_CCSC_HOLIDAY)
                
                GoTo SKIP_COMMAND_REGISTRATION
                    
HOLIDAY_NO_ENTRY_FOUND:
                    
            Else
            
                Call subr_RegisterSpecialDay(glo_ArrHoliday, loc_CommandParam)
                GoTo SKIP_COMMAND_REGISTRATION
                
            End If
        '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++'
        
        
        '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++'
        ' COMMAND > CUSTOM
        '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++'
        Case con_CCSC_CUSTOM
            
            If IsDate(loc_CommandParam) = False And loc_CommandParam = "none" Then
            
                On Error GoTo CUSTOM_NO_ENTRY_FOUND
                loc_CollectionValue = glo_CCSFileCommand.Item(con_CCSC_CUSTOM)
                
                GoTo SKIP_COMMAND_REGISTRATION
                    
CUSTOM_NO_ENTRY_FOUND:
                    
            Else
            
                Call subr_RegisterSpecialDay(glo_ArrCustom, loc_CommandParam)
                GoTo SKIP_COMMAND_REGISTRATION
                
            End If
        '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++'
        
        
        '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++'
        ' COMMAND > TITLE
        '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++'
        Case con_CCSC_TITLE
        
            On Error GoTo TITLE_NO_ENTRY_FOUND
            loc_CollectionValue = glo_CCSFileCommand.Item(con_CCSC_TITLE)
            
            GoTo SKIP_COMMAND_REGISTRATION
        
TITLE_NO_ENTRY_FOUND:

            loc_CommandParam = loc_ArrCommandComp(1)
        '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++'
        
        
        '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++'
        ' COMMAND > SUBTITLE
        '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++'
        Case con_CCSC_SUBTITLE
            
            On Error GoTo SUBTITLE_NO_ENTRY_FOUND
            loc_CollectionValue = glo_CCSFileCommand.Item(con_CCSC_SUBTITLE)
            
            GoTo SKIP_COMMAND_REGISTRATION
        
SUBTITLE_NO_ENTRY_FOUND:

            loc_CommandParam = loc_ArrCommandComp(1)
        '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++'
        
        
        '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++'
        ' COMMAND > CUSTOMTEXT
        '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++'
        Case con_CCSC_CSTEXT
            
            On Error GoTo CUSTOMTEXT_NO_ENTRY_FOUND
            loc_CollectionValue = glo_CCSFileCommand.Item(con_CCSC_CSTEXT)
            
            GoTo SKIP_COMMAND_REGISTRATION
        
CUSTOMTEXT_NO_ENTRY_FOUND:

            loc_CommandParam = loc_ArrCommandComp(1)
        '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++'
        
        
        '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++'
        ' COMMAND > HOLIDAYTEXT
        '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++'
        Case con_CCSC_HLTEXT
        
            On Error GoTo HOLIDAYTEXT_NO_ENTRY_FOUND
            loc_CollectionValue = glo_CCSFileCommand.Item(con_CCSC_HLTEXT)
            
            GoTo SKIP_COMMAND_REGISTRATION
        
HOLIDAYTEXT_NO_ENTRY_FOUND:

            loc_CommandParam = loc_ArrCommandComp(1)
        '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++'
        
        
        '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++'
        ' COMMAND > FONTNAME
        '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++'
        Case con_CCSC_FONTNAME
        
            On Error GoTo FONTNAME_NO_ENTRY_FOUND
            loc_CollectionValue = glo_CCSFileCommand.Item(con_CCSC_FONTNAME)
            
            GoTo SKIP_COMMAND_REGISTRATION
        
FONTNAME_NO_ENTRY_FOUND:
        '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++'
        
        
        '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++'
        Case Else
        
            GoTo SKIP_COMMAND_REGISTRATION
        '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++'
            
    End Select
    
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++'
    ' > Registering Command
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++'
    glo_CCSFileCommand.Add loc_CommandParam, loc_CommandName
    
    Exit Sub
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++'
        
        
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++'
' SKIP_COMMAND_REGISTRATION is triggered when the command encountered was previously registered (e.g. command exists twice or more in the CCS File)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++'
SKIP_COMMAND_REGISTRATION:
    
End Sub

Private Sub subr_RegisterSpecialDay(ByRef ref_ArrSpDay() As String, ByVal inp_SpDayText As String)
    
    Dim loc_Year           As String
    
    Dim loc_ArrDateEntry() As String
    Dim loc_ArrDateRange() As String
    
    Dim loc_Counter00      As Integer
    
    Dim loc_DateFr         As Date
    Dim loc_DateTo         As Date
    
    Dim loc_Date_STR       As String
    Dim loc_Date           As Date
    Dim loc_Day            As String
    
    loc_Year = glo_CCSFileCommand(con_CCSC_YEAR)
    
    loc_ArrDateEntry = Split(inp_SpDayText, ",", , vbTextCompare)
    
    For loc_Counter00 = 0 To UBound(loc_ArrDateEntry)
        
        On Error Resume Next
        
        ' > Single Date Entry
        If InStr(1, loc_ArrDateEntry(loc_Counter00), "-", vbTextCompare) = 0 Then
        
            loc_Date_STR = loc_ArrDateEntry(loc_Counter00) + "/" + loc_Year
            loc_Date = DateValue(loc_Date_STR)
            
            If IsDate(loc_Date_STR) = False Then
            
                GoTo NEXT_DATE_ENTRY
                
            End If
            
            loc_Day = Day(loc_Date)
            
            ref_ArrSpDay(Month(loc_Date)) = ref_ArrSpDay(Month(loc_Date)) + "(" + loc_Day + ")"
        
        ' > Date Range Entry
        Else
        
            loc_ArrDateRange = Split(loc_ArrDateEntry(loc_Counter00), "-", , vbTextCompare)
            
            loc_ArrDateRange(0) = loc_ArrDateRange(0) + "/" + loc_Year
            loc_ArrDateRange(1) = loc_ArrDateRange(1) + "/" + loc_Year
            
            loc_DateFr = DateValue(loc_ArrDateRange(0))
            loc_DateTo = DateValue(loc_ArrDateRange(1))
            
            If IsDate(loc_ArrDateRange(0)) = False Or IsDate(loc_ArrDateRange(1)) = False Then
            
                GoTo NEXT_DATE_ENTRY
                
            End If
            
            If loc_DateFr >= loc_DateTo Then
                
                GoTo NEXT_DATE_ENTRY
                
            End If
            
            While loc_DateFr <= loc_DateTo
                
                loc_Day = Day(loc_DateFr)
                ref_ArrSpDay(Month(loc_DateFr)) = ref_ArrSpDay(Month(loc_DateFr)) + "(" + loc_Day + ")"
                
                loc_DateFr = loc_DateFr + 1
            
            Wend
        
        End If
        
NEXT_DATE_ENTRY:
        
    Next

End Sub
