Attribute VB_Name = "FileOperations"
Option Explicit

Public Function func_VerifyFileName() As Boolean

    Dim loc_ExpectedDocName     As String
    Dim loc_ArrDocNameToken()   As String
    
    loc_ExpectedDocName = con_FILE_PUPLISHER + "_" + con_CALENDAR_NAME + "_" + con_PAPER_SIZE
    
    loc_ArrDocNameToken = Split(ThisDocument.Name, ".", , vbTextCompare)
    
    If UCase(loc_ArrDocNameToken(0)) <> loc_ExpectedDocName Then
        
        func_VerifyFileName = False
        Exit Function
        
    End If
    
    func_VerifyFileName = True

End Function

Public Sub subr_ReadCCSFile()
    
    Dim loc_FileName        As String
    Dim loc_FileID          As Integer
    Dim loc_FileSize        As Long
    

    loc_FileName = func_GetCCSFileName
    
    If loc_FileName = "" Then
        
        Exit Sub
        
    End If
            
    loc_FileID = FreeFile
    
    Open loc_FileName For Input Access Read As loc_FileID
    
    loc_FileSize = FileLen(loc_FileName)
    
    glo_CCSFileContent = Input$(loc_FileSize, loc_FileID)

    Close loc_FileID
    
End Sub

Private Function func_GenerateCCSFileName() As String
    
    Dim loc_CCSFileName As String
    
    loc_CCSFileName = con_FILE_PUPLISHER + "_" + con_CALENDAR_NAME + "_" + con_CSSFilePostfix + "." + con_CCSFileExtension
    
    func_GenerateCCSFileName = loc_CCSFileName

End Function

Private Function func_GetCCSFileName() As String

    Dim loc_CCSLgFileName As String
    
    loc_CCSLgFileName = ThisDocument.Path + "\" + func_GenerateCCSFileName
    
    If Dir(loc_CCSLgFileName) = "" Then
        
        Exit Function
        
    End If
    
    func_GetCCSFileName = loc_CCSLgFileName
    
End Function
