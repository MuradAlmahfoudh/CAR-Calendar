VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
    
Private Sub Document_Open()
    
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++'
    Dim loc_Copyright As String
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++'
    
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++'
    If func_VerifyFileName = False Then
    
        ThisDocument.Close
    
    End If

    Call subr_ReadCCSFile
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++'
    
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++'
    ThisDocument.Unprotect "Q2FsZW5kYXI=496eU2Vjb25kcw=="
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++'
            
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++'
    Call subr_AnalyzeCCSFileContent
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++'
    
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++'
    On Error GoTo START_CUSTOMIZATION
    If glo_CCSFileCommand.Item(con_CCSC_REGENERATE) = con_CCSC_REGEN_OFF Then
    
        GoTo EXIT_SUB
    
    End If
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++'
        
START_CUSTOMIZATION:

    Call subr_StartCustomization
        
    UndoClear

EXIT_SUB:
    
    ThisDocument.Protect wdAllowOnlyReading, , "Q2FsZW5kYXI=496eU2Vjb25kcw=="
    
    ThisDocument.Save

End Sub
