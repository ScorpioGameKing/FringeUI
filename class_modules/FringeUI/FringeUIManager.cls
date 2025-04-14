VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FringeUIManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----- FRINGEUI: VBA Wasn't Meant For This
' Author: ScorpioGameKing
' Version: 0.3.0

'----- XML INJECTION PHASE
Sub BuildFringeUI(package As Collection, Optional DEBUGMODE As Boolean)

    ClearFringeUI DEBUGMODE

    '----- Prep Variables
    Dim hFile As Long: hFile = FreeFile
    Dim path As String: path = "C:\Users\" & Environ("Username") & "\AppData\Local\Microsoft\Office\"
    Dim fileName As String: fileName = "Excel.officeUI"
    Dim ribbonXML As String
    
    If Not RealityCheck.IsClassModuleLoaded("FringeUIMultiLoader") Then
        Dim reloadButton As String, reloadGroup As String
        reloadButton = "<mso:group id='FringeUIManagerTools'" & _
        " label='Fringe UI Tools'" & _
        " autoScale='true'>" & vbNewLine & _
        "<mso:button id='FringeUIReload'" & _
        " label='Re-Load FringeUI'" & _
        " imageMso='GroupPivotTableData'" & _
        " onAction='FringeUIReLoader.ReLoadUI'/>" & vbNewLine & "</mso:group>" & vbNewLine
    End If
    
    '----- Nothing to do
    ribbonXML = "<mso:customUI xmlns:mso='http://schemas.microsoft.com/office/2009/07/customui'>" & vbNewLine & _
    "<mso:ribbon startFromScratch='false'>" & vbNewLine & _
    "<mso:tabs>" & vbNewLine
    
    '----- Start Tab Injections
    For Each tab_check In package
    
        '----- Validate
        If tab_check.Item(1) = "TAB" Then
        
            '----- Insert Header, Iterate
            ribbonXML = ribbonXML & tab_check.Item(2)
            If Not RealityCheck.IsClassModuleLoaded("FringeUIMultiLoader") Then
                ribbonXML = ribbonXML & reloadButton
            End If
            For Each tab_group_check In tab_check.Item(3)
            
                '----- Validate
                If tab_group_check.Item(1) = "GROUP" Then
                
                    '----- Insert Header, Iterate
                    ribbonXML = ribbonXML & tab_group_check.Item(2)
                    For Each tab_obj_check In tab_group_check.Item(3)
                    
                        '----- Validate
                        If tab_obj_check.Item(1) = "BUTTON" Then
                        
                            '----- Insert Button
                            ribbonXML = ribbonXML & tab_obj_check.Item(2)
                        End If
                    Next
                    
                    '----- Insert Group Exit
                    ribbonXML = ribbonXML & "</mso:group>" & vbNewLine
                End If
            Next
            
            '----- Insert Tab Exit
            ribbonXML = ribbonXML & "</mso:tab>" & vbNewLine
        End If
    Next
    
    '----- Close Out Misc Tags
    ribbonXML = ribbonXML & "</mso:tabs>" & vbNewLine & _
    "</mso:ribbon>" & vbNewLine & _
    "</mso:customUI>"
    
    '----- Format to proper XML
    ribbonXML = Replace(ribbonXML, "'", """")
    
    '----- Inject the custom ribbon
    Open path & fileName For Output Access Write As hFile
    Print #hFile, ribbonXML
    Close hFile
    
    '----- Debug Display
    If DEBUGMODE Then MsgBox ribbonXML
    If DEBUGMODE Then Debug.Print ribbonXML

End Sub
'-----

'----- CLEAN UP AFTER OURSELVES
Sub ClearFringeUI(Optional DEBUGMODE As Boolean)

    Dim hFile As Long: hFile = FreeFile
    Dim path As String: path = "C:\Users\" & Environ("Username") & "\AppData\Local\Microsoft\Office\"
    Dim fileName As String: fileName = "Excel.officeUI"
    Dim ribbonXML As String
    
    ribbonXML = Replace("<mso:customUI xmlns:mso='http://schemas.microsoft.com/office/2009/07/customui'>" & _
    "<mso:ribbon></mso:ribbon></mso:customUI>" _
    , "'", """")
    
    Open path & fileName For Output Access Write As hFile
    Print #hFile, ribbonXML
    Close hFile
    
    '----- Debug Display
    If DEBUGMODE Then MsgBox ribbonXML
    If DEBUGMODE Then Debug.Print ribbonXML

End Sub
'-----




