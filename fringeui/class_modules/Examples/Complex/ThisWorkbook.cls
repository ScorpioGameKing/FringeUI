VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private UIManager As New FringeUIManager
Private MyCustomUI As New FringeUIPackage
Private multiLoader As Object

Private Sub Workbook_Activate()
    If RealityCheck.IsClassModuleLoaded("FringeUIMultiLoader") * RealityCheck.IsClassModuleLoaded("FringeUIManager") * RealityCheck.IsClassModuleLoaded("FringeUIPackage") Then
        ' Multiple UI Packages
        On Error Resume Next
        multiLoader.BuildMultiUIPackage
        FringeUIReloader.SetUIPackage multiLoader.MultiUIPackage
        UIManager.BuildFringeUI multiLoader.MultiUIPackage, False
    ElseIf RealityCheck.IsClassModuleLoaded("FringeUIManager") * RealityCheck.IsClassModuleLoaded("FringeUIPackage") Then
        ' Single UI Package
        PPTLinker.Init "TEMPLATE/PPTLinking", "Output/PPTLinking", ".pptx"
    Else
        ' Default Launch
    End If
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    UIManager.ClearFringeUI
End Sub

Private Sub Workbook_Deactivate()
    UIManager.ClearFringeUI
End Sub


