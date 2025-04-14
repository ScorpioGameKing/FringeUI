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
    Application.ScreenUpdating = False
    If RealityCheck.IsClassModuleLoaded("FringeUIMultiLoader") * RealityCheck.IsClassModuleLoaded("FringeUIManager") * RealityCheck.IsClassModuleLoaded("FringeUIPackage") Then
        ' Multiple UI Packages
        On Error Resume Next
        Set multiLoader = New FringeUIMultiLoader
        PPTLinker.Init "PPTLinking", "PPTLinking", ".pptx", multiLoader
        MyCoolMod.InitUI multiLoader
        multiLoader.BuildMultiUIPackage
        FringeUIReloader.SetUIPackage multiLoader.MultiUIPackage
        UIManager.BuildFringeUI multiLoader.MultiUIPackage, False
    ElseIf RealityCheck.IsClassModuleLoaded("FringeUIManager") * RealityCheck.IsClassModuleLoaded("FringeUIPackage") Then
        ' Single UI Package
        PPTLinker.Init "PPTLinking", "PPTLinking", ".pptx"
    Else
        ' Default Launch
    End If
    Application.ScreenUpdating = True
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Application.ScreenUpdating = False
    UIManager.ClearFringeUI
    Application.ScreenUpdating = True
End Sub

Private Sub Workbook_Deactivate()
    Application.ScreenUpdating = False
    UIManager.ClearFringeUI
    Application.ScreenUpdating = True
End Sub
