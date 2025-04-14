Attribute VB_Name = "FringeUIReloader"
Private uiPackage As New Collection
Private UIManager As New FringeUIManager

Public Sub SetUIPackage(newUIPackage As Collection)
    Set uiPackage = newUIPackage
End Sub

Public Sub ReLoadUI()
    If RealityCheck.IsClassModuleLoaded("FringeUIMultiLoader") * RealityCheck.IsClassModuleLoaded("FringeUIManager") * RealityCheck.IsClassModuleLoaded("FringeUIPackage") Then
        ' Multiple UI Packages
        UIManager.BuildFringeUI uiPackage
    ElseIf RealityCheck.IsClassModuleLoaded("FringeUIManager") * RealityCheck.IsClassModuleLoaded("FringeUIPackage") Then
        ' Single UI Package
        UIManager.BuildFringeUI uiPackage
    Else
        ' No UIPackage
    End If
    Toaster.Toast "Re-Loaded All Custom UI Elements!", 1, "FringeUI Re-Loader", 4096
    ThisWorkbook.Activate
End Sub
