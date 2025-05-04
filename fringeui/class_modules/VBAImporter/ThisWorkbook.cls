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

Sub TestLoad()
    'VBAImporter.LoadVBComp ActiveWorkbook, "C:\Users\scorpio\OneDrive\Documents\ImportExport\Exports\FringeUIManager.cls"
    'VBAImporter.LoadVBComp ActiveWorkbook, "C:\Users\scorpio\OneDrive\Documents\ImportExport\Exports\FringeUIMultiLoader.cls"
    'VBAImporter.LoadVBComp ActiveWorkbook, "C:\Users\scorpio\OneDrive\Documents\ImportExport\Exports\FringeUIPackage.cls"
    'VBAImporter.LoadVBComp ActiveWorkbook, "C:\Users\scorpio\OneDrive\Documents\ImportExport\Exports\FringeUIReloader.bas"
    'VBAImporter.LoadVBComp ActiveWorkbook, "C:\Users\scorpio\OneDrive\Documents\ImportExport\Exports\RealityCheck.bas"
    'VBAImporter.LoadVBComp ActiveWorkbook, "C:\Users\scorpio\OneDrive\Documents\ImportExport\Exports\Toaster.bas"
    ExportVBComps ActiveWorkbook, "C:\Users\scorpio\OneDrive\Documents\ImportExport\Exports"
End Sub

Private Sub Workbook_Activate()
    TestLoad
    If RealityCheck.IsClassModuleLoaded("FringeUIMultiLoader") * RealityCheck.IsClassModuleLoaded("FringeUIManager") * RealityCheck.IsClassModuleLoaded("FringeUIPackage") Then
        ' Multiple UI Packages
        Set multiLoader = New FringeUIMultiLoader
        VBAImporter.InitUI multiLoader
        
        On Error Resume Next
        multiLoader.BuildMultiUIPackage
        
        FringeUIReloader.SetUIPackage multiLoader.MultiUIPackage
        UIManager.BuildFringeUI multiLoader.MultiUIPackage, False
    ElseIf RealityCheck.IsClassModuleLoaded("FringeUIManager") * RealityCheck.IsClassModuleLoaded("FringeUIPackage") Then
        ' Single UI Package
        
        VBAImporter.InitUI

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
