Attribute VB_Name = "VBAImporter"
Private FringeUI As Object
Private uiPackage As Object

Sub LoadVBComp(ByVal wb As Workbook, ByVal path As String)
    If Dir(path) <> "" Then
        Dim m As VBComponent: Dim n As String
        n = Split(StrReverse(Split(StrReverse(path), "\")(0)), ".")(0)

        For Each m In wb.VBProject.VBComponents
            If (((m.Type = vbext_ct_StdModule) _
                Or (m.Type = vbext_ct_ClassModule)) _
                And (m.name = n)) _
                Then GoTo FOUND
        Next m

LOAD:
        On Error GoTo DONE
        wb.VBProject.VBComponents.IMPORT path
        GoTo DONE

FOUND:
        DeleteVBComp wb, n
        GoTo LOAD
    End If

DONE:
    Set wb = Nothing
End Sub

Sub ExportVBComps(ByVal wb As Workbook, ByVal path As String)
    Dim comp As VBComponent

    If Right(path, 1) <> "\" Then path = path & "\"

    For Each comp In wb.VBProject.VBComponents
        If comp.Type = vbext_ct_StdModule Then
            wb.VBProject.VBComponents.Item(comp.name).Export (path & comp.name & ".bas")
        ElseIf comp.Type = vbext_ct_ClassModule Then
            wb.VBProject.VBComponents.Item(comp.name).Export (path & comp.name & ".cls")
        End If
    Next comp
End Sub

Sub DeleteVBComp(ByVal wb As Workbook, ByVal m As String)
    Application.DisplayAlerts = False
    On Error Resume Next
    wb.VBProject.VBComponents.Remove wb.VBProject.VBComponents(m)
    On Error GoTo 0
    Application.DisplayAlerts = True
End Sub

Sub UILoadVBComp()
    path = InputBox("Please Provide A File Path", "File Import", "")
    LoadVBComp ActiveWorkbook, path
    Toaster.Toast "VBA Module Has Successfuly Been Imported", 1, "Success", 4096
End Sub

Sub UIExportVBComp()
    path = InputBox("Please Provide An Export Path", "File Export", "")
    ExportVBComps ActiveWorkbook, path
    Toaster.Toast "VBA Modules Have Successfuly Been Exported", 1, "Success", 4096
End Sub

Sub InitUI(Optional multiLoader As Variant)
    If FringeUI Is Nothing Then Set FringeUI = New FringeUIManager
    If uiPackage Is Nothing Then Set uiPackage = New FringeUIPackage
    
    uiPackage.AddTab "FringeUIMultiLoaderToolsTab", "FringeUI Tools", "mso:TabFormat"
    uiPackage.AddGroup "FringeUIMultiLoaderToolsTab", "VBAImporterGroup", "VBA Import Export Tools", "true"
    uiPackage.AddButton "FringeUIMultiLoaderToolsTab", "VBAImporterGroup", "VBAUIImporter", "Import VBA File", "SaveAsQuery", "VBAImporter.UILoadVBComp"
    uiPackage.AddButton "FringeUIMultiLoaderToolsTab", "VBAImporterGroup", "VBAUIExporter", "Export VBA to Folder", "LoadFromQuery", "VBAImporter.UIExportVBComp"
        
    If IsMissing(multiLoader) Then
        FringeUIReloader.SetUIPackage uiPackage.uiPackage
        FringeUI.BuildFringeUI uiPackage.uiPackage, True
    Else
        multiLoader.AddUIPackage uiPackage, "VBAImporter"
    End If
End Sub

