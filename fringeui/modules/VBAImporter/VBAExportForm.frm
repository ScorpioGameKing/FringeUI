VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} VBAExportForm 
   Caption         =   "VBA Export Options"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "VBAExportForm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "VBAExportForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ExportButton_Click()
    Dim path As String: path = VBAExportForm.FilePathInput.Text
    Dim Name As String: Name = VBAExportForm.ModuleNameCombo.Value
    
    If path = "" Then
        MsgBox "Invalid Path. Please re-enter the Path", vbOKOnly, "INVALID PATH"
        Exit Sub
    Else
        If Name = "ALL" Or Name = "" Then
            VBAImporter.ExportVBComps ActiveWorkbook, path
        Else
            VBAImporter.ExportVBCompByName ActiveWorkbook, path, Name
        End If
    End If
    MsgBox "VBA Has Successfully Been Exported!", vbOKOnly, "EXPORT SUCCESS"
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    Dim m As VBComponent
    VBAExportForm.ModuleNameCombo.AddItem "ALL"
    For Each m In ActiveWorkbook.VBProject.VBComponents
        If ((m.Type = vbext_ct_StdModule) _
        Or (m.Type = vbext_ct_ClassModule)) Then
        VBAExportForm.ModuleNameCombo.AddItem m.Name
        End If
    Next
End Sub
