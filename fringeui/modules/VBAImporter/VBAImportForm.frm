VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} VBAImportForm 
   Caption         =   "VBA Import Options"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "VBAImportForm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "VBAImportForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ImportButton_Click()
    Dim path As String: path = VBAImportForm.FilePathInput.Text
    Dim Name As String: Name = VBAImportForm.ModuleNameCombo.Value
    
    If path = "" Then
        MsgBox "Invalid Path. Please re-enter the Path", vbOKOnly, "INVALID PATH"
        Exit Sub
    Else
        If Right(path, 1) <> "\" Then path = path & "\"
        If Name = "ALL" Or Name = "" Then
            For i = 1 To VBAImportForm.ModuleNameCombo.ListCount - 1
                VBAImporter.LoadVBComp ActiveWorkbook, path & VBAImportForm.ModuleNameCombo.List(i)
            Next
        Else
            VBAImporter.LoadVBComp ActiveWorkbook, path & Name
        End If
    End If
    MsgBox "VBA Has Successfully Been Imported!", vbOKOnly, "IMPORT SUCCESS"
    Unload Me
End Sub

Private Sub ModuleComboReload_Click()
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim folder As Object: Set folder = fso.GetFolder(VBAImportForm.FilePathInput.Text)
    Dim file As Object
    
    VBAImportForm.ModuleNameCombo.AddItem "ALL"
    For Each file In folder.Files
        VBAImportForm.ModuleNameCombo.AddItem file.Name
    Next
End Sub
