Attribute VB_Name = "PPTLinker"
Private ppt_app As PowerPoint.Application
Private ppt_presentation As PowerPoint.Presentation
Private template As String
Private file_name As String
Private file_path As String
Private init_run As Boolean

Private FringeUI As Object
Private uiPackage As Object

Public Sub Init(temp As String, file As String, ext As String, Optional multiLoader As Variant)
    If Not init_run Then
        Dim dsh As Worksheet: Set dsh = Nothing
        On Error Resume Next: Set dsh = ActiveWorkbook.Worksheets("PPT_LINKER_PRGDATA"): On Error GoTo 0
        If dsh Is Nothing Then: Worksheets.Add.name = "PPT_LINKER_PRGDATA"
        Set ppt_app = CreateObject("Powerpoint.Application")
        template = temp & ext
        SetFileName file, ext
        SetFilePath
        SavePRGDATA ext
        init_run = True
    End If
    If RealityCheck.IsClassModuleLoaded("FringeUIManager") * RealityCheck.IsClassModuleLoaded("FringeUIPackage") Then
        InitUI multiLoader
    End If
    Toaster.Toast "PPT Linker Intialized!", 1, "PPT Re-Init", 4096
End Sub

Public Sub reINIT()
    Application.ScreenUpdating = False
    If Not init_run Then
        If Sheets("PPT_LINKER_PRGDATA").Visible = xlSheetHidden Or Sheets("PPT_LINKER_PRGDATA").Visible = xlSheetVeryHidden Then Sheets("PPT_LINKER_PRGDATA").Visible = xlSheetVisible
        Set ppt_app = CreateObject("Powerpoint.Application")
        template = Sheets("PPT_LINKER_PRGDATA").Range("A1").Value
        SetFileName Sheets("PPT_LINKER_PRGDATA").Range("A2").Value, Sheets("PPT_LINKER_PRGDATA").Range("A3").Value
        SetFilePath
        Sheets("PPT_LINKER_PRGDATA").Visible = xlSheetVeryHidden
        init_run = True
    End If
    Application.ScreenUpdating = True
    Toaster.Toast "PPT Linker Re-Intialized!", 1, "PPT Re-Init", 4096
End Sub

Private Sub InitUI(Optional multiLoader As Variant)
    If FringeUI Is Nothing Then Set FringeUI = New FringeUIManager
    If uiPackage Is Nothing Then Set uiPackage = New FringeUIPackage
    
    uiPackage.AddTab "PPTLinker", "PPT Linker", "mso:TabFormat"
    uiPackage.AddGroup "PPTLinker", "PPTLinkerGroup", "PowerPoint Linking", "true"
    uiPackage.AddButton "PPTLinker", "PPTLinkerGroup", "PPTReInit", "Re-Intialize PPT Linker", "Repeat", "PPTLinker.reINIT"
    uiPackage.AddButton "PPTLinker", "PPTLinkerGroup", "PPTOpenLinked", "Open Linked PowerPoint", "MicrosoftPowerPoint", "PPTLinker.OpenPPT"
    uiPackage.AddButton "PPTLinker", "PPTLinkerGroup", "PPTUpdateLinked", "Update Linked PowerPoint", "PictureInsertMenu", "PPTLinker.UpdatePPT"
    
    If IsMissing(multiLoader) Then
        FringeUIReloader.SetUIPackage uiPackage.uiPackage
        FringeUI.BuildFringeUI uiPackage.uiPackage, True
    Else
        multiLoader.AddUIPackage uiPackage, "PPTLinkerUI"
    End If
End Sub

Sub SetFileName(name As String, ext As String)
    file_name = name & " " & Replace(Date, "/", "-") & ext
End Sub

Sub SetFilePath()
    file_path = Application.ActiveWorkbook.path & "/"
End Sub

Sub OpenPPT()
    If ppt_app.Presentations.Count > 0 Then
        Set ppt_presentation = FindOpenPPT
    Else
        Set ppt_presentation = FindSavedPPT
    End If
End Sub

Sub UpdatePPT()
    OpenPPT
    ClearSlideOf 1, "PPHBoard"
    ClearSlideOf 2, "PPHBoard"
    ClearSlideOf 3, "PPHBoard"
    PasteOnSlide 1, "Leaderboard 1", "A1", "B6", "P"
    PasteOnSlide 2, "Leaderboard 2", "A1", "B6", "P"
    PasteOnSlide 3, "Leaderboard 3", "A1", "B6", "P"
    Toaster.Toast "Slides have been updated!", 1, "Slide Update", 4096
End Sub

Function FindOpenPPT() As PowerPoint.Presentation
    Dim open_file As Boolean: open_file = False
    For Each p In ppt_app.Presentations
        If p.FullName = file_path & file_name Then
            Set FindOpenPPT = p
            open_file = True
        End If
    Next
    If Not open_file Then Set FindOpenPPT = FindSavedPPT
End Function

Function FindSavedPPT() As PowerPoint.Presentation
    On Error GoTo NoFile
    Set FindSavedPPT = ppt_app.Presentations.Open(file_path & file_name)
    Exit Function
NoFile:
    Set FindSavedPPT = CreateTemplatedPPT
End Function

Function CreateTemplatedPPT() As PowerPoint.Presentation
    Dim temp_ppt As PowerPoint.Presentation: Set temp_ppt = ppt_app.Presentations.Open(file_path & template)
    temp_ppt.SaveAs (file_path & file_name)
    Set CreateTemplatedPPT = temp_ppt
End Function

Function FindCaptureArea(sheet As String, topLeft As String, innerLeft As String, bottomRight As String) As String
    FindCaptureArea = topLeft & ":" & bottomRight & Sheets(sheet).Range(innerLeft).End(xlDown).Row
End Function

Sub ClearSlideOf(index As Integer, name As String)
    For Each s In ppt_presentation.Slides(index).Shapes
        If s.name = name Then
            s.Delete
        End If
    Next
End Sub

Sub PasteOnSlide(index As Integer, sheet As String, topLeft As String, innerLeft As String, bottomRight As String)
    Dim cap_area As Range: Set cap_area = Sheets(sheet).Range(FindCaptureArea(sheet, topLeft, innerLeft, bottomRight))
    cap_area.CopyPicture xlScreen, xlBitmap
    With ppt_presentation.Slides(index).Shapes.PasteSpecial(ppPasteBitmap)
        .name = "PPHBoard"
    End With
End Sub

Sub SavePRGDATA(file_ext As String)
    If Sheets("PPT_LINKER_PRGDATA").Visible = xlSheetHidden Or Sheets("PPT_LINKER_PRGDATA").Visible = xlSheetVeryHidden Then Sheets("PPT_LINKER_PRGDATA").Visible = xlSheetVisible
    Sheets("PPT_LINKER_PRGDATA").Range("A1").Value = template
    Sheets("PPT_LINKER_PRGDATA").Range("A2").Value = file_name
    Sheets("PPT_LINKER_PRGDATA").Range("A3").Value = file_ext
    Sheets("PPT_LINKER_PRGDATA").Visible = xlSheetVeryHidden
End Sub
