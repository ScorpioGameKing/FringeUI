Attribute VB_Name = "RealityCheck"
'----- Suite of Exists, Contains, Find, Match, etc style Functions

'----- Module Checkers
'----- REQUIRES: Microsoft Visual Basic For Applications Extensibility 5.3
Public Function IsClassModuleLoaded(name As String, Optional wb As Workbook) As Boolean
    Dim j As Long
    Dim vbcomp As VBComponent
    Dim modules As New Collection
    IsClassModuleLoaded = False

    '----- Check if values are set
    If wb Is Nothing Then: Set wb = ThisWorkbook

    '----- Collect Class Module Names
    For Each vbcomp In ThisWorkbook.VBProject.VBComponents
        If (vbcomp.Type = vbext_ct_ClassModule) Then: modules.Add vbcomp.name
    Next vbcomp

    '----- Compare for the target
    For j = 1 To modules.Count
        If (name = modules.Item(j)) Then: IsClassModuleLoaded = True
    Next j
    j = 0

    '----- Missing Jump
    If (IsClassModuleLoaded = False) Then
        Debug.Print ("CLASS MODULE: " & name & " is not installed please add")
    End If
End Function

Public Function IsStandardModuleLoaded(name As String, Optional wb As Workbook) As Boolean
    Dim j As Long
    Dim vbcomp As VBComponent
    Dim modules As New Collection
    IsStandardModuleLoaded = False

    '----- Check if values are set
    If wb Is Nothing Then: Set wb = ThisWorkbook: End If
    If (name = "") Then: GoTo errorname: End If

    '----- Collect Class Module Names
    For Each vbcomp In ThisWorkbook.VBProject.VBComponents
        If (vbcomp.Type = vbext_ct_StdModule) Then: modules.Add vbcomp.name: End If
    Next vbcomp

    '----- Compare for the target
    For j = 1 To modules.Count
        If (name = modules.Item(j)) Then: IsStandardModuleLoaded = True: End If
    Next j
    j = 0

    '----- Missing Module
    If (IsStandardModuleLoaded = False) Then
        Debug.Print ("STANDARD MODULE: " & name & " is not installed please add")
    End If
End Function

Public Function IsAnyModuleLoaded(name As String, Optional wb As Workbook) As Boolean
    Dim j As Long
    Dim vbcomp As VBComponent
    Dim modules As New Collection
    IsAnyModuleLoaded = False

    '----- Check if values are set
    If wb Is Nothing Then: Set wb = ThisWorkbook: End If
    If (name = "") Then: GoTo errorname: End If

    '----- Collect Class Module Names
    For Each vbcomp In ThisWorkbook.VBProject.VBComponents
        If (vbcomp.Type = vbext_ct_StdModule) Then: modules.Add vbcomp.name: End If
    Next vbcomp

    '----- Compare for the target
    For j = 1 To modules.Count
        If (name = modules.Item(j)) Then: IsAnyModuleLoaded = True: End If
    Next j
    j = 0

    '----- Missing Module
    If (IsAnyModuleLoaded = False) Then
        Debug.Print ("Any MODULE: " & name & " is not installed please add")
    End If
End Function

'----- Exists For Collections
Public Function InCollection(col As Collection, key As String) As Boolean
  Dim var As Variant: Set var = Nothing
  Dim errNumber As Long
  InCollection = False: Err.Clear
  On Error Resume Next: var = col.Item(key): errNumber = CLng(Err.Number): On Error GoTo 0
  'If errNumber = 5 Then: InCollection = False: Else: InCollection = True: End If '5 is not in, 0 and 438 represent incollection
  InCollection = IIf(errNumber = 5, False, True)
End Function

'----- Find substring from delimiter to delimiter
Public Function ReturnBetweenElements(sConCat As String, sFirstElement As String, sSecondElement As String) As String
    Dim arArray1 As Variant, arArray2 As Variant
    Dim sReturn As String, sElement1 As String, sElement2 As String

    arArray1 = VBA.Split(sConCat, sFirstElement)           'removes the first element and creates an array of before and after
    sElement1 = arArray1(1)                                'returns the string after the first element

    arArray2 = VBA.Split(sElement1, sSecondElement)        'removes the second element and create an array of before and after
    sElement2 = arArray2(0)                                'returns the string before the second element
    
    ReturnBetweenElements = Trim(sElement2)
End Function

