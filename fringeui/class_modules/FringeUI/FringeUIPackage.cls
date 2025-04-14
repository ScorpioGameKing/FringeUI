VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FringeUIPackage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----- FRINGEUI: VBA Wasn't Meant For This
' Author: ScorpioGameKing
' Version: 0.3.0

'----- Package and Struct Vars
Public uiPackage As New Collection

'----- TAB STRUCT: An ID to search, An ID for build, the XML to inject, and a Collection to hold Groups
Private Sub AddTabStruct(id As String, label As String, insertBeforeQ As String)
    Dim tab_struct As New Collection
    Dim tab_contents As New Collection
    Dim header As String
    
    header = "<mso:tab id='" & id & _
    "' label='" & label & _
    "' insertBeforeQ='" & insertBeforeQ & "'>" & vbNewLine
    header = Replace(header, "'", """")
    
    tab_struct.Add "TAB": tab_struct.Add header: tab_struct.Add tab_contents, "CONTENTS"
    If Not RealityCheck.InCollection(uiPackage, id) Then
        uiPackage.Add tab_struct, id
    Else
        uiPackage.Remove (id)
        uiPackage.Add tab_struct, id
    End If
End Sub

Public Sub AddTab(id As String, label As String, insertBeforeQ As String)
    AddTabStruct id, label, insertBeforeQ
End Sub

'----- GROUP STRUCT: An ID to search, An ID for build, the XML to inject, and a Collection to hold Buttons, Menus, ETC
Private Sub AddGroupStruct(tId As String, id As String, label As String, autoScale As String)
    Dim group_struct As New Collection
    Dim group_contents As New Collection
    Dim header As String
    
    header = "<mso:group id='" & id & _
    "' label='" & label & _
    "' autoScale='" & autoScale & "'>" & vbNewLine
    header = Replace(header, "'", """")
    
    group_struct.Add "GROUP": group_struct.Add header: group_struct.Add group_contents, "CONTENTS"
    If Not RealityCheck.InCollection(uiPackage(tId).Item("CONTENTS"), id) Then
        uiPackage(tId).Item("CONTENTS").Add group_struct, id
    Else
        uiPackage(tId).Item("CONTENTS").Remove (id)
        uiPackage(tId).Item("CONTENTS").Add group_struct, id
    End If
End Sub

Public Sub AddGroup(tId As String, id As String, label As String, autoScale As String)
    AddGroupStruct tId, id, label, autoScale
End Sub

'----- BUTTON STRUCT: A TAG for build, XML to Inject , Tab ID to sort
Private Sub AddButtonXML(tId As String, gId As String, id As String, label As String, image As String, callback As String)
    Dim button_struct As Collection: Set button_struct = New Collection
    
    bt = "<mso:button id='" & id & _
    "' label='" & label & _
    "' imageMso='" & image & _
    "' onAction='" & callback & "'/>" & vbNewLine
    bt = Replace(bt, "'", """")
    
    button_struct.Add "BUTTON": button_struct.Add bt
    If Not RealityCheck.InCollection(uiPackage.Item(tId).Item("CONTENTS").Item(gId).Item("CONTENTS"), id) Then
        uiPackage.Item(tId).Item("CONTENTS").Item(gId).Item("CONTENTS").Add button_struct, id
    Else
        uiPackage.Item(tId).Item("CONTENTS").Item(gId).Item("CONTENTS").Remove (id)
        uiPackage.Item(tId).Item("CONTENTS").Item(gId).Item("CONTENTS").Add button_struct, id
    End If
End Sub

Public Sub AddButton(tId As String, gId As String, id As String, label As String, image As String, callback As String)
    AddButtonXML tId, gId, id, label, image, callback
End Sub



