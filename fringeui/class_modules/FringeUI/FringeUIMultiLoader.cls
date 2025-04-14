VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FringeUIMultiLoader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----- FRINGEUI: VBA Wasn't Meant For This
' Author: ScorpioGameKing
' Version: 0.2.0
Private ToolsPackage As New FringeUIPackage
Private UIPackages As New Collection
Public MultiUIPackage As New Collection

Private Sub Class_Initialize()
    BuildToolsTab
End Sub

Public Sub AddUIPackage(uiPackage As FringeUIPackage, id As String)
    If Not RealityCheck.InCollection(UIPackages, id) Then UIPackages.Add uiPackage, id
End Sub

Public Sub RemoveUIPackage(uiPackage As FringeUIPackage, id As String)
    If RealityCheck.InCollection(UIPackages, id) Then UIPackages.Remove uiPackage, id
End Sub

Private Sub BuildToolsTab()
    ToolsPackage.AddTab "FringeUIMultiLoaderToolsTab", "FringeUI Tools", "mso:TabFormat"
    ToolsPackage.AddGroup "FringeUIMultiLoaderToolsTab", "FringeUIMultiLoaderToolsGroup", "Multi Loader Tools", "true"
    ToolsPackage.AddButton "FringeUIMultiLoaderToolsTab", "FringeUIMultiLoaderToolsGroup", "FringeUIMultiLoaderReloader", "Re-Load FringeUI", "GroupPivotTableData", "FringeUIReLoader.ReLoadUI"
    AddUIPackage ToolsPackage, "FringeUIToolsPackage"
End Sub

Public Sub BuildMultiUIPackage()
    Dim temp_tab_id As String, temp_group_id As String, temp_comp_id As String

    For Each sub_pack_tab In UIPackages
        Dim multiTab As New Collection: Set multiTab = New Collection
        Dim multiContents As New Collection: Set multiContents = New Collection
        temp_tab_id = RealityCheck.ReturnBetweenElements(sub_pack_tab.uiPackage.Item(1).Item(2), "id=", " ")
        
        If Not RealityCheck.InCollection(MultiUIPackage, temp_tab_id) Then
            multiTab.Add "TAB": multiTab.Add sub_pack_tab.uiPackage.Item(1).Item(2): multiTab.Add multiContents, "CONTENTS"
            MultiUIPackage.Add multiTab, temp_tab_id
        End If
        
        For Each sub_pack_group In sub_pack_tab.uiPackage.Item(1).Item("CONTENTS")
            Dim multiGroup As New Collection: Set multiGroup = New Collection
            Dim multiGroupContents As New Collection: Set multiGroupContents = New Collection
            temp_group_id = RealityCheck.ReturnBetweenElements(sub_pack_group.Item(2), "id=", " ")
            
            If Not RealityCheck.InCollection(MultiUIPackage.Item(temp_tab_id).Item("CONTENTS"), temp_group_id) Then
                multiGroup.Add "GROUP": multiGroup.Add sub_pack_group.Item(2): multiGroup.Add multiGroupContents, "CONTENTS"
                MultiUIPackage.Item(temp_tab_id).Item("CONTENTS").Add multiGroup, temp_group_id
            End If
            
            For Each sub_pack_component In sub_pack_group.Item("CONTENTS")
                Dim multiComponent As New Collection: Set multiComponent = New Collection
                Dim multiCompContents As New Collection: Set multiCompContents = New Collection
                temp_comp_id = RealityCheck.ReturnBetweenElements(sub_pack_component.Item(2), "id=", " ")
                temp_comp_type = RealityCheck.ReturnBetweenElements(sub_pack_component.Item(2), "<mso:", " ")
                
                If Not RealityCheck.InCollection(MultiUIPackage.Item(temp_tab_id).Item("CONTENTS").Item(temp_group_id).Item("CONTENTS"), temp_comp_id) Then
                    multiComponent.Add UCase(temp_comp_type): multiComponent.Add sub_pack_component.Item(2): multiComponent.Add multiCompContents, "CONTENTS"
                    MultiUIPackage.Item(temp_tab_id).Item("CONTENTS").Item(temp_group_id).Item("CONTENTS").Add multiComponent, temp_comp_id
                End If
            Next
        Next
    Next
End Sub

