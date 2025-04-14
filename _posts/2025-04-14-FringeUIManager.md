---
title: FringeUIManager
date: 2025-04-14 4:24:45 -0500
categories: [Modules, Core]
tags: [documentation, class-module, module]
description: An overview of how the FringeUIManager works and should be used.
---

## FringeUIManager Overview

The FringeUIManager is the "Core" of FringeUI. It handles the final steps of injecting
a [FringeUIPackage](https://scorpiogameking.github.io/FringeUI/posts/FringeUIPackage/) 
into the RibbonXML. It also handles the cleaning before exiting to restore the default
Ribbon.

To create and UI a FringeUIManager include the following parts  in your "*ThisWorkbook*" 
file:

```vb
Private UIManager As New FringeUIManager
Private multiLoader As Object

Private Sub Workbook_Activate()
    ' Multiple UI Packages
    If RealityCheck.IsClassModuleLoaded("FringeUIMultiLoader") * RealityCheck.IsClassModuleLoaded("FringeUIManager") * RealityCheck.IsClassModuleLoaded("FringeUIPackage") Then
        On Error Resume Next
        ' Set multiLoader = New FringeUIMultiLoader
        MyCoolMod.InitUI multiLoader ' Your Module or UIPackage Building Here

        multiLoader.BuildMultiUIPackage
        FringeUIReloader.SetUIPackage multiLoader.MultiUIPackage
        UIManager.BuildFringeUI multiLoader.MultiUIPackage, False
    ' Single UI Package
    ElseIf RealityCheck.IsClassModuleLoaded("FringeUIManager") * RealityCheck.IsClassModuleLoaded("FringeUIPackage") Then
        MyCoolMod.InitUI ' Your Module or UIPackage Building Here
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
```

We begin by declaring our FringeUIManager and a stand-in for a [FringeUIMultiLoader](https://scorpiogameking.github.io/FringeUI/posts/FringeUIMultiLoader/).
When the workbook is activated we want to check for the other [Core FringeUI Modules](https://scorpiogameking.github.io/FringeUI/categories/core/).
If the MultiLoader is found then we need to add our UIPackage to the MultiLoader. MultiLoader installed regardless, you will always end of by passing a UIPackage to
[BuildFringeUI](#buildfringeui), in this example `MyCoolMod.InitUI` passes the module's
UIPacakge to either a UIManager or the MultiLoader if provided.

To handle the cleaning your workbook should also include calls to [ClearFringeUI](#clearfringeui)
in either `Workbook_BeforeClose` and or `Workbook_Deactivate`.

## BuildFringeUI

> WIP
{: .prompt-info }

## ClearFringeUI

> WIP
{: .prompt-info }+

## Link to Module

[FringeUIManager on Github](https://github.com/ScorpioGameKing/FringeUI/blob/main/fringeui/class_modules/FringeUI/FringeUIManager.cls)