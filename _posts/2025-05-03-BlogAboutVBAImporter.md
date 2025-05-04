---
title: VBAImporter - The VBA Editor causes me pain
date: 2025-05-03 8:40:30 -0500
categories: [Modules, External]
tags: [blog, documentation, standard-module, class-module, module]
description: I dislike the VBA Editor and it's file handling so I made a FringeUI File Handler.
---

## The VBA Editor Feels Restrictive

I am personally not a major fan of the VBA Editor provided by Excel. I won't outright call it 
*bad* but there's many minor things that feel like they are *weighing* me down. And let's be
honest, at least from my understanding while learning, the VBA Editor is not anywhere close
to a priority for Microsoft. 

# The UI of Pain

![VBA UI](https://scorpiogameking.github.io/FringeUI/git_assets/images/UIofPain.png)

This UI on it's own cause me such a headache. It's so close to feeling decent for me but with
limitations such as **Not allowing multiple windows** it starts to feel like a box. A very
small and kind of unplesant box. The ability to modify color themes is limited in both scope 
and ability, I found what I feel is an alright theme but sadly the majority of the UI remains
as plesant as a flashbang. 

The ability to actually do multi-file work stops being feasible at 2 files. And this leaves 
me personally torn. A lot of VBA code can be made *multi-line* by including the `_` at the end
(?) of a line. Now do I try to follow the Nameing and `_` patterns I see, the general 
*verbosity* is a double edged sword to me, or do I keep to terse names and aim for line 
reduction? 

And every time I go to refactor one way or another, the Editor will gladly remind me with 
every line:

![VBA Error](https://scorpiogameking.github.io/FringeUI/git_assets/images/VBAError.png)

*I know* 

Now most of this is personal and I reconigize that. If you love the VBA Editor then I'm not 
here to stop you from using it. I'm simply explaining why *I* felt like there was a problem to
solve. You may have been saying, "Hey you can just Import and Export with the Editor 
already!" and yeah, I don't like how it works. Often times I'm working across several modules 
and if I was to use the default File Handling, I need to import and export every file 
indivually. A minor time sink but one none the less. So that brings us to VBAImporter.

## VBAImporter

Right of the bat, I should refactor this name as it's become more of a file handler and I plan
to extend it in that way if possible. Moving past this, what is it? VBAImporter is a module 
for single or bulk, class and standard module importing and exporting. For Ease of Use it is
intended for use with FringUI by providing a simple and easy to use interface. Let's break it
down.

# LoadVBComp

This is currently the only Import method, it handles loading single VBA files into the 
project. 

```vb
Sub LoadVBComp(ByVal wb As Workbook, ByVal path As String)
    If Dir(path) <> "" Then
```

We start of by asking for a `Workbook` and a `Path`. This path needs to be a `FullPath`, from
the Drive to the Extension. From there we check the Path right away, curently we just check to
see if it's at least something. 

```vb
        Dim m As VBComponent: Dim n As String
        n = Split(StrReverse(Split(StrReverse(path), "\")(0)), ".")(0)
```
If we pass through, we'll initialize 2 variables, `m` which we'll use to iterate through the 
project `modules` and `n`; `n` is *fancy* string spliting relying on the nature of a 
`FullPath`. 

```vb

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
```

We begin to interate through components of the given `Workbook` and if the `module` name is 
found in the project we jump to the FOUND label and we remove the module. Once removed we 
jump back to the LOAD label and continue as normal. If the module name is not found in the 
project the execution will naturally fall to the LOAD label code. Either way, the given file 
is imported into given `workbook` and we clean up a little bit on the way out. 

# ExportVBComps

The only exporter currently, it dumps all valid modules, standard and class, into a given 
folder path. 

```vb
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
```

Incredibly simple in nature and execution. The safety is mostly non existant but we do attempt
to ensure the `Path` given is to the inside of a folder. From there we just export the files 
with the proper extensions.

# DeleteVBComp

We saw this earlier and is another simple sub. It tries to remove the given module from the
project.

```vb
Sub DeleteVBComp(ByVal wb As Workbook, ByVal m As String)
    Application.DisplayAlerts = False
    On Error Resume Next
    wb.VBProject.VBComponents.Remove wb.VBProject.VBComponents(m)
    On Error GoTo 0
    Application.DisplayAlerts = True
End Sub
```

## FringeUI Additions

With everything given above you can easily implement and use this module in code but included
is a set of FringeUI Compatible Buttons. By Default they will show up in the FringeUI Tools 
Tab and will be treated as a pack in tool.

# UILoadVBComp

A FringeUI Button Callback that allows the user to load or update the `ActiveWorkbook` project
modules

```vb
Sub UILoadVBComp()
    path = InputBox("Please Provide A File Path", "File Import", "")
    LoadVBComp ActiveWorkbook, path
    Toaster.Toast "VBA Module Has Successfuly Been Imported", 1, "Success", 4096
End Sub
```

![VBA Import](https://scorpiogameking.github.io/FringeUI/git_assets/images/VBImportBut.png)

# UIExportVBComp

A FringeUI Button Callback that allows the user to dump all of the `ActiveWorkbook` project
modules

```vb
Sub UIExportVBComp()
    path = InputBox("Please Provide An Export Path", "File Export", "")
    ExportVBComps ActiveWorkbook, path
    Toaster.Toast "VBA Modules Have Successfuly Been Exported", 1, "Success", 4096
End Sub
```

![VBA Export](https://scorpiogameking.github.io/FringeUI/git_assets/images/VBExportBut.png)

# InitUI

The Standard FringeUIPackage Building Sub

```vb
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
```

![VBA Importer Group](https://scorpiogameking.github.io/FringeUI/git_assets/images/VBImportTools.png)

## What do we gain?

So was this worth it? Personally I feel by making the process of getting files in and out
of the `Workbook` will actually let development in more *modern* editors like VSCode or
even Notepad++ far more feasible. I plan on working on implementing a Sub for installing from
a folder and potentially implmenting a default folder structure. Thinking of VBAImporter as
a proper File Handler has changed the scope, hopefully allowing for these broader concepts to 
develop into powerful features.

I'll spend some time over the next few days hopefully creating the proper documentation note
posts for further reference but I wanted to start trying a more personable approach by 
actually talking through some of the development. Thanks for your time, hopefully we both
learned something with this.

> "It's All Right! You Can Stop Now!" - Sora, Kingdom Hearts 3
