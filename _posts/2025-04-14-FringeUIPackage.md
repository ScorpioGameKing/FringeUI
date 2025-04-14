---
title: FringeUIPackage
date: 2025-04-14 4:29:30 -0500
categories: [Modules, Core]
tags: [documentation, class-module, module]
description: An overview of how the FringeUIPackage works and should be used.
---

## FringeUIPackage Overview

The FringeUIPackage is the actual skeleton of your Custom UI. To create a new UIPackage add the following
anywhere you can `Set`:

```vb
Dim myNewUIPackage As FringeUIPackage: Set myNewUIPackage = New FringeUIPackage
```

Once you have a UIPackage the first thing you'll need to do is create your Main Tab and Group. If your UIpackage
does not contain a Tab with a Group you'll not be able to add any components. To continue building the example out:

```vb
myNewUIPackage.AddTab "MyCoolTab", "My New Tab", "mso:TabFormat"
myNewUIPackage.AddGroup "MyCoolTab", "MyNewGroup", "My New Group", "true"
```

At this point your package is ready to add any components to. But don't let this convince you that we
can only have 1 Custom Tab and Group. You simply need 1 Tab and Group to have a component work, you can
add multiple tabs, groups and any number of components to each as needed. If you really want 20 tabs with
1 button each then you can make it.

## uiPackage
This is the actual structured Collection containing the TAGS and XML data for injection by [FringeUIManager](https://scorpiogameking.github.io/FringeUI/posts/FringeUIManager/).
It should not be manipulated directly and instead handled by the relavant subroutines.

## AddTabStruct

> WIP
{: .prompt-info }

## AddTab

> WIP
{: .prompt-info }

## AddGroupStruct

> WIP
{: .prompt-info }

## AddGroup

> WIP
{: .prompt-info }

## AddButtonXML

> WIP
{: .prompt-info }

## AddButton

> WIP
{: .prompt-info }

## Link to Module

[FringeUIPackage on Github](https://github.com/ScorpioGameKing/FringeUI/blob/main/fringeui/class_modules/FringeUI/FringeUIPackage.cls)
