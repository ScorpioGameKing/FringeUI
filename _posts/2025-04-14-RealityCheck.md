---
title: RealityCheck
date: 2025-04-14 4:29:30 -0500
categories: [Modules, External]
tags: [documentation, standard-module, module]
description: An overview of how the RealityCheck works and should be used.
---

## RealityCheck Overview

A Collection of helpful search Functions for checking modules, collections and
dynamic string subsections.

## IsClassModuleLoaded

Used to check if a workbook has a class module installed. Defaults to current active workbook.
```vb
Public Function IsClassModuleLoaded(name As String, Optional wb As Workbook) As Boolean
```

To Call:
```vb
moduleCheck = RealityCheck.IsClassModuleLoaded ("TargetModule")
```

Returns 
```vb
Boolean
```

## IsStandardModuleLoaded

Used to check if a workbook has a standard module installed. Defaults to current active workbook.
```vb
Public Function IsStandardModuleLoaded(name As String, Optional wb As Workbook) As Boolean
```

To Call:
```vb
moduleCheck = RealityCheck.IsStandardModuleLoaded ("TargetModule")
```

Returns 
```vb
Boolean
```

## IsAnyModuleLoaded

Used to check if a workbook has a module installed. Defaults to current active workbook.
```vb
Public Function IsAnyModuleLoaded(name As String, Optional wb As Workbook) As Boolean
```

To Call:
```vb
moduleCheck = RealityCheck.IsAnyModuleLoaded ("TargetModule")
```

Returns 
```vb
Boolean
```

## InCollection

Used to check if a collection contains a string key.
```vb
Public Function InCollection(col As Collection, key As String) As Boolean
```

To Call:
```vb
collectionCheck = RealityCheck.InCollection(myCollection, "myKey")
```

Returns:
```vb
Boolean
```

## ReturnBetweenElements

Used to find the substring between keys.

```vb
Public Function ReturnBetweenElements(sConCat As String, sFirstElement As String, sSecondElement As String) As String
```

To Call:
```vb
stringBetweenKeys = RealityCheck.ReturnBetweenElements("garbageHello World!data", "garbage", "data")
```

Returns:
```vb
"Hello World!"
```

## Link to Module

[RealityCheck on Github](https://github.com/ScorpioGameKing/FringeUI/blob/main/fringeui/modules/FringeUI/RealityCheck.bas)
