---
title: FringeUIMultiLoader
date: 2025-04-14 4:28:30 -0500
categories: [Modules, Core]
tags: [documentation, class-module, module]
description: An overview of how the FringeUIMultiLoader works and should be used.
---

## FringeUIMultiLoader Overview

The FringeUIMultiLoader is used to merge multiple modules with custom FringeUI components.
It works first in first out so the subsequent modules will be appended to existing Tabs
and groups if they exist. To prepare the multiloader, multiple [FringeUIPackages](https://scorpiogameking.github.io/FringeUI/posts/FringeUIPackage/)
need to be passed to the MultiLoader using [AddUIPackage](#adduipackage). Once all UIPackages
have been passed to the MultiLoader you will need to call [BuildMultiUIPackage](#buildmultiuipackage)
and the MultiLoader will begin to collapse the UIPackages into a single valid UIPackage.

### Class_Initialize

> WIP
{: .prompt-info }

### AddUIPackage

> WIP
{: .prompt-info }

### RemoveUIPackage

> WIP
{: .prompt-info }

### BuildToolsTab

> WIP
{: .prompt-info }

### BuildMultiUIPackage

> WIP
{: .prompt-info }

### Link to Module

[FringeUIMultiLoader on Github](https://github.com/ScorpioGameKing/FringeUI/blob/main/fringeui/class_modules/FringeUI/FringeUIMultiLoader.cls)
