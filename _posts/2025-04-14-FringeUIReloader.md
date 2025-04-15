---
title: FringeUIReloader
date: 2025-04-14 4:31:15 -0500
categories: [Modules, Core]
tags: [documentation, standard-module, module]
description: An overview of how the FringeUIReloader works and should be used.
---

## FringeUIReloader Overview

The *Idea* of the FringeUIRekoader is to allow us to update the current [FringeUIPackage](https://scorpiogameking.github.io/FringeUI/posts/FringeUIPackage/)
or [FringeUIMultiLoader](https://scorpiogameking.github.io/FringeUI/posts/FringeUIMultiLoader/). 

> The implementation is currently slightly flawed and only works with
> the MultiLoader branch.
{: .prompt-warning }

Adding to the existing custom UIPackage is fairly simple but the current limitation is removing old components. Without a reference to the `IRibbionUI` we
can't directly invalidate the UI so the feasiblity is still unknown.

### ReLoadUI

> WIP
{: .prompt-info }

### Link to Module

[FringeUIReloader on Github](https://github.com/ScorpioGameKing/FringeUI/blob/main/fringeui/modules/FringeUI/FringeUIReloader.bas)
