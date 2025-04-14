---
title: FringeUI First Post!
date: 2025-04-13 9:46:45 -0500
categories: [News, FringeUI]
tags: [readme, documentation]
description: The intial post of the FringeUI Documentation Site. This is just a mirror of the Github Repo
---

# FringeUI
A VBA based Solution for your Office Customization Needs

![Example Custom Ribbon with Default Re-Loader and User Made Hello World Buttons](https://scorpiogameking.github.io/FringeUI/git_assets/images/HelloWordExampleBanner.png)

## What is FringeUI?
I'm glad you asked. `FringeUI` is an incredibly Fringe use-case of VBA to decorate the "Fringe", 
aka the Ribbon, of your Office App Window with your own Custom Tabs, Groups, Buttons and more! 
Do you have that one Macro you use every day but it's a *total pain* to open the menu every 
time? Got that ugly **CommandButton** you threw on a blank sheet to help clean up stray data? No
More! With `FringeUI` you're a few simple lines away from having a stylish and clean button in
the Ribbon that works like any other. 

## Current Features
- Custom Tabs
    > Create, name and organize your own custom tabs
- Custom Groups
    > Group Components using custom groups to keep related features together
- Custom Buttons
    > Create Custom Buttons using built-in icons and your own custom callbacks

## Planned Features
- Custom Menus
    > When a Group is not enough, list Components together in a handy dropdown

### How to Install Modules
1. Find and click the "Developer" Tab in the Ribbon.


> If you are missing the "Developer" Tab go to File -> Options -> Customize Ribbon and Click the 
> checkbox next to "Developer"
{: .prompt-tip }

2. Find and click "Visual Basic", This will open another Window, "Microsoft Visual Basic for Applications"
3. In this Window, Find in the top left the "File" Dropdown and select "Import Module" (Ctrl + M)
4. Find and install each of the Required Modules Below as needed

## Install Process (Single UI)
### Required Modules
- Class Module List
    - FringeUIManager
    - FringeUIPackage
- Module List
    - RealityCheck
    - Toaster

## Install Process (MultiUILoader Module)
### Required Modules
- Class Module List
    - FringeUIManager
    - FringeUIMultiLoader
    - FringeUIPackage

- Module List
    - FringeUIReloader
    - RealityCheck
    - Toaster

## Make Your First Tab, Group and Button (Tutorial)
> WIP

## Extend Modules to Support MultiUILoader (Tutorial)
> WIP

## Module Overview

### FringeUIManager
`FringeUIManager` is the core UI Injection Class. `FringeUI` takes advantage of the fact the a User has a local instance of the OfficeUI file.
It finds and grabs this file and using a `FringeUIPackage` created through the `FringeUIPackage` Class it will inject prebuilt XML strings for each
component a User Defines. 

### FringeUIMultiLoader
`FringeUIMultiLoader` is the supporting framework for allowing indivdual User Modules to provide their own `FringeUIPackages` to be sorted and merged.
When the `FringeUIMultiLoader` is in use, instead of the User Module building it's `FringeUIPackage`, it will be sent to the `FringeUIMultiLoader` Packages. 
When all `FringeUIPackages` have been recived the `FringeUIMultiloader` needs to collapse all valid posiblities into a single package, always adding to 
existing Components. For example, If User Mod 1 defines a Tab, group and button while User Mod 2 defines the same tab but a different group and button it 
will take the group and button of Mod 2 and append them to Mod 1.

![Simplified Flowchart of the MultiLoader Build Process](https://scorpiogameking.github.io/FringeUI/git_assets/images/MultiLoaderFlowChartSimple.png)

### FringeUIPackage
> This will be refactored due potential to name confusion
{: .prompt-note }

`FringeUIPackage` is the `FringeUIPackage` Builder. All `FringeUIPackages` require a Tab and Group to be able add a component. To create and build a package 
simply call `YourUIPackage.AddComponentNameHere arg1, arg2, ...` to add your components and pass the `YourUIPackage.uiPackage` to either the `FringeUIMultiLoader` 
or `FringeUIManager`

### FringeUIReloader
Small Standard Module included to provide a default ReLoad Method for `FringeUIManager`'s and `FringeUIMultiLoader`'s default Tool Menu. 

### RealityCheck
Helpful Libray of Sanity Checks.

### Toaster
Simple Timed MsgBox Alternative
