# Create a Universal Windows Platform (UWP) app

## Prerequisites

Before you start this demo, you should have [Visual Studio](https://visualstudio.microsoft.com/vs/) installed on a computer running Windows 10 with [Developer mode turned on](https://docs.microsoft.com/windows/uwp/get-started/enable-your-device-for-development). If you do not have Visual Studio, visit the previous link for download options.

> **Note:** This tutorial was written with Visual Studio 2017 version 15.8.1. The steps in this guide may work with other versions, but that has not been tested.

Open Visual Studio, and select **File > New > Project**. In the **New Project** dialog, do the following:

1. Select **Templates > Visual C# > Windows Universal**.
1. Select **Blank App (Universal Windows)**.
1. Enter **graph-tutorial** for the Name of the project.

![Visual Studio 2017 create new project dialog](/Images/vs-newproj-01.png)

> **Note:** Ensure that you enter the exact same name for the Visual Studio Project that is specified in these lab instructions. The Visual Studio Project name becomes part of the namespace in the code. The code inside these instructions depends on the namespace matching the Visual Studio Project name specified in these instructions. If you use a different project name the code will not compile unless you adjust all the namespaces to match the Visual Studio Project name you enter when you create the project.

Select **OK**. In the **New Universal Windows Platform Project** dialog, ensure that the **Minimum version** is set to `Windows 10 Creators Update (10.0; Build 15063)` or later and select **OK**.