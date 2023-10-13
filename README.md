# WebDS-PTCiptaAlamBI (AccurateOnline Web Automation)
This is a web automation project that uses Playwright Driver and C#.NET downoad several report of Accurate Online, a cloud-based accounting software. The project aims to automate the following scenarios:

- Login to Accurate Online with valid credentials
- Moving through several link to open report
- Downloading reports in Excel file format
  Logout from Accurate Online
- Zipping download files and log file then send to Fairbanc Cloud


## Prerequisites
To run this project, you need to have the following installed on your machine:

- Visual Studio
- .NET 7 SDK or later
- Playwright Library & WebDriver
- Suggested OS : Windows 10 or later or Windows Server 20-8 R2 or later

## Running project and publishing app
To run this project, follow these steps:

- Clone or download this repository to your local machine
- Open the solution file `WebDS-PTCiptaAlamBI.sln` in Visual Studio
- Restore the NuGet packages by right-clicking on the solution and selecting `Restore NuGet Packages` or Playwright for .NET as a NuGet package by running this command in the Package Manager Console: Install-Package Microsoft.Playwright


Then to publish tha app
- Build the solution
- Publish the solution by right-clicking on the solution and selecting Publish > Publish -to Folder
- Choose a folder where you want to publish the files and click on Publish
- Copy the published files to your target machine or server
- On the deployment folder open Windows power shell by  running this command in PowerShell: playwright.ps1 install
- Install C# version 7 in you server by running this command ' winget install Microsoft.DotNet.SDK.7 '.
- Run the .exe file .

## App Configuration
This is a documentation for the app.config file in theWebDS-PTCiptaAlamBI project. The app.config file contains settings that are used by the program.cs file to configure and run the web automation tests. The app.config file is an XML document that has the following structure:

```xml
<?xml version="1.0" encoding="utf-8"?>
<configuration>
  <appSettings>
    <add key="IncludeAccountingData" value="Y"/>
    <add key="BrowserIsVisible" value="N" />
    <add key="ConnectToSandbox" value="Y"/>
    <add key="LogFolderName" value="logs"/>
    <add key="ConsoleScreenLogMessage" value="Y"/>
    <add key="TraceElement" value="N" />
    <add key="DT-Id" value="###"/>
    <add key="DT-Name" value="x-x-x-x-x-x"/>
    <add key="LoginId" value="####"/>
    <add key="Password" value="****"/>
    <add key="AppDownloadFolder" value="webui01"/>
    <add key="UploadFolder" value="uploads"/>
    <add key="DataFolder" value="sharing"/>
  </appSettings>
</configuration>
```
The appSettings section contains a list of key-value pairs that define the settings for the web automation tests. Each setting has a description and an example value as follows:

- `IncludeAccountingData`: Whether to include accounting data or not. It can be either `Y` or `N`. The default value is `Y`.
- `BrowserIsVisible`: Whether to make the browser visible or not. It can be either `Y` or `N`. The default value is `N`.
- `ConnectToSandbox`: Whether to connect to the sandbox or not. It can be either `Y` or `N`. The default value is `Y`.
- `LogFolderName`: The name of the log folder where the test results and screenshots are stored. The default value is `logs`.
- `ConsoleScreenLogMessage`: Whether to show the console screen log message or not. It can be either `Y` or `N`. The default value is `Y`.
- `TraceElement`: Whether to trace the element or not. It can be either `Y` or `N`. The default value is `N`.
- `DT-Id`: The ID of the company to use for testing. The default value is `539`.
- `DT-Name`: The name of the company to use for testing. 
- `LoginId`: The login ID of the Accurate Online account to use for testing. 
- `Password`: The password of the Accurate Online account to use for testing. 
- `AppDownloadFolder`: The name of the app download folder where the downloaded files are stored. The default value is `webui01`.
- `UploadFolder`: The name of the upload folder where the files to be uploaded are stored. The default value is `uploads`.
- `DataFolder`: The name of the data folder where the data files are stored. The default value is `sharing`.

To change any of these settings, you can edit the app.config file and save it before running the program.cs file.

## Troubleshooting
If you encounter any issues while running this project, you can try the following solutions:

- Make sure that you have installed all the prerequisites and configured all the settings correctly
- Make sure that your WebDriver matches your browser version and is placed in a folder that is accessible by your system path
- Make sure that your internet connection is stable and that Accurate Online website is reachable
- Make sure that your Accurate Online account has sufficient permissions and credits to perform the test scenarios







