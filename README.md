# GraphHelper

## Summary
This project is a collection of extension methods that aim to provide simpler means of obtaining some of the most used Graph functionalities, mostly related to MailFolders

## Status
[![NuGet Badge](https://buildstats.info/nuget/GHelper)](https://www.nuget.org/packages/GHelper/)

## Extension methods for Microsoft Graph types:
- MailFolder
```Csharp
GetMailFolderByName()
GetMailFolderByID()
GetMailFolderID()
```
- IMailFolderMessagesCollectionPage
```Csharp
GetInboxEmails()
GetSpecificFolderEmails()
```
- Message
```Csharp
SendExceptionEmail()
```
- GraphServiceClient
```Csharp
GetGraphClient()
```

## Installation
**DISCLAIMER:** Plese note that this package is still under development and bugs may be present. If you spot a bug, please [open a new Issue](https://github.com/Banovvv/GHelper/issues/new)

You can install the NuGet library into your project using:

Package Manager:
```
Install-Package GHelper -Version 1.0.0
```

.NET CLI:
```
dotnet add package GHelper --version 1.0.0
```

## License
Copyright © 2022 Ivan Gechev.

This package has MIT license. Refer to the [LICENSE](https://github.com/Banovvv/GHelper/blob/e2ca6eb3f858a887a7141b9442957cf6c76aaf3f/LICENSE) for detailed information.

## Questions, comments or additions
If you have a feature request or bug report, [open a new Issue](https://github.com/Banovvv/GHelper/issues/new) or [send a Pull request](https://github.com/Banovvv/GHelper/pulls).
