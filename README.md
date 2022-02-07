# GraphHelper

## Summary
This project is a collection of extension methods that aim to provide simpler means of obtaining some of the most used Graph functionalities, mostly related to MailFolders

## Status
[![NuGet Badge](https://buildstats.info/nuget/GraphHelper)](https://www.nuget.org/packages/GraphHelper/1.1.0)

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
**DISCLAIMER:** Plese note that this package is still under development and bugs may be present. If you spot a bug, please [open a new Issue](https://github.com/Banovvv/GraphHelper/issues/new)

You can install the NuGet library into your project using:

Package Manager:
```
Install-Package GraphHelper -Version 1.1.0
```

.NET CLI:
```
dotnet add package GraphHelper --version 1.1.0
```

## License
Copyright © 2022 Ivan Gechev.

This package has MIT license. Refer to the [LICENSE](https://github.com/Banovvv/GraphHelper/blob/e2ca6eb3f858a887a7141b9442957cf6c76aaf3f/LICENSE) for detailed information.

## Questions, comments or additions
If you have a feature request or bug report, [open a new Issue](https://github.com/Banovvv/GraphHelper/issues/new) or [send a Pull request](https://github.com/Banovvv/GraphHelper/pulls).
