Excel Add-in for handling multiple types of templates

### [adding the Add-In to Excel](#add-in)
disclaimer:
> this has only been tested on my windows computer\
> and may be different another windows machine and will\
> definitely would be different for a mac machine

1. Open "Registry Editor".
1. At the top, paste `Computer\HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Developer`.
1. Click `Edit > New > String Value` and give it a name.
1. Double click on the name and set the value to the absolute path of the `taskPaneManifest.xml`
1. For example: `C:\Users\{username}\Downloads\ExcelTamplateHandler/taskPaneManifest.xml`
1. Next time you open Excel at the top there will be a "developer Add-In" you can use.

File > Options > Trust Center > Trust Center Settings > Trusted Add-in Catalogs > 