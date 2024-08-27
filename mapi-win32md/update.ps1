param($mapi_rs_path)

& dotnet build

Copy-Item .\bin\Microsoft.Office.Outlook.MAPI.Win32.winmd "$mapi_rs_path\crates\update-bindings\winmd\"
