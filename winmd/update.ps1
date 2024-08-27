New-Item -Force -Type Directory .\out\ > $null

cmake -B .\out\ -S .\mapi-scrubbed\ -G Ninja
cmake --build .\out\ -j

& dotnet tool install ClangSharpPInvokeGenerator -g
& dotnet build

Copy-Item .\bin\Microsoft.Office.Outlook.MAPI.Win32.winmd ..\crates\update-bindings\winmd\
