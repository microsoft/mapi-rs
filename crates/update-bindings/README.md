# update-bindings

This crate is a utility which will regenerate the bindings in [outlook-mapi-sys](https://crates.io/crates/outlook-mapi-sys). This utility should only typically be run by maintainers of [microsoft/mapi-rs](https://github.com/microsoft/mapi-rs) when preparing a new version with updated dependencies, e.g.:

```cmd
> cargo run update-bindings
```

## Scrubbing the MAPI Headers

The [/crates/update-bindings/mapi-scrubbed](./mapi-scrubbed/) directory contains a C++/CMake project which will generate a `MAPIConstants.h` convenience header from the SDK in [microsoft/MAPIStubLibrary](https://github.com/microsoft/MAPIStubLibrary) that the [microsoft/win32metadata](https://github.com/microsoft/win32metadata) generator can handle. You don't need to be able to manually build it, `update-bindings` will do that for you. You do need to have MSVC and CMake installed and in your PATH environment variable, though.

To generate bindings for an update to [microsoft/MAPIStubLibrary](https://github.com/microsoft/MAPIStubLibrary), update the `GIT_TAG` in this block from [/crates/update-bindings/mapi-scrubbed/CMakeLists.txt](./mapi-scrubbed/CMakeLists.txt) with a new commit hash or release tag:

```cmake
# Fetch the https://github.com/microsoft/MAPIStubLibrary repository.
include(FetchContent)
FetchContent_Declare(MAPIStubLibrary
  GIT_REPOSITORY "https://github.com/microsoft/MAPIStubLibrary"
  GIT_TAG "18655afee37164ea62052a8dd451402b91bb7c37"
  GIT_PROGRESS TRUE)
FetchContent_GetProperties(MAPIStubLibrary)
```

The next time that you run `update-bindings`, it should automatically pull down the new version of [microsoft/MAPIStubLibrary](https://github.com/microsoft/MAPIStubLibrary) and ingest any newly defined constants.

## Generating `Microsoft.Office.Outlook.MAPI.Win32.winmd`

The [/crates/update-bindings/winmd](./winmd/) directory contains a [dotnet](https://learn.microsoft.com/en-us/dotnet/core/tools/dotnet) (.NET Core) [/crates/update-bindings/winmd/MAPIWin32Metadata.proj](./winmd/MAPIWin32Metadata.proj) project which uses the scrubbed headers to generate `Microsoft.Office.Outlook.MAPI.Win32.winmd` with [microsoft/win32metadata](https://github.com/microsoft/win32metadata). Again, `update-bindings` will automate building this project to generate the `winmd` file. You do need to have `dotnet` installed and in your PATH environment variable, though.

To maintain compatibility with new versions of [microsoft/windows-rs](https://github.com/microsoft/windows-rs), the [microsoft/win32metadata](https://github.com/microsoft/win32metadata) SDK version needs to match the version used to generate the `Windows.Win32.winmd` file in [microsoft/windows-rs:/crates/libs/bindgen/default](https://github.com/microsoft/windows-rs/blob/master/crates/libs/bindgen/default/readme.md#windowswin32winmd). For example, in `windows = "0.58"`, `Windows.Win32.winmd` uses version `61.0.15`:

> ## `Windows.Win32.winmd`
>
> - Source: <https://www.nuget.org/packages/Microsoft.Windows.SDK.Win32Metadata/>
> - Version: `61.0.15`

That version number is based on the SDK version used to generate the `winmd` file, and we should use the same version in our [/crates/update-bindings/winmd/MAPIWin32Metadata.proj](./winmd/MAPIWin32Metadata.proj) project, but with the major version changed to a minor version:

```xml
<Project Sdk="Microsoft.Windows.WinmdGenerator/0.61.15-preview">
```

You can double check which versions of the `Microsoft.Windows.WinmdGenerator` SDK are available at [https://www.nuget.org/packages/Microsoft.Windows.WinmdGenerator/](https://www.nuget.org/packages/Microsoft.Windows.WinmdGenerator/#versions-body-tab).

Please also update the `WinmdVersion` property to match as well, so it is clear with which version of `Windows.Win32.winmd` it is compatible:

```xml
<WinmdVersion>0.61.0.15</WinmdVersion>
```
