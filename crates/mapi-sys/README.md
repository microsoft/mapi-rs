# outlook-mapi-sys
This crate implements unsafe Rust bindings for the [Outlook MAPI](https://learn.microsoft.com/en-us/office/client-developer/outlook/mapi/outlook-mapi-reference) COM APIs using the [Windows](https://github.com/microsoft/windows-rs) crate.

## Getting Started
This crate has a friendlier wrapper in [outlook-mapi](https://crates.io/crates/outlook-mapi).

## Windows Metadata
The Windows crate requires a Windows Metadata (`winmd`) file describing the API. The one used in this crate is automatically regenerated as part of running `update-bindings`:
```cmd
> cargo run update-bindings
```
