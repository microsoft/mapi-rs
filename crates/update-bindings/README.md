# update-bindings
This crate is a utility which will regenerate the bindings in [outlook-mapi-sys](https://crates.io/crates/outlook-mapi-sys).

## Windows Metadata
The Windows crate requires a Windows Metadata (`winmd`) file describing the API. The one used in this crate was generated with the [mapi-win32md](https://github.com/wravery/mapi-win32md) project.
