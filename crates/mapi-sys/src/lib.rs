// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

//! This crate implements unsafe Rust bindings for the
//! [Outlook MAPI](https://learn.microsoft.com/en-us/office/client-developer/outlook/mapi/outlook-mapi-reference)
//! COM APIs using the [Windows](https://github.com/microsoft/windows-rs) crate.

use windows::Win32::{Foundation::*, System::LibraryLoader::*};

#[cfg(feature = "olmapi32")]
mod load_mapi;

fn get_mapi_module() -> HMODULE {
    use std::sync::OnceLock;
    use windows_core::*;

    static MAPI_MODULE: OnceLock<usize> = OnceLock::new();
    HMODULE(*MAPI_MODULE.get_or_init(|| unsafe {
        #[cfg(feature = "olmapi32")]
        if let Ok(module) = load_mapi::ensure_olmapi32() {
            return module.0 as usize;
        }

        LoadLibraryW(w!("mapi32"))
            .expect("mapi32 should be loaded on demand")
            .0 as usize
    }) as *mut _)
}

#[cfg(feature = "olmapi32")]
pub use load_mapi::ensure_olmapi32;

#[macro_use]
extern crate outlook_mapi_stub;

#[allow(non_snake_case)]
pub mod Microsoft;
