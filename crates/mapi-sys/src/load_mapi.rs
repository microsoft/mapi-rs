// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

use std::{iter, path::PathBuf};
use windows::Win32::{
    Foundation::*,
    System::{ApplicationInstallationAndServicing::*, LibraryLoader::*},
};
use windows_core::*;

const OLMAPI32_MODULE: PCWSTR = w!("olmapi32.dll");

const O16_CATEGORY_GUID_CORE_OFFICE_RETAIL: PCWSTR = w!("{5812C571-53F0-4467-BEFA-0A4F47A9437C}");
const O15_CATEGORY_GUID_CORE_OFFICE_RETAIL: PCWSTR = w!("{E83B4360-C208-4325-9504-0D23003A74A5}");
const O14_CATEGORY_GUID_CORE_OFFICE_RETAIL: PCWSTR = w!("{1E77DE88-BCAB-4C37-B9E5-073AF52DFD7A}");
const O12_CATEGORY_GUID_CORE_OFFICE_RETAIL: PCWSTR = w!("{24AAE126-0911-478F-A019-07B875EB9996}");
const O11_CATEGORY_GUID_CORE_OFFICE_RETAIL: PCWSTR = w!("{BC174BAD-2F53-4855-A1D5-0D575C19B1EA}");
const O11_CATEGORY_GUID_CORE_OFFICE_DEBUG: PCWSTR = w!("{BC174BAD-2F53-4855-A1D5-1D575C19B1EA}");

pub const OUTLOOK_QUALIFIED_COMPONENTS: [PCWSTR; 6] = [
    O16_CATEGORY_GUID_CORE_OFFICE_RETAIL,
    O15_CATEGORY_GUID_CORE_OFFICE_RETAIL,
    O14_CATEGORY_GUID_CORE_OFFICE_RETAIL,
    O12_CATEGORY_GUID_CORE_OFFICE_RETAIL,
    O11_CATEGORY_GUID_CORE_OFFICE_RETAIL,
    O11_CATEGORY_GUID_CORE_OFFICE_DEBUG,
];

pub unsafe fn get_outlook_mapi_path(category: PCWSTR, qualifier: PCWSTR) -> Result<PathBuf> {
    let mut size = 0;
    if WIN32_ERROR(unsafe {
        MsiProvideQualifiedComponentW(
            category,
            qualifier,
            INSTALLMODE_DEFAULT,
            None,
            Some(&mut size),
        )
    }) != ERROR_SUCCESS
    {
        return Err(Error::from(E_INVALIDARG));
    }

    size += 1;
    let mut buffer = vec![0; size as usize];
    if WIN32_ERROR(unsafe {
        MsiProvideQualifiedComponentW(
            category,
            qualifier,
            INSTALLMODE_DEFAULT,
            Some(PWSTR::from_raw(buffer.as_mut_ptr())),
            Some(&mut size),
        )
    }) != ERROR_SUCCESS
        || size as usize != buffer.len() - 1
    {
        return Err(Error::from(E_INVALIDARG));
    }

    let mut path = PathBuf::from(String::from_utf16(&buffer[0..(buffer.len())])?);
    if !path.pop() {
        return Err(Error::from(E_INVALIDARG));
    }

    path.push("olmapi32.dll");
    Ok(path)
}

pub fn ensure_olmapi32() -> Result<HMODULE> {
    unsafe {
        // If olmapi32.dll is already loaded, we're done.
        let module = GetModuleHandleW(OLMAPI32_MODULE);
        if module.is_ok() {
            return module;
        }

        // Use our new installation detection to find Office/MAPI installations
        use crate::installation::{check_outlook_mapi_installation, InstallationState};
        
        match check_outlook_mapi_installation() {
            InstallationState::Installed(_, dll_path, _) => {
                let path_str = dll_path
                    .to_str()
                    .ok_or_else(|| Error::from(E_INVALIDARG))?;
                let buffer: Vec<_> = path_str
                    .encode_utf16()
                    .chain(iter::once(0))
                    .collect();
                return LoadLibraryW(PCWSTR::from_raw(buffer.as_ptr()));
            }
            InstallationState::NotInstalled => {
                // Fall back to legacy method for backward compatibility
                #[cfg(target_arch = "x86_64")]
                const QUALIFIER: PCWSTR = w!("outlook.x64.exe");
                #[cfg(not(target_arch = "x86_64"))]
                const QUALIFIER: PCWSTR = w!("outlook.exe");

                for category in OUTLOOK_QUALIFIED_COMPONENTS {
                    if let Ok(path) = get_outlook_mapi_path(category, QUALIFIER) {
                        let buffer: Vec<_> = path
                            .to_str()
                            .ok_or_else(|| Error::from(E_INVALIDARG))?
                            .encode_utf16()
                            .chain(iter::once(0))
                            .collect();
                        return LoadLibraryW(PCWSTR::from_raw(buffer.as_ptr()));
                    }
                }
            }
        }
    }

    Err(Error::from(E_NOTIMPL))
}
