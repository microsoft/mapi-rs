// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

//! MAPI Library Loading and Office Detection
//!
//! This module provides functionality to detect and load MAPI (Messaging Application
//! Programming Interface) libraries from Microsoft Office installations.
//!
//! # Official Support
//!
//! Microsoft officially supports MAPI through proper Office installations that include
//! MAPI components. Modern Office installations can include MAPI without requiring Outlook.
//!
//! # Experimental Fallback Support
//!
//! This module also includes **experimental** fallback detection for Office
//! applications that may have MAPI components available. This functionality:
//!
//! - Is **NOT officially supported** by Microsoft
//! - May not work reliably across all Office configurations  
//! - May break in future Office updates without notice
//! - Should be considered a temporary workaround
//!
//! **This fallback approach is experimental while we develop a more robust long-term solution.**

use std::{iter, path::PathBuf};
use windows::Win32::{
    Foundation::*,
    System::{ApplicationInstallationAndServicing::*, LibraryLoader::*},
};
use windows_core::*;

const OLMAPI32_MODULE: PCWSTR = w!("olmapi32.dll");

// EXPERIMENTAL: Office application fallback qualifiers for MAPI detection
//
// WARNING: This fallback detection method is NOT officially supported by Microsoft.
// While Office applications may share MAPI infrastructure with Outlook, this behavior
// is not guaranteed and may change in future Office versions without notice.
//
// This experimental approach is provided as a temporary workaround for environments
// where standard MAPI detection fails but other Office applications are installed.
// We are actively working on a more robust long-term solution for MAPI detection.
pub const OFFICE_QUALIFIERS: [PCWSTR; 6] = [
    // Excel
    w!("excel.exe"),
    // Word
    w!("winword.exe"),
    // PowerPoint
    w!("powerpnt.exe"),
    // Access
    w!("msaccess.exe"),
    // OneNote
    w!("onenote.exe"),
    // Publisher
    w!("mspub.exe"),
];

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

// Get the path to the MAPI DLL for Outlook, with installation checks
pub unsafe fn get_outlook_mapi_path(category: PCWSTR, qualifier: PCWSTR) -> Result<PathBuf> {
    unsafe {
        get_office_component_path(
            category,
            qualifier,
            Some("olmapi32.dll"),
            INSTALLMODE_DEFAULT,
        )
    }
}

// Get the path to the MAPI DLL for Outlook, without installation checks
pub unsafe fn get_office_mapi_path_no_install(
    category: PCWSTR,
    qualifier: PCWSTR,
) -> Result<PathBuf> {
    unsafe {
        get_office_component_path(
            category,
            qualifier,
            Some("olmapi32.dll"),
            INSTALLMODE_EXISTING,
        )
    }
}

// Get the path to the Office executable (e.g., winword.exe)
pub unsafe fn get_office_executable_path(category: PCWSTR, qualifier: PCWSTR) -> Result<PathBuf> {
    unsafe { get_office_component_path(category, qualifier, None, INSTALLMODE_EXISTING) }
}

unsafe fn get_office_component_path(
    category: PCWSTR,
    qualifier: PCWSTR,
    component: Option<&str>,
    install_mode: INSTALLMODE,
) -> Result<PathBuf> {
    let mut size = 0;
    if WIN32_ERROR(unsafe {
        MsiProvideQualifiedComponentW(category, qualifier, install_mode, None, Some(&mut size))
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
            install_mode,
            Some(PWSTR::from_raw(buffer.as_mut_ptr())),
            Some(&mut size),
        )
    }) != ERROR_SUCCESS
        || size as usize != buffer.len() - 1
    {
        return Err(Error::from(E_INVALIDARG));
    }

    let mut path = PathBuf::from(String::from_utf16(&buffer[0..(buffer.len())])?);

    match component {
        Some(comp) => {
            // For components like olmapi32.dll, pop the executable and add the component
            if !path.pop() {
                return Err(Error::from(E_INVALIDARG));
            }
            path.push(comp);
        }
        None => {
            // For executables, return the path as-is
        }
    }

    Ok(path)
}

pub fn ensure_olmapi32() -> Result<HMODULE> {
    unsafe {
        // If olmapi32.dll is already loaded, we're done.
        let module = GetModuleHandleW(OLMAPI32_MODULE);
        if module.is_ok() {
            return module;
        }

        #[cfg(target_arch = "x86_64")]
        const QUALIFIER: PCWSTR = w!("outlook.x64.exe");
        #[cfg(not(target_arch = "x86_64"))]
        const QUALIFIER: PCWSTR = w!("outlook.exe");

        // First, try the standard Outlook qualified components
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

        // Try fallback Office app qualifiers (without installation)
        //
        // EXPERIMENTAL FALLBACK: Attempt to locate MAPI through other Office applications.
        // This is NOT officially supported.
        // We are working on a more robust long-term solution for comprehensive MAPI detection.
        // This behavior may break in future Office updates without notice.
        for category in OUTLOOK_QUALIFIED_COMPONENTS {
            for qualifier in OFFICE_QUALIFIERS {
                if let Ok(path) = get_office_mapi_path_no_install(category, qualifier) {
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

    Err(Error::from(E_NOTIMPL))
}
