use std::path::{Path, PathBuf};
use windows::Win32::Storage::FileSystem::GetBinaryTypeW;
use windows_core::{PCWSTR, w};

use crate::load_mapi::{
    OFFICE_QUALIFIERS, OUTLOOK_QUALIFIED_COMPONENTS, get_office_executable_path,
    get_office_mapi_path_no_install, get_outlook_mapi_path,
};

#[derive(Debug, Clone, Copy, PartialEq)]
pub enum Architecture {
    X64,
    X86,
}

/// Represents the state of MAPI installation detection
///
/// The third boolean parameter in `Installed` indicates whether this is an
/// official Outlook installation (true) or a fallback Office installation (false).
/// Fallback installations are experimental and not officially supported.
pub enum InstallationState {
    /// MAPI installation found: (Architecture, DLL Path, Is Outlook Installation)
    Installed {
        /// Platform architecture that is installed.
        architecture: Architecture,
        /// Path to `olmapi32.dll`.
        dll_path: PathBuf,
        /// Indicates whether this is an official Outlook installation (true) or a fallback Office installation (false).
        /// Fallback installations are experimental and not officially supported.
        is_outlook_installed: bool,
    },
    NotInstalled,
}

fn get_binary_architecture(file_path: &Path) -> Result<Architecture, Box<dyn std::error::Error>> {
    let path_str = file_path.to_string_lossy();
    let path_wide: Vec<u16> = path_str.encode_utf16().chain(std::iter::once(0)).collect();
    let mut binary_type: u32 = 0;

    unsafe {
        GetBinaryTypeW(PCWSTR::from_raw(path_wide.as_ptr()), &mut binary_type)?;

        match binary_type {
            0 => Ok(Architecture::X86), // SCS_32BIT_BINARY
            6 => Ok(Architecture::X64), // SCS_64BIT_BINARY
            _ => Err(format!("Unsupported binary type: {binary_type}").into()),
        }
    }
}

fn try_office_installation(category: PCWSTR, qualifier: PCWSTR) -> Option<InstallationState> {
    // EXPERIMENTAL: Try to find MAPI through non-Outlook Office applications
    // This is not officially supported and may not work reliably across all configurations.
    // We are working on a better long-term approach for comprehensive MAPI detection.

    // Try to get the executable path for architecture detection
    let exe_path = unsafe { get_office_executable_path(category, qualifier) }.ok()?;

    // Detect architecture from the executable
    let actual_arch = get_binary_architecture(&exe_path).ok()?;

    // Get the corresponding MAPI DLL path
    let dll_path = unsafe { get_office_mapi_path_no_install(category, qualifier) }.ok()?;

    Some(InstallationState::Installed {
        architecture: actual_arch,
        dll_path,
        is_outlook_installed: false,
    })
}

pub fn check_outlook_mapi_installation() -> InstallationState {
    const OUTLOOK_QUALIFIERS: [(Architecture, PCWSTR); 2] = [
        (Architecture::X64, w!("outlook.x64.exe")),
        (Architecture::X86, w!("outlook.exe")),
    ];

    // First, try the standard Outlook qualified components (officially supported)
    for category in OUTLOOK_QUALIFIED_COMPONENTS {
        for (bitness, qualifier) in OUTLOOK_QUALIFIERS {
            if let Ok(path) = unsafe { get_outlook_mapi_path(category, qualifier) } {
                return InstallationState::Installed {
                    architecture: bitness,
                    dll_path: path,
                    is_outlook_installed: true,
                };
            }
        }
    }

    // EXPERIMENTAL FALLBACK: If Outlook is not found, try other Office applications
    //
    // WARNING: This fallback method is NOT officially supported by Microsoft.
    // While technically functional, it relies on Office applications sharing MAPI
    // infrastructure, which may not be guaranteed in future versions.
    //
    // We are actively working on a more robust long-term solution for comprehensive
    // MAPI detection that will replace this experimental approach.
    for category in OUTLOOK_QUALIFIED_COMPONENTS {
        for qualifier in OFFICE_QUALIFIERS {
            if let Some(installation) = try_office_installation(category, qualifier) {
                return installation;
            }
        }
    }

    InstallationState::NotInstalled
}
