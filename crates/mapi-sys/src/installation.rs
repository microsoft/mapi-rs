use std::path::PathBuf;

use windows_core::{PCWSTR, w};

use crate::load_mapi::{get_outlook_mapi_path, OUTLOOK_QUALIFIED_COMPONENTS};

pub enum Architecture {
    X64,
    X86,
}

pub enum InstallationState {
    Installed(Architecture, PathBuf),
    NotInstalled,
}

pub fn check_outlook_mapi_installation() -> InstallationState {
    const OUTLOOK_QUALIFIERS: [(Architecture, PCWSTR); 2] = [
        (Architecture::X64, w!("outlook.x64.exe")),
        (Architecture::X86, w!("outlook.exe")),
    ];

    for category in OUTLOOK_QUALIFIED_COMPONENTS {
        for (bitness, qualifier) in OUTLOOK_QUALIFIERS {
            if let Ok(path) = unsafe { get_outlook_mapi_path(category, qualifier) } {
                return InstallationState::Installed(bitness, path)
            }
        }
    }

    InstallationState::NotInstalled
}