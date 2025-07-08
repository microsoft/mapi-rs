use std::path::PathBuf;

use crate::load_mapi::{OFFICE_QUALIFIERS, OUTLOOK_QUALIFIED_COMPONENTS, get_outlook_mapi_path};

#[derive(Debug, Clone, Copy, PartialEq)]
pub enum Architecture {
    X64,
    X86,
}

pub enum InstallationState {
    Installed(Architecture, PathBuf),
    NotInstalled,
}

pub fn check_outlook_mapi_installation() -> InstallationState {
    for category in OUTLOOK_QUALIFIED_COMPONENTS {
        for (arch, qualifier) in OFFICE_QUALIFIERS {
            if let Ok(path) = unsafe { get_outlook_mapi_path(category, qualifier) } {
                return InstallationState::Installed(arch, path);
            }
        }
    }

    InstallationState::NotInstalled
}
