use std::path::PathBuf;

use windows::Win32::System::Registry::*;
use windows_core::{HSTRING, PCWSTR, w};

use crate::load_mapi::{OUTLOOK_QUALIFIED_COMPONENTS, get_outlook_mapi_path};

#[derive(Debug, Clone, Copy, PartialEq)]
pub enum Architecture {
    X64,
    X86,
}

#[derive(Debug, Clone, Copy, PartialEq)]
pub enum DetectionMethod {
    /// Found via Outlook-specific Windows Installer API (legacy method)
    OutlookInstaller,
    /// Found via Office ClickToRun registry detection
    OfficeClickToRun,
    /// Found via Office MSI registry detection  
    OfficeMsi,
}

pub enum InstallationState {
    Installed(Architecture, PathBuf, DetectionMethod),
    NotInstalled,
}

#[derive(Debug, Clone)]
pub struct OfficeInstallation {
    pub architecture: Architecture,
    pub version: String,
    pub install_path: PathBuf,
    pub mapi_dll_path: PathBuf,
    pub detection_method: DetectionMethod,
}

// Registry paths for Office installations
// Only checking Office 16.0+ as earlier versions don't have the MAPI support we need
const OFFICE_REGISTRY_PATHS: &[&str] = &[
    r"SOFTWARE\Microsoft\Office\ClickToRun\Configuration", // Modern Office 365/2016+
    r"SOFTWARE\Microsoft\Office\16.0\Common\InstallRoot",  // Traditional MSI Office 2016
];

pub fn check_outlook_mapi_installation() -> InstallationState {
    // First try the original Outlook-specific method for backward compatibility
    if let InstallationState::Installed(arch, path, detection_method) =
        check_outlook_installation_legacy()
    {
        return InstallationState::Installed(arch, path, detection_method);
    }

    // Then check for Office installations via registry
    if let Some(installation) = find_office_installations().into_iter().next() {
        return InstallationState::Installed(
            installation.architecture,
            installation.mapi_dll_path,
            installation.detection_method,
        );
    }

    InstallationState::NotInstalled
}

/// Legacy method using Windows Installer API to find Outlook installations
fn check_outlook_installation_legacy() -> InstallationState {
    const OUTLOOK_QUALIFIERS: [(Architecture, PCWSTR); 2] = [
        (Architecture::X64, w!("outlook.x64.exe")),
        (Architecture::X86, w!("outlook.exe")),
    ];

    for category in OUTLOOK_QUALIFIED_COMPONENTS {
        for (bitness, qualifier) in OUTLOOK_QUALIFIERS {
            if let Ok(path) = unsafe { get_outlook_mapi_path(category, qualifier) } {
                return InstallationState::Installed(
                    bitness,
                    path,
                    DetectionMethod::OutlookInstaller,
                );
            }
        }
    }

    InstallationState::NotInstalled
}

/// Find all Office installations that include MAPI support
pub fn find_office_installations() -> Vec<OfficeInstallation> {
    let mut installations = Vec::new();

    // Check both 64-bit and 32-bit registry views
    for &wow64_flag in &[KEY_WOW64_64KEY, KEY_WOW64_32KEY] {
        let arch = if wow64_flag == KEY_WOW64_64KEY {
            Architecture::X64
        } else {
            Architecture::X86
        };

        installations.extend(check_office_registry_paths(arch, wow64_flag));
    }

    // Remove duplicates and sort by version (newest first)
    installations.sort_by(|a, b| b.version.cmp(&a.version));
    installations.dedup_by(|a, b| a.install_path == b.install_path);

    installations
}

fn check_office_registry_paths(
    arch: Architecture,
    wow64_flag: REG_SAM_FLAGS,
) -> Vec<OfficeInstallation> {
    let mut installations = Vec::new();

    for &registry_path in OFFICE_REGISTRY_PATHS {
        if let Some(installation) = check_office_registry_path(registry_path, arch, wow64_flag) {
            installations.push(installation);
        }
    }

    installations
}

fn check_office_registry_path(
    registry_path: &str,
    arch: Architecture,
    wow64_flag: REG_SAM_FLAGS,
) -> Option<OfficeInstallation> {
    unsafe {
        let mut hkey = HKEY::default();
        let path_hstring = HSTRING::from(registry_path);

        if RegOpenKeyExW(
            HKEY_LOCAL_MACHINE,
            &path_hstring,
            Some(0),
            KEY_READ | wow64_flag,
            &mut hkey,
        )
        .is_ok()
        {
            // Try to get the install path - different key names for different registry paths
            let install_path = if registry_path.contains("ClickToRun") {
                // For ClickToRun installations, get InstallationPath and construct root\Office16 path
                if let Some(base_path) = read_registry_string(&hkey, w!("InstallationPath")) {
                    let mut office_path = PathBuf::from(base_path);
                    office_path.push("root");
                    office_path.push("Office16");
                    Some(office_path)
                } else {
                    None
                }
            } else {
                // For traditional MSI installations, use the Path key directly
                read_registry_string(&hkey, w!("Path")).map(PathBuf::from)
            };

            if let Some(install_path) = install_path {
                // Check for MAPI DLL in the installation
                let mapi_paths = [
                    install_path.join("olmapi32.dll"),
                    install_path.join("mapi32.dll"),
                ];

                for mapi_path in &mapi_paths {
                    if mapi_path.exists() {
                        let version = extract_version_from_path(registry_path);
                        let detection_method = if registry_path.contains("ClickToRun") {
                            DetectionMethod::OfficeClickToRun
                        } else {
                            DetectionMethod::OfficeMsi
                        };
                        let _ = RegCloseKey(hkey);
                        return Some(OfficeInstallation {
                            architecture: arch,
                            version,
                            install_path: install_path.clone(),
                            mapi_dll_path: mapi_path.clone(),
                            detection_method,
                        });
                    }
                }
            }

            let _ = RegCloseKey(hkey);
        }
    }

    None
}

fn read_registry_string(hkey: &HKEY, value_name: PCWSTR) -> Option<String> {
    unsafe {
        let mut buffer_size = 0u32;

        // Get the required buffer size
        if RegQueryValueExW(*hkey, value_name, None, None, None, Some(&mut buffer_size)).is_ok()
            && buffer_size > 0
        {
            let mut buffer = vec![0u16; (buffer_size / 2) as usize];
            let mut actual_size = buffer_size;

            if RegQueryValueExW(
                *hkey,
                value_name,
                None,
                None,
                Some(buffer.as_mut_ptr() as *mut u8),
                Some(&mut actual_size),
            )
            .is_ok()
            {
                // Remove null terminator and convert to String
                if let Some(null_pos) = buffer.iter().position(|&x| x == 0) {
                    buffer.truncate(null_pos);
                }
                return String::from_utf16(&buffer).ok();
            }
        }
    }

    None
}

fn extract_version_from_path(registry_path: &str) -> String {
    // Extract version info from registry path
    if registry_path.contains("ClickToRun") {
        "16.0-ClickToRun".to_string()
    } else if let Some(start) = registry_path.find(r"\Office\") {
        let version_start = start + 8; // Length of "\Office\"
        if let Some(end) = registry_path[version_start..].find('\\') {
            return registry_path[version_start..version_start + end].to_string();
        }
        "16.0-MSI".to_string()
    } else {
        "Unknown".to_string()
    }
}
