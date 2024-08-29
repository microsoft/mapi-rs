//! Define [`Logon`] and [`LogonFlags`].

use crate::{sys, Initialize};
use std::{iter, ptr, sync::Arc};
use windows::Win32::Foundation::*;
use windows_core::*;

/// Set of flags that can be passed to [`sys::MAPILogonEx`].
#[derive(Default)]
pub struct LogonFlags {
    /// Pass [`sys::MAPI_ALLOW_OTHERS`].
    pub allow_others: bool,

    /// Pass [`sys::MAPI_BG_SESSION`].
    pub bg_session: bool,

    /// Pass [`sys::MAPI_EXPLICIT_PROFILE`].
    pub explicit_profile: bool,

    /// Pass [`sys::MAPI_EXTENDED`].
    pub extended: bool,

    /// Pass [`sys::MAPI_FORCE_DOWNLOAD`].
    pub force_download: bool,

    /// Pass [`sys::MAPI_LOGON_UI`].
    pub logon_ui: bool,

    /// Pass [`sys::MAPI_NEW_SESSION`].
    pub new_session: bool,

    /// Pass [`sys::MAPI_NO_MAIL`].
    pub no_mail: bool,

    /// Pass [`sys::MAPI_NT_SERVICE`].
    pub nt_service: bool,

    /// Pass [`sys::MAPI_SERVICE_UI_ALWAYS`].
    pub service_ui_always: bool,

    /// Pass [`sys::MAPI_TIMEOUT_SHORT`].
    pub timeout_short: bool,

    /// Pass [`sys::MAPI_UNICODE`].
    pub unicode: bool,

    /// Pass [`sys::MAPI_USE_DEFAULT`].
    pub use_default: bool,
}

impl From<LogonFlags> for u32 {
    fn from(value: LogonFlags) -> Self {
        let allow_others = if value.allow_others {
            sys::MAPI_ALLOW_OTHERS
        } else {
            0
        };
        let bg_session = if value.bg_session {
            sys::MAPI_BG_SESSION
        } else {
            0
        };
        let explicit_profile = if value.explicit_profile {
            sys::MAPI_EXPLICIT_PROFILE
        } else {
            0
        };
        let extended = if value.extended {
            sys::MAPI_EXTENDED
        } else {
            0
        };
        let force_download = if value.force_download {
            sys::MAPI_FORCE_DOWNLOAD
        } else {
            0
        };
        let logon_ui = if value.logon_ui {
            sys::MAPI_LOGON_UI
        } else {
            0
        };
        let new_session = if value.new_session {
            sys::MAPI_NEW_SESSION
        } else {
            0
        };
        let no_mail = if value.no_mail { sys::MAPI_NO_MAIL } else { 0 };
        let nt_service = if value.nt_service {
            sys::MAPI_NT_SERVICE
        } else {
            0
        };
        let service_ui_always = if value.service_ui_always {
            sys::MAPI_SERVICE_UI_ALWAYS
        } else {
            0
        };
        let timeout_short = if value.timeout_short {
            sys::MAPI_TIMEOUT_SHORT
        } else {
            0
        };
        let unicode = if value.unicode { sys::MAPI_UNICODE } else { 0 };
        let use_default = if value.use_default {
            sys::MAPI_USE_DEFAULT
        } else {
            0
        };

        allow_others
            | bg_session
            | explicit_profile
            | extended
            | force_download
            | logon_ui
            | new_session
            | no_mail
            | nt_service
            | service_ui_always
            | timeout_short
            | unicode
            | use_default
    }
}

/// Call [`sys::MAPILogonEx`] and hold on to the [`sys::IMAPISession`].
///
/// This helper also holds onto an `Arc<Initialize>`, which ensures that there are balanced calls
/// to [`sys::MAPIInitialize`] and [`sys::MAPIUninitialize`] around every [`Logon`] object that
/// shares a reference to that instance of [`Initialize`].
pub struct Logon {
    /// Access the [`sys::IMAPISession`].
    pub session: sys::IMAPISession,

    _initialized: Arc<Initialize>,
}

impl Logon {
    pub fn new(
        initialized: Arc<Initialize>,
        ui_param: HWND,
        profile_name: Option<&str>,
        password: Option<&str>,
        flags: LogonFlags,
    ) -> Result<Self> {
        let mut profile_name: Option<Vec<_>> =
            profile_name.map(|value| value.bytes().chain(iter::once(0)).collect());
        let profile_name = profile_name
            .as_mut()
            .map(|value| value.as_mut_ptr())
            .unwrap_or(ptr::null_mut());
        let mut password: Option<Vec<_>> =
            password.map(|value| value.bytes().chain(iter::once(0)).collect());
        let password = password
            .as_mut()
            .map(|value| value.as_mut_ptr())
            .unwrap_or(ptr::null_mut());

        Ok(Self {
            _initialized: initialized,
            session: unsafe {
                let mut session = None;
                sys::MAPILogonEx(
                    ui_param.0 as usize,
                    profile_name as *mut _,
                    password as *mut _,
                    flags.into(),
                    ptr::from_mut(&mut session),
                )?;
                session
            }
            .ok_or_else(|| Error::from(E_FAIL))?,
        })
    }
}
