// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

//! Define [`Initialize`] and [`InitializeFlags`].

use crate::sys;
use core::ptr;
use std::sync::Arc;
use windows_core::*;

/// Set of flags that can be passed to [`sys::MAPIInitialize`] through the
/// [`sys::MAPIINIT::ulFlags`] member.
#[derive(Default)]
pub struct InitializeFlags {
    /// Pass [`sys::MAPI_MULTITHREAD_NOTIFICATIONS`].
    pub multithread_notifications: bool,

    /// Pass [`sys::MAPI_NT_SERVICE`].
    pub nt_service: bool,

    /// Pass [`sys::MAPI_NO_COINIT`].
    pub no_coinit: bool,
}

impl From<InitializeFlags> for u32 {
    fn from(value: InitializeFlags) -> Self {
        let multithread_notifications = if value.multithread_notifications {
            sys::MAPI_MULTITHREAD_NOTIFICATIONS
        } else {
            0
        };
        let nt_service = if value.nt_service {
            sys::MAPI_NT_SERVICE
        } else {
            0
        };
        let no_coinit = if value.no_coinit {
            sys::MAPI_NO_COINIT
        } else {
            0
        };

        multithread_notifications | nt_service | no_coinit
    }
}

/// Call [`sys::MAPIInitialize`] in the constructor, and balance it with a call to
/// [`sys::MAPIUninitialize`] in the destructor.
pub struct Initialize();

impl Initialize {
    /// Call [`sys::MAPIInitialize`] with the specified flags in [`InitializeFlags`].
    pub fn new(flags: InitializeFlags) -> Result<Arc<Self>> {
        unsafe {
            sys::MAPIInitialize(ptr::from_mut(&mut sys::MAPIINIT {
                ulVersion: sys::MAPI_INIT_VERSION,
                ulFlags: flags.into(),
            }) as *mut _)?;
        }

        Ok(Arc::new(Self()))
    }
}

impl Drop for Initialize {
    /// Call [`sys::MAPIUninitialize`].
    fn drop(&mut self) {
        unsafe {
            sys::MAPIUninitialize();
        }
    }
}
