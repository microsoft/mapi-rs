// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

use core::{ptr, slice};
use outlook_mapi::{sys::*, *};
use windows_core::*;

fn main() -> Result<()> {
    println!("Initializing MAPI...");
    let initialized = Initialize::new(Default::default()).expect("failed to initialize MAPI");
    println!("Trying to logon to the default profile...");
    let logon = Logon::new(
        initialized,
        Default::default(),
        None,
        None,
        LogonFlags {
            extended: true,
            unicode: true,
            logon_ui: true,
            use_default: true,
            ..Default::default()
        },
    )
    .expect("should be able to logon to the default MAPI profile");
    println!("Success!");

    // Now try to list the stores in the default MAPI profile.
    SizedSPropTagArray! { PropTagArray[2] }
    let mut prop_tag_array = PropTagArray {
        aulPropTag: [PR_ENTRYID, PR_DISPLAY_NAME_W],
        ..Default::default()
    };
    SizedSSortOrderSet! { SortOrderSet[1] }
    let mut sort_order_set: SortOrderSet = SortOrderSet {
        aSort: [SSortOrder {
            ulPropTag: PR_DISPLAY_NAME_W,
            ulOrder: TABLE_SORT_ASCEND,
        }],
        ..Default::default()
    };
    let mut rows: RowSet = Default::default();
    unsafe {
        let stores_table = logon.session.GetMsgStoresTable(0)?;
        HrQueryAllRows(
            &stores_table,
            prop_tag_array.as_mut_ptr(),
            ptr::null_mut(),
            sort_order_set.as_mut_ptr(),
            50,
            rows.as_mut_ptr(),
        )?;
    }

    println!("Found {rows} stores", rows = rows.len());
    for (idx, row) in rows.into_iter().enumerate() {
        // Use 1-based indices for messages.
        let idx = idx + 1;

        assert_eq!(2, row.len());
        let mut values = row.iter();

        let Some(PropValue {
            tag: PropTag(PR_ENTRYID),
            value: PropValueData::Binary(entry_id),
        }) = values.next()
        else {
            eprintln!("Store {idx}: missing entry ID");
            continue;
        };

        let Some(PropValue {
            tag: PropTag(PR_DISPLAY_NAME_W),
            value: PropValueData::Unicode(display_name),
        }) = values.next()
        else {
            eprintln!("Store {idx}: missing display name");
            continue;
        };
        let display_name = unsafe { PCWSTR::from_raw(display_name.as_ptr()).to_string() }
            .unwrap_or_else(|err| format!("bad display name: {err}"));

        println!(
            "Store {idx}: {display_name} ({entry_id} byte ID)",
            entry_id = entry_id.len()
        );

        unsafe {
            let mut store = None;
            logon.session.OpenMsgStore(
                0,
                entry_id.len() as u32,
                entry_id.as_ptr() as *mut _,
                &<IMsgStore as Interface>::IID as *const _ as *mut _,
                MAPI_BEST_ACCESS | MAPI_DEFERRED_ERRORS | MDB_NO_DIALOG | MDB_NO_MAIL,
                &mut store,
            )?;
            let Some(store) = store else {
                eprintln!("OpenMsgStore succeeded but store is None");
                continue;
            };

            let mut names = [MAPINAMEID {
                lpguid: &PS_PUBLIC_STRINGS as *const _ as *mut _,
                ulKind: MNID_STRING,
                Kind: MAPINAMEID_0 {
                    lpwstrName: PWSTR(w!("Keywords").0 as *mut _),
                },
            }];
            let mut prop_ids: MAPIOutParam<SPropTagArray> = Default::default();
            store.GetIDsFromNames(
                names.len() as u32,
                &mut ((&mut names) as *mut _),
                0,
                prop_ids.as_mut_ptr() as *mut _,
            )?;
            let Some(prop_ids) = prop_ids.as_mut() else {
                eprintln!("GetIDsFromNames succeeded but prop_ids is None");
                continue;
            };
            let prop_ids = slice::from_raw_parts_mut(
                prop_ids.aulPropTag.as_mut_ptr(),
                prop_ids.cValues as usize,
            );
            for prop_tag in prop_ids.iter().map(|tag| PropTag(*tag)) {
                let prop_type: u32 = prop_tag.prop_type().into();
                if prop_type != PT_UNSPECIFIED {
                    eprintln!("Unexpected prop type");
                    continue;
                }
                println!("Found prop id: {}", prop_tag.prop_id());
            }
        }
    }

    Ok(())
}
