// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

//! Public macros and `const` functions to support SizedXXX types originally from `MAPIDefs.h`.

#![allow(non_snake_case)]

use crate::sys;
use core::mem;

/// All of the SizedXXX structs are declared with 1 ([`sys::MAPI_DIM`]) element in accordance with
/// C/C++ syntax rules that say you can't declare a zero-length array. We need to deduct that
/// placeholder element from the size of the container and re-add the size of the variable-length
/// elements to get the total size of a SizedXXX struct in memory.
const fn size_of_container<Container, Element>(count: usize) -> usize
where
    Container: Sized,
    Element: Sized,
{
    let mapi_dim_element_size = mem::size_of::<Element>() * sys::MAPI_DIM as usize;
    let base_container_size = mem::size_of::<Container>() - mapi_dim_element_size;
    let elements_size = mem::size_of::<Element>() * count;
    base_container_size + elements_size
}

/// Get the size of a [`sys::ENTRYID`] struct with `count` bytes in [`sys::ENTRYID::ab`].
pub const fn CbNewENTRYID(count: usize) -> usize {
    size_of_container::<sys::ENTRYID, u8>(count)
}

/// Get the size of a [`sys::ENTRYID`] struct with `count` bytes in [`sys::ENTRYID::ab`]. Since
/// there is no count of bytes stored in a member of [`sys::ENTRYID`] itself, this is just an alias
/// for [`CbNewENTRYID`].
pub const fn CbENTRYID(count: usize) -> usize {
    CbNewENTRYID(count)
}

/// Declare a variable length struct with the same layout as [`sys::ENTRYID`] and implement casting
/// functions:
///
/// - `fn as_ptr(&self) -> *const sys::ENTRYID`
/// - `fn as_mut_ptr(&mut self) -> *mut sys::ENTRYID`.
///
/// ### Sample
/// ```
/// # use outlook_mapi::{sys, SizedENTRYID};
/// #
/// SizedENTRYID! { EntryId[12] }
///
/// let entry_id = EntryId {
///     abFlags: [0x0, 0x1, 0x2, 0x3],
///     ab: [0x4, 0x5, 0x6, 0x7, 0x8, 0x9, 0xa, 0xb, 0xc, 0xd, 0xe, 0xf],
/// };
///
/// let entry_id: *const sys::ENTRYID = entry_id.as_ptr();
/// ```
#[macro_export]
#[allow(non_snake_case)]
macro_rules! SizedENTRYID {
    ($name:ident [ $count:expr ]) => {
        #[repr(C)]
        #[allow(non_snake_case)]
        struct $name {
            pub abFlags: [u8; 4],
            pub ab: [u8; $count],
        }

        $crate::impl_sized_struct_casts!($name, $crate::sys::ENTRYID);
    };
}

/// Get the size of a [`sys::SPropTagArray`] struct with `count` entries in
/// [`sys::SPropTagArray::aulPropTag`].
pub const fn CbNewSPropTagArray(count: usize) -> usize {
    size_of_container::<sys::SPropTagArray, u32>(count)
}

/// Get the size of a [`sys::SPropTagArray`] struct with [`sys::SPropTagArray::cValues`] entries in
/// [`sys::SPropTagArray::aulPropTag`].
pub const fn CbSPropTagArray(prop_tag_array: &sys::SPropTagArray) -> usize {
    CbNewSPropTagArray(prop_tag_array.cValues as usize)
}

/// Declare a variable length struct with the same layout as [`sys::SPropTagArray`] and implement
/// casting functions:
///
/// - `fn as_ptr(&self) -> *const sys::SPropTagArray`
/// - `fn as_mut_ptr(&mut self) -> *mut sys::SPropTagArray`.
///
/// ### Sample
/// ```
/// # use outlook_mapi::{sys, SizedSPropTagArray};
/// #
/// SizedSPropTagArray! { PropTagArray[2] }
///
/// let prop_tag_array = PropTagArray {
///     aulPropTag: [
///         sys::PR_ENTRYID,
///         sys::PR_DISPLAY_NAME_W,
///     ],
///     ..Default::default()
/// };
///
/// let prop_tag_array: *const sys::SPropTagArray = prop_tag_array.as_ptr();
/// ```
#[macro_export]
#[allow(non_snake_case)]
macro_rules! SizedSPropTagArray {
    ($name:ident [ $count:expr ]) => {
        #[repr(C)]
        #[allow(non_snake_case)]
        struct $name {
            pub cValues: u32,
            pub aulPropTag: [u32; $count],
        }

        $crate::impl_sized_struct_casts!($name, $crate::sys::SPropTagArray);

        $crate::impl_sized_struct_default!($name {
            cValues: $count as u32,
            aulPropTag: [$crate::sys::PR_NULL; $count],
        });
    };
}

/// Get the size of a [`sys::SPropProblemArray`] struct with `count` entries in
/// [`sys::SPropProblemArray::aProblem`].
pub const fn CbNewSPropProblemArray(count: usize) -> usize {
    size_of_container::<sys::SPropProblemArray, sys::SPropProblem>(count)
}

/// Get the size of a [`sys::SPropProblemArray`] struct with [`sys::SPropProblemArray::cProblem`]
/// entries in [`sys::SPropProblemArray::aProblem`].
pub const fn CbSPropProblemArray(prop_problem_array: &sys::SPropProblemArray) -> usize {
    CbNewSPropProblemArray(prop_problem_array.cProblem as usize)
}

/// Declare a variable length struct with the same layout as [`sys::SPropProblemArray`] and
/// implement casting functions:
///
/// - `fn as_ptr(&self) -> *const sys::SPropProblemArray`
/// - `fn as_mut_ptr(&mut self) -> *mut sys::SPropProblemArray`.
///
/// ### Sample
/// ```
/// # use outlook_mapi::{sys, SizedSPropProblemArray};
/// #
/// SizedSPropProblemArray! { PropProblemArray[2] }
///
/// let prop_problem_array = PropProblemArray {
///     aProblem: [
///         sys::SPropProblem {
///             ulIndex: 0,
///             ulPropTag: sys::PR_ENTRYID,
///             scode: sys::MAPI_E_NOT_FOUND.0
///         },
///         sys::SPropProblem {
///             ulIndex: 1,
///             ulPropTag: sys::PR_DISPLAY_NAME_W,
///             scode: sys::MAPI_E_NOT_FOUND.0
///         },
///     ],
///     ..Default::default()
/// };
///
/// let prop_problem_array: *const sys::SPropProblemArray = prop_problem_array.as_ptr();
/// ```
#[macro_export]
#[allow(non_snake_case)]
macro_rules! SizedSPropProblemArray {
    ($name:ident [ $count:expr ]) => {
        #[repr(C)]
        #[allow(non_snake_case)]
        struct $name {
            pub cProblem: u32,
            pub aProblem: [$crate::sys::SPropProblem; $count],
        }

        $crate::impl_sized_struct_casts!($name, $crate::sys::SPropProblemArray);

        {
            const DEFAULT_VALUE: $crate::sys::SPropProblem = $crate::sys::SPropProblem {
                ulIndex: 0,
                ulPropTag: $crate::sys::PR_NULL,
                scode: 0,
            };

            $crate::impl_sized_struct_default!($name {
                cProblem: $count as u32,
                aProblem: [DEFAULT_VALUE; $count],
            });
        }
    };
}

/// Get the size of a [`sys::FLATENTRY`] struct with `count` bytes in [`sys::FLATENTRY::abEntry`].
pub const fn CbNewFLATENTRY(count: usize) -> usize {
    size_of_container::<sys::FLATENTRY, u8>(count)
}

/// Get the size of a [`sys::FLATENTRY`] struct with [`sys::FLATENTRY::cb`] bytes in
/// [`sys::FLATENTRY::abEntry`].
pub const fn CbFLATENTRY(flat_entry: &sys::FLATENTRY) -> usize {
    CbNewFLATENTRY(flat_entry.cb as usize)
}

/// Get the size of a [`sys::FLATENTRYLIST`] struct with `count` bytes in
/// [`sys::FLATENTRYLIST::abEntries`].
pub const fn CbNewFLATENTRYLIST(count: usize) -> usize {
    size_of_container::<sys::FLATENTRYLIST, u8>(count)
}

/// Get the size of a [`sys::FLATENTRYLIST`] struct with [`sys::FLATENTRYLIST::cbEntries`] bytes in
/// [`sys::FLATENTRYLIST::abEntries`].
pub const fn CbFLATENTRYLIST(flat_entry_list: &sys::FLATENTRYLIST) -> usize {
    CbNewFLATENTRYLIST(flat_entry_list.cbEntries as usize)
}

/// Get the size of a [`sys::MTSID`] struct with `count` bytes in [`sys::MTSID::ab`].
pub const fn CbNewMTSID(count: usize) -> usize {
    size_of_container::<sys::MTSID, u8>(count)
}

/// Get the size of a [`sys::MTSID`] struct with [`sys::MTSID::cb`] bytes in [`sys::MTSID::ab`].
pub const fn CbMTSID(mtsid: &sys::MTSID) -> usize {
    CbNewMTSID(mtsid.cb as usize)
}

/// Get the size of a [`sys::FLATMTSIDLIST`] struct with `count` bytes in
/// [`sys::FLATMTSIDLIST::abMTSIDs`].
pub const fn CbNewFLATMTSIDLIST(count: usize) -> usize {
    size_of_container::<sys::FLATMTSIDLIST, u8>(count)
}

/// Get the size of a [`sys::FLATMTSIDLIST`] struct with [`sys::FLATMTSIDLIST::cbMTSIDs`] bytes in
/// [`sys::FLATMTSIDLIST::abMTSIDs`].
pub const fn CbFLATMTSIDLIST(mtsid_list: &sys::FLATMTSIDLIST) -> usize {
    CbNewFLATMTSIDLIST(mtsid_list.cbMTSIDs as usize)
}

/// Get the size of a [`sys::ADRLIST`] struct with `count` entries in [`sys::ADRLIST::aEntries`].
pub const fn CbNewADRLIST(count: usize) -> usize {
    size_of_container::<sys::ADRLIST, sys::ADRENTRY>(count)
}

/// Get the size of a [`sys::ADRLIST`] struct with [`sys::ADRLIST::cEntries`]
/// entries in [`sys::ADRLIST::aEntries`].
pub const fn CbADRLIST(adr_list: &sys::ADRLIST) -> usize {
    CbNewADRLIST(adr_list.cEntries as usize)
}

/// Declare a variable length struct with the same layout as [`sys::ADRLIST`] and implement casting
/// functions:
///
/// - `fn as_ptr(&self) -> *const sys::ADRLIST`
/// - `fn as_mut_ptr(&mut self) -> *mut sys::ADRLIST`.
///
/// ### Sample
/// ```
/// use core::ptr;
/// # use outlook_mapi::{sys, SizedADRLIST};
///
/// SizedADRLIST! { AdrList[2] }
///
/// let adr_list = AdrList {
///     aEntries: [
///         sys::ADRENTRY {
///             ulReserved1: 0,
///             cValues: 0,
///             rgPropVals: ptr::null_mut(),
///         },
///         sys::ADRENTRY {
///             ulReserved1: 0,
///             cValues: 0,
///             rgPropVals: ptr::null_mut(),
///         },
///     ],
///     ..Default::default()
/// };
///
/// let adr_list: *const sys::ADRLIST = adr_list.as_ptr();
/// ```
#[macro_export]
#[allow(non_snake_case)]
macro_rules! SizedADRLIST {
    ($name:ident [ $count:expr ]) => {
        #[repr(C)]
        #[allow(non_snake_case)]
        struct $name {
            pub cEntries: u32,
            pub aEntries: [$crate::sys::ADRENTRY; $count],
        }

        $crate::impl_sized_struct_casts!($name, $crate::sys::ADRLIST);

        {
            const DEFAULT_VALUE: $crate::sys::ADRENTRY = $crate::sys::ADRENTRY {
                ulReserved1: 0,
                cValues: 0,
                rgPropVals: core::ptr::null_mut(),
            };

            $crate::impl_sized_struct_default!($name {
                cEntries: $count as u32,
                aEntries: [DEFAULT_VALUE; $count],
            });
        }
    };
}

/// Get the size of a [`sys::SRowSet`] struct with `count` entries in [`sys::SRowSet::aRow`].
pub const fn CbNewSRowSet(count: usize) -> usize {
    size_of_container::<sys::SRowSet, sys::SRow>(count)
}

/// Get the size of a [`sys::SRowSet`] struct with [`sys::SRowSet::cRows`]
/// entries in [`sys::SRowSet::aRow`].
pub const fn CbSRowSet(row_set: &sys::SRowSet) -> usize {
    CbNewSRowSet(row_set.cRows as usize)
}

/// Declare a variable length struct with the same layout as [`sys::SRowSet`] and implement casting
/// functions:
///
/// - `fn as_ptr(&self) -> *const sys::SRowSet`
/// - `fn as_mut_ptr(&mut self) -> *mut sys::SRowSet`.
///
/// ### Sample
/// ```
/// use core::ptr;
/// # use outlook_mapi::{sys, SizedSRowSet};
///
/// SizedSRowSet! { RowSet[2] }
///
/// let row_set = RowSet {
///     aRow: [
///         sys::SRow {
///             ulAdrEntryPad: 0,
///             cValues: 0,
///             lpProps: ptr::null_mut(),
///         },
///         sys::SRow {
///             ulAdrEntryPad: 0,
///             cValues: 0,
///             lpProps: ptr::null_mut(),
///         },
///     ],
///     ..Default::default()
/// };
///
/// let row_set: *const sys::SRowSet = row_set.as_ptr();
/// ```
#[macro_export]
#[allow(non_snake_case)]
macro_rules! SizedSRowSet {
    ($name:ident [ $count:expr ]) => {
        #[repr(C)]
        #[allow(non_snake_case)]
        struct $name {
            pub cRows: u32,
            pub aRow: [$crate::sys::SRow; $count],
        }

        $crate::impl_sized_struct_casts!($name, $crate::sys::SRowSet);

        {
            const DEFAULT_VALUE: $crate::sys::SRow = $crate::sys::SRow {
                ulAdrEntryPad: 0,
                cValues: 0,
                lpProps: core::ptr::null_mut(),
            };

            $crate::impl_sized_struct_default!($name {
                cRows: $count as u32,
                aRow: [DEFAULT_VALUE; $count],
            });
        }
    };
}

/// Get the size of a [`sys::SSortOrderSet`] struct with `count` entries in
/// [`sys::SSortOrderSet::aSort`].
pub const fn CbNewSSortOrderSet(count: usize) -> usize {
    size_of_container::<sys::SSortOrderSet, sys::SSortOrder>(count)
}

/// Get the size of a [`sys::SSortOrderSet`] struct with [`sys::SSortOrderSet::cSorts`]
/// entries in [`sys::SSortOrderSet::aSort`].
pub const fn CbSSortOrderSet(sort_order_set: &sys::SSortOrderSet) -> usize {
    CbNewSSortOrderSet(sort_order_set.cSorts as usize)
}

/// Declare a variable length struct with the same layout as [`sys::SSortOrderSet`] and implement
/// casting functions:
///
/// - `fn as_ptr(&self) -> *const sys::SSortOrderSet`
/// - `fn as_mut_ptr(&mut self) -> *mut sys::SSortOrderSet`.
///
/// ### Sample
/// ```
/// # use outlook_mapi::{sys, SizedSSortOrderSet};
/// #
/// SizedSSortOrderSet! { SortOrderSet[3] }
///
/// let sort_order_set = SortOrderSet {
///     cCategories: 1,
///     cExpanded: 1,
///     aSort: [
///         sys::SSortOrder {
///             ulPropTag: sys::PR_CONVERSATION_TOPIC_W,
///             ulOrder: sys::TABLE_SORT_DESCEND,
///         },
///         sys::SSortOrder {
///             ulPropTag: sys::PR_MESSAGE_DELIVERY_TIME,
///             ulOrder: sys::TABLE_SORT_CATEG_MAX,
///         },
///         sys::SSortOrder {
///             ulPropTag: sys::PR_CONVERSATION_INDEX,
///             ulOrder: sys::TABLE_SORT_ASCEND,
///         },
///     ],
///     ..Default::default()
/// };
///
/// let sort_order_set: *const sys::SSortOrderSet = sort_order_set.as_ptr();
/// ```
#[macro_export]
#[allow(non_snake_case)]
macro_rules! SizedSSortOrderSet {
    ($name:ident [ $count:expr ]) => {
        #[repr(C)]
        #[allow(non_snake_case)]
        struct $name {
            pub cSorts: u32,
            pub cCategories: u32,
            pub cExpanded: u32,
            pub aSort: [$crate::sys::SSortOrder; $count],
        }

        $crate::impl_sized_struct_casts!($name, $crate::sys::SSortOrderSet);

        {
            const DEFAULT_VALUE: $crate::sys::SSortOrder = $crate::sys::SSortOrder {
                ulPropTag: $crate::sys::PR_NULL,
                ulOrder: $crate::sys::TABLE_SORT_ASCEND,
            };

            $crate::impl_sized_struct_default!($name {
                cSorts: $count as u32,
                cCategories: 0,
                cExpanded: 0,
                aSort: [DEFAULT_VALUE; $count],
            });
        }
    };
}

/// Declare a variable length struct with the same layout as [`sys::DTBLLABEL`] and implement
/// casting functions:
///
/// - `fn as_ptr(&self) -> *const sys::DTBLLABEL`
/// - `fn as_mut_ptr(&mut self) -> *mut sys::DTBLLABEL`
///
/// It also initializes the [`sys::DTBLLABEL::ulbLpszLabelName`] and [`sys::DTBLLABEL::ulFlags`]
/// members and implements either of these accessors to fill in the string buffer, depending on
/// whether it is declared with [`u8`] or [`u16`]:
///
/// - [`u8`]: `fn label_name(&mut self) -> &mut [u8]`
/// - [`u16`]: `fn label_name(&mut self) -> &mut [u16]`
///
/// ### Sample
/// ```
/// # use outlook_mapi::{sys, SizedDtblLabel};
/// use windows_core::PCSTR;
///
/// const LABEL_NAME: &str = "Label Name";
///
/// SizedDtblLabel! { DisplayTableLabelA[u8; LABEL_NAME.len()] }
///
/// let mut display_table_label = DisplayTableLabelA::default();
/// let label_name: Vec<_> = LABEL_NAME.bytes().collect();
/// display_table_label.label_name().copy_from_slice(label_name.as_slice());
/// unsafe {
///     assert_eq!(
///         PCSTR::from_raw(display_table_label.lpszLabelName.as_ptr())
///             .to_string()
///             .expect("invalid string"),
///         LABEL_NAME);
/// }
///
/// let display_table_label: *const sys::DTBLLABEL = display_table_label.as_ptr();
/// ```
#[macro_export]
#[allow(non_snake_case)]
macro_rules! SizedDtblLabel {
    ($name:ident [ $char:ident; $count:expr ]) => {
        #[repr(C)]
        #[allow(non_snake_case)]
        struct $name {
            ulbLpszLabelName: u32,
            ulFlags: u32,
            pub lpszLabelName: [$char; $count + 1],
        }

        $crate::impl_sized_struct_casts!($name, $crate::sys::DTBLLABEL);

        $crate::impl_sized_struct_default!($name {
            ulbLpszLabelName: core::mem::size_of::<$crate::sys::DTBLLABEL>() as u32,
            ulFlags: $crate::display_table_default_flags!($char, $crate::sys::MAPI_UNICODE),
            lpszLabelName: [0; $count + 1],
        });

        impl $name {
            pub fn label_name(&mut self) -> &mut [$char] {
                &mut self.lpszLabelName[..$count]
            }
        }
    };
}

/// Declare a variable length struct with the same layout as [`sys::DTBLEDIT`] and implement
/// casting functions:
///
/// - `fn as_ptr(&self) -> *const sys::DTBLEDIT`
/// - `fn as_mut_ptr(&mut self) -> *mut sys::DTBLEDIT`
///
/// It also initializes the [`sys::DTBLEDIT::ulbLpszCharsAllowed`] and [`sys::DTBLEDIT::ulFlags`]
/// members and implements either of these accessors to fill in the string buffer, depending on
/// whether it is declared with [`u8`] or [`u16`]:
///
/// - [`u8`]: `fn chars_allowed(&mut self) -> &mut [u8]`
/// - [`u16`]: `fn chars_allowed(&mut self) -> &mut [u16]`
///
/// ### Sample
/// ```
/// # use outlook_mapi::{sys, SizedDtblEdit};
/// use windows_core::PCSTR;
///
/// const ALLOWED: &str = "Allowed Characters";
///
/// SizedDtblEdit! { DisplayTableEditA[u8; ALLOWED.len()] }
///
/// let mut display_table_edit = DisplayTableEditA {
///     ulNumCharsAllowed: ALLOWED.len() as u32,
///     ulPropTag: sys::PR_DISPLAY_NAME_A,
///     ..Default::default()
/// };
/// let allowed: Vec<_> = ALLOWED.bytes().collect();
/// display_table_edit.chars_allowed().copy_from_slice(allowed.as_slice());
/// unsafe {
///     assert_eq!(
///         PCSTR::from_raw(display_table_edit.lpszCharsAllowed.as_ptr())
///             .to_string()
///             .expect("invalid string"),
///         ALLOWED);
/// }
///
/// let display_table_edit: *const sys::DTBLEDIT = display_table_edit.as_ptr();
/// ```
#[macro_export]
#[allow(non_snake_case)]
macro_rules! SizedDtblEdit {
    ($name:ident [ $char:ident; $count:expr ]) => {
        #[repr(C)]
        #[allow(non_snake_case)]
        struct $name {
            ulbLpszCharsAllowed: u32,
            ulFlags: u32,
            pub ulNumCharsAllowed: u32,
            pub ulPropTag: u32,
            pub lpszCharsAllowed: [$char; $count + 1],
        }

        $crate::impl_sized_struct_casts!($name, $crate::sys::DTBLEDIT);

        $crate::impl_sized_struct_default!($name {
            ulbLpszCharsAllowed: core::mem::size_of::<$crate::sys::DTBLEDIT>() as u32,
            ulFlags: $crate::display_table_default_flags!($char, $crate::sys::MAPI_UNICODE),
            ulNumCharsAllowed: 0,
            ulPropTag: $crate::sys::PR_NULL,
            lpszCharsAllowed: [0; $count + 1],
        });

        impl $name {
            pub fn chars_allowed(&mut self) -> &mut [$char] {
                &mut self.lpszCharsAllowed[..$count]
            }
        }
    };
}

/// Declare a variable length struct with the same layout as [`sys::DTBLCOMBOBOX`] and implement
/// casting functions:
///
/// - `fn as_ptr(&self) -> *const sys::DTBLCOMBOBOX`
/// - `fn as_mut_ptr(&mut self) -> *mut sys::DTBLCOMBOBOX`
///
/// It also initializes the [`sys::DTBLCOMBOBOX::ulbLpszCharsAllowed`] and
/// [`sys::DTBLCOMBOBOX::ulFlags`] members and implements either of these accessors to fill in the
/// string buffer, depending on whether it is declared with [`u8`] or [`u16`]:
///
/// - [`u8`]: `fn chars_allowed(&mut self) -> &mut [u8]`
/// - [`u16`]: `fn chars_allowed(&mut self) -> &mut [u16]`
///
/// ### Sample
/// ```
/// # use outlook_mapi::{sys, SizedDtblComboBox};
/// use windows_core::PCSTR;
///
/// const ALLOWED: &str = "Allowed Characters";
///
/// SizedDtblComboBox! { DisplayTableComboBoxA[u8; ALLOWED.len()] }
///
/// let mut display_table_combo_box = DisplayTableComboBoxA {
///     ulNumCharsAllowed: ALLOWED.len() as u32,
///     ulPRPropertyName: sys::PR_DISPLAY_NAME_A,
///     ulPRTableName: sys::PR_MESSAGE_DELIVERY_TIME,
///     ..Default::default()
/// };
/// let allowed: Vec<_> = ALLOWED.bytes().collect();
/// display_table_combo_box.chars_allowed().copy_from_slice(allowed.as_slice());
/// unsafe {
///     assert_eq!(
///         PCSTR::from_raw(display_table_combo_box.lpszCharsAllowed.as_ptr())
///             .to_string()
///             .expect("invalid string"),
///         ALLOWED);
/// }
///
/// let display_table_combo_box: *const sys::DTBLCOMBOBOX = display_table_combo_box.as_ptr();
/// ```
#[macro_export]
#[allow(non_snake_case)]
macro_rules! SizedDtblComboBox {
    ($name:ident [ $char:ident; $count:expr ]) => {
        #[repr(C)]
        #[allow(non_snake_case)]
        struct $name {
            ulbLpszCharsAllowed: u32,
            ulFlags: u32,
            pub ulNumCharsAllowed: u32,
            pub ulPRPropertyName: u32,
            pub ulPRTableName: u32,
            pub lpszCharsAllowed: [$char; $count + 1],
        }

        $crate::impl_sized_struct_casts!($name, $crate::sys::DTBLCOMBOBOX);

        $crate::impl_sized_struct_default!($name {
            ulbLpszCharsAllowed: core::mem::size_of::<$crate::sys::DTBLCOMBOBOX>() as u32,
            ulFlags: $crate::display_table_default_flags!($char, $crate::sys::MAPI_UNICODE),
            ulNumCharsAllowed: 0,
            ulPRPropertyName: $crate::sys::PR_NULL,
            ulPRTableName: $crate::sys::PR_NULL,
            lpszCharsAllowed: [0; $count + 1],
        });

        impl $name {
            pub fn chars_allowed(&mut self) -> &mut [$char] {
                &mut self.lpszCharsAllowed[..$count]
            }
        }
    };
}

/// Declare a variable length struct with the same layout as [`sys::DTBLCHECKBOX`] and implement
/// casting functions:
///
/// - `fn as_ptr(&self) -> *const sys::DTBLCHECKBOX`
/// - `fn as_mut_ptr(&mut self) -> *mut sys::DTBLCHECKBOX`
///
/// It also initializes the [`sys::DTBLCHECKBOX::ulbLpszLabel`] and [`sys::DTBLCHECKBOX::ulFlags`]
/// members and implements either of these accessors to fill in the string buffer, depending on
/// whether it is declared with [`u8`] or [`u16`]:
///
/// - [`u8`]: `fn label(&mut self) -> &mut [u8]`
/// - [`u16`]: `fn label(&mut self) -> &mut [u16]`
///
/// ### Sample
/// ```
/// # use outlook_mapi::{sys, SizedDtblCheckBox};
/// use windows_core::PCSTR;
///
/// const LABEL: &str = "Label";
///
/// SizedDtblCheckBox! { DisplayTableCheckBoxA[u8; LABEL.len()] }
///
/// let mut display_table_check_box = DisplayTableCheckBoxA {
///     ulPRPropertyName: sys::PR_DISPLAY_NAME_A,
///     ..Default::default()
/// };
/// let label: Vec<_> = LABEL.bytes().collect();
/// display_table_check_box.label().copy_from_slice(label.as_slice());
/// unsafe {
///     assert_eq!(
///         PCSTR::from_raw(display_table_check_box.lpszLabel.as_ptr())
///             .to_string()
///             .expect("invalid string"),
///         LABEL);
/// }
///
/// let display_table_check_box: *const sys::DTBLCHECKBOX = display_table_check_box.as_ptr();
/// ```
#[macro_export]
#[allow(non_snake_case)]
macro_rules! SizedDtblCheckBox {
    ($name:ident [ $char:ident; $count:expr ]) => {
        #[repr(C)]
        #[allow(non_snake_case)]
        struct $name {
            ulbLpszLabel: u32,
            ulFlags: u32,
            pub ulPRPropertyName: u32,
            pub lpszLabel: [$char; $count + 1],
        }

        $crate::impl_sized_struct_casts!($name, $crate::sys::DTBLCHECKBOX);

        $crate::impl_sized_struct_default!($name {
            ulbLpszLabel: core::mem::size_of::<$crate::sys::DTBLCHECKBOX>() as u32,
            ulFlags: $crate::display_table_default_flags!($char, $crate::sys::MAPI_UNICODE),
            ulPRPropertyName: $crate::sys::PR_NULL,
            lpszLabel: [0; $count + 1],
        });

        impl $name {
            pub fn label(&mut self) -> &mut [$char] {
                &mut self.lpszLabel[..$count]
            }
        }
    };
}

/// Declare a variable length struct with the same layout as [`sys::DTBLGROUPBOX`] and implement
/// casting functions:
///
/// - `fn as_ptr(&self) -> *const sys::DTBLGROUPBOX`
/// - `fn as_mut_ptr(&mut self) -> *mut sys::DTBLGROUPBOX`
///
/// It also initializes the [`sys::DTBLGROUPBOX::ulbLpszLabel`] and [`sys::DTBLGROUPBOX::ulFlags`]
/// members and implements either of these accessors to fill in the string buffer, depending on
/// whether it is declared with [`u8`] or [`u16`]:
///
/// - [`u8`]: `fn label(&mut self) -> &mut [u8]`
/// - [`u16`]: `fn label(&mut self) -> &mut [u16]`
///
/// ### Sample
/// ```
/// # use outlook_mapi::{sys, SizedDtblGroupBox};
/// use windows_core::PCSTR;
///
/// const LABEL: &str = "Label";
///
/// SizedDtblGroupBox! { DisplayTableGroupBoxA[u8; LABEL.len()] }
///
/// let mut display_table_group_box = DisplayTableGroupBoxA::default();
/// let label: Vec<_> = LABEL.bytes().collect();
/// display_table_group_box.label().copy_from_slice(label.as_slice());
/// unsafe {
///     assert_eq!(
///         PCSTR::from_raw(display_table_group_box.lpszLabel.as_ptr())
///             .to_string()
///             .expect("invalid string"),
///         LABEL);
/// }
///
/// let display_table_group_box: *const sys::DTBLGROUPBOX = display_table_group_box.as_ptr();
/// ```
#[macro_export]
#[allow(non_snake_case)]
macro_rules! SizedDtblGroupBox {
    ($name:ident [ $char:ident; $count:expr ]) => {
        #[repr(C)]
        #[allow(non_snake_case)]
        struct $name {
            ulbLpszLabel: u32,
            ulFlags: u32,
            pub lpszLabel: [$char; $count + 1],
        }

        $crate::impl_sized_struct_casts!($name, $crate::sys::DTBLGROUPBOX);

        $crate::impl_sized_struct_default!($name {
            ulbLpszLabel: core::mem::size_of::<$crate::sys::DTBLGROUPBOX>() as u32,
            ulFlags: $crate::display_table_default_flags!($char, $crate::sys::MAPI_UNICODE),
            lpszLabel: [0; $count + 1],
        });

        impl $name {
            pub fn label(&mut self) -> &mut [$char] {
                &mut self.lpszLabel[..$count]
            }
        }
    };
}

/// Declare a variable length struct with the same layout as [`sys::DTBLBUTTON`] and implement
/// casting functions:
///
/// - `fn as_ptr(&self) -> *const sys::DTBLBUTTON`
/// - `fn as_mut_ptr(&mut self) -> *mut sys::DTBLBUTTON`
///
/// It also initializes the [`sys::DTBLBUTTON::ulbLpszLabel`] and [`sys::DTBLBUTTON::ulFlags`]
/// members and implements either of these accessors to fill in the string buffer, depending on
/// whether it is declared with [`u8`] or [`u16`]:
///
/// - [`u8`]: `fn label(&mut self) -> &mut [u8]`
/// - [`u16`]: `fn label(&mut self) -> &mut [u16]`
///
/// ### Sample
/// ```
/// # use outlook_mapi::{sys, SizedDtblButton};
/// use windows_core::PCSTR;
///
/// const LABEL: &str = "Label";
///
/// SizedDtblButton! { DisplayTableButtonA[u8; LABEL.len()] }
///
/// let mut display_table_button = DisplayTableButtonA {
///     ulPRControl: sys::PR_DISPLAY_NAME_A,
///     ..Default::default()
/// };
/// let label: Vec<_> = LABEL.bytes().collect();
/// display_table_button.label().copy_from_slice(label.as_slice());
/// unsafe {
///     assert_eq!(
///         PCSTR::from_raw(display_table_button.lpszLabel.as_ptr())
///             .to_string()
///             .expect("invalid string"),
///         LABEL);
/// }
///
/// let display_table_button: *const sys::DTBLBUTTON = display_table_button.as_ptr();
/// ```
#[macro_export]
#[allow(non_snake_case)]
macro_rules! SizedDtblButton {
    ($name:ident [ $char:ident; $count:expr ]) => {
        #[repr(C)]
        #[allow(non_snake_case)]
        struct $name {
            ulbLpszLabel: u32,
            ulFlags: u32,
            pub ulPRControl: u32,
            pub lpszLabel: [$char; $count + 1],
        }

        $crate::impl_sized_struct_casts!($name, $crate::sys::DTBLBUTTON);

        $crate::impl_sized_struct_default!($name {
            ulbLpszLabel: core::mem::size_of::<$crate::sys::DTBLBUTTON>() as u32,
            ulFlags: $crate::display_table_default_flags!($char, $crate::sys::MAPI_UNICODE),
            ulPRControl: $crate::sys::PR_NULL,
            lpszLabel: [0; $count + 1],
        });

        impl $name {
            pub fn label(&mut self) -> &mut [$char] {
                &mut self.lpszLabel[..$count]
            }
        }
    };
}

/// Declare a variable length struct with the same layout as [`sys::DTBLPAGE`] and implement
/// casting functions:
///
/// - `fn as_ptr(&self) -> *const sys::DTBLPAGE`
/// - `fn as_mut_ptr(&mut self) -> *mut sys::DTBLPAGE`
///
/// It also initializes the [`sys::DTBLPAGE::ulbLpszLabel`], [`sys::DTBLPAGE::ulFlags`], and
/// [`sys::DTBLPAGE::ulbLpszComponent`] members and implements either of these accessor pairs to
/// fill in the string buffers, depending on whether it is declared with [`u8`] or [`u16`]:
///
/// - [`u8`]: `fn label(&mut self) -> &mut [u8]`, and `fn component(&mut self) -> &mut [u8]`
/// - [`u16`]: `fn label(&mut self) -> &mut [u16]`, and `fn component(&mut self) -> &mut [u16]`
///
/// ### Sample
/// ```
/// # use outlook_mapi::{sys, SizedDtblPage};
/// use windows_core::PCSTR;
///
/// const LABEL: &str = "Label";
/// const COMPONENT: &str = "Component";
///
/// SizedDtblPage! { DisplayTablePageA[u8; LABEL.len(); COMPONENT.len()] }
///
/// let mut display_table_page = DisplayTablePageA {
///     ulContext: 10,
///     ..Default::default()
/// };
/// let label: Vec<_> = LABEL.bytes().collect();
/// display_table_page.label().copy_from_slice(label.as_slice());
/// let component: Vec<_> = COMPONENT.bytes().collect();
/// display_table_page.component().copy_from_slice(component.as_slice());
/// unsafe {
///     assert_eq!(
///         PCSTR::from_raw(display_table_page.lpszLabel.as_ptr())
///             .to_string()
///             .expect("invalid string"),
///         LABEL);
///     assert_eq!(
///         PCSTR::from_raw(display_table_page.lpszComponent.as_ptr())
///             .to_string()
///             .expect("invalid string"),
///         COMPONENT);
/// }
///
/// let display_table_page: *const sys::DTBLPAGE = display_table_page.as_ptr();
/// ```
#[macro_export]
#[allow(non_snake_case)]
macro_rules! SizedDtblPage {
    ($name:ident [ $char:ident; $count1:expr; $count2:expr ]) => {
        #[repr(C)]
        #[allow(non_snake_case)]
        struct $name {
            ulbLpszLabel: u32,
            ulFlags: u32,
            ulbLpszComponent: u32,
            pub ulContext: u32,
            pub lpszLabel: [$char; $count1 + 1],
            pub lpszComponent: [$char; $count2 + 1],
        }

        $crate::impl_sized_struct_casts!($name, $crate::sys::DTBLPAGE);

        $crate::impl_sized_struct_default!($name {
            ulbLpszLabel: core::mem::size_of::<$crate::sys::DTBLPAGE>() as u32,
            ulFlags: $crate::display_table_default_flags!($char, $crate::sys::MAPI_UNICODE),
            ulbLpszComponent: (core::mem::size_of::<$crate::sys::DTBLPAGE>()
                + core::mem::size_of::<[$char; $count1 + 1]>()) as u32,
            ulContext: 0,
            lpszLabel: [0; $count1 + 1],
            lpszComponent: [0; $count2 + 1],
        });

        impl $name {
            pub fn label(&mut self) -> &mut [$char] {
                &mut self.lpszLabel[..$count1]
            }

            pub fn component(&mut self) -> &mut [$char] {
                &mut self.lpszComponent[..$count2]
            }
        }
    };
}

/// Declare a variable length struct with the same layout as [`sys::DTBLRADIOBUTTON`] and implement
/// casting functions:
///
/// - `fn as_ptr(&self) -> *const sys::DTBLRADIOBUTTON`
/// - `fn as_mut_ptr(&mut self) -> *mut sys::DTBLRADIOBUTTON`
///
/// It also initializes the [`sys::DTBLRADIOBUTTON::ulbLpszLabel`] and
/// [`sys::DTBLRADIOBUTTON::ulFlags`] members and implements either of these accessors to fill in
/// the string buffer, depending on whether it is declared with [`u8`] or [`u16`]:
///
/// - [`u8`]: `fn label(&mut self) -> &mut [u8]`
/// - [`u16`]: `fn label(&mut self) -> &mut [u16]`
///
/// ### Sample
/// ```
/// # use outlook_mapi::{sys, SizedDtblRadioButton};
/// use windows_core::PCSTR;
///
/// const LABEL: &str = "Label";
///
/// SizedDtblRadioButton! { DisplayTableRadioButtonA[u8; LABEL.len()] }
///
/// let mut display_table_radio_button = DisplayTableRadioButtonA {
///     ulcButtons: 10,
///     ulPropTag: sys::PR_DISPLAY_NAME_A,
///     lReturnValue: -1,
///     ..Default::default()
/// };
/// let label: Vec<_> = LABEL.bytes().collect();
/// display_table_radio_button.label().copy_from_slice(label.as_slice());
/// unsafe {
///     assert_eq!(
///         PCSTR::from_raw(display_table_radio_button.lpszLabel.as_ptr())
///             .to_string()
///             .expect("invalid string"),
///         LABEL);
/// }
///
/// let display_table_radio_button: *const sys::DTBLRADIOBUTTON = display_table_radio_button.as_ptr();
/// ```
#[macro_export]
#[allow(non_snake_case)]
macro_rules! SizedDtblRadioButton {
    ($name:ident [ $char:ident; $count:expr ]) => {
        #[repr(C)]
        #[allow(non_snake_case)]
        struct $name {
            ulbLpszLabel: u32,
            ulFlags: u32,
            pub ulcButtons: u32,
            pub ulPropTag: u32,
            pub lReturnValue: i32,
            pub lpszLabel: [$char; $count + 1],
        }

        $crate::impl_sized_struct_casts!($name, $crate::sys::DTBLRADIOBUTTON);

        $crate::impl_sized_struct_default!($name {
            ulbLpszLabel: core::mem::size_of::<$crate::sys::DTBLRADIOBUTTON>() as u32,
            ulFlags: $crate::display_table_default_flags!($char, $crate::sys::MAPI_UNICODE),
            ulcButtons: 0,
            ulPropTag: $crate::sys::PR_NULL,
            lReturnValue: 0,
            lpszLabel: [0; $count + 1],
        });

        impl $name {
            pub fn label(&mut self) -> &mut [$char] {
                &mut self.lpszLabel[..$count]
            }
        }
    };
}

mod impl_macros {
    /// Build the common casting function `impl` block for all of the SizedXXX macros.
    #[macro_export]
    #[doc(hidden)]
    macro_rules! impl_sized_struct_casts {
        ($name:ident, $sys_type:path) => {
            #[allow(dead_code)]
            impl $name {
                pub fn as_ptr(&self) -> *const $sys_type {
                    unsafe { std::mem::transmute::<&Self, &$sys_type>(self) }
                }

                pub fn as_mut_ptr(&mut self) -> *mut $sys_type {
                    unsafe { std::mem::transmute::<&mut Self, &mut $sys_type>(self) }
                }
            }
        };
    }

    /// Build an optional `impl Default` block for any of the SizedXXX macros.
    #[macro_export]
    #[doc(hidden)]
    macro_rules! impl_sized_struct_default {
    ($name:ident $body:tt) => {
        #[allow(dead_code)]
        impl Default for $name {
            fn default() -> Self {
                Self $body
            }
        }
    };
}

    /// Get the `ulFlags` default value for any of the display table SizedXXX macros.
    #[macro_export]
    #[doc(hidden)]
    macro_rules! display_table_default_flags {
        (u8, $unicode:expr) => {
            0
        };
        (u16, $unicode:expr) => {
            $unicode
        };
    }
}

#[cfg(test)]
mod tests {
    use crate::*;
    use core::{mem, ptr};
    use windows_core::{PCSTR, PCWSTR};

    #[test]
    fn sized_entry_id() {
        SizedENTRYID! { EntryId[12] }

        assert_eq!(mem::size_of::<EntryId>(), CbNewENTRYID(12));
        let entry_id = EntryId {
            abFlags: [0x0, 0x1, 0x2, 0x3],
            ab: [0x4, 0x5, 0x6, 0x7, 0x8, 0x9, 0xa, 0xb, 0xc, 0xd, 0xe, 0xf],
        };

        assert_eq!(mem::size_of::<sys::ENTRYID>(), CbNewENTRYID(1));
        assert_eq!(mem::size_of::<sys::ENTRYID>(), CbENTRYID(1));
        let entry_id: *const sys::ENTRYID = entry_id.as_ptr();
        let entry_id = unsafe { entry_id.as_ref() }.unwrap();
        assert_eq!(entry_id.abFlags, [0x0, 0x1, 0x2, 0x3]);
        assert_eq!(
            entry_id.ab,
            [0x4],
            "can only see the first entry in the sys type"
        );
    }

    #[test]
    fn sized_prop_tag_array() {
        SizedSPropTagArray!(PropTagArray[2]);

        assert_eq!(mem::size_of::<PropTagArray>(), CbNewSPropTagArray(2));
        let prop_tag_array = PropTagArray {
            aulPropTag: [sys::PR_ENTRYID, sys::PR_DISPLAY_NAME_W],
            ..Default::default()
        };

        assert_eq!(mem::size_of::<sys::SPropTagArray>(), CbNewSPropTagArray(1));
        let prop_tag_array: *const sys::SPropTagArray = prop_tag_array.as_ptr();
        let prop_tag_array = unsafe { prop_tag_array.as_ref() }.unwrap();
        assert_eq!(CbNewSPropTagArray(2), CbSPropTagArray(prop_tag_array));
        assert_eq!(prop_tag_array.cValues, 2);
        assert_eq!(
            prop_tag_array.aulPropTag,
            [sys::PR_ENTRYID],
            "can only see the first entry in the sys type"
        );
    }

    #[test]
    fn sized_prop_problem_array() {
        SizedSPropProblemArray!(PropProblemArray[2]);

        assert_eq!(
            mem::size_of::<PropProblemArray>(),
            CbNewSPropProblemArray(2)
        );
        let prop_problem_array = PropProblemArray {
            aProblem: [
                sys::SPropProblem {
                    ulIndex: 0,
                    ulPropTag: sys::PR_ENTRYID,
                    scode: sys::MAPI_E_NOT_FOUND.0,
                },
                sys::SPropProblem {
                    ulIndex: 1,
                    ulPropTag: sys::PR_DISPLAY_NAME_W,
                    scode: sys::MAPI_E_NOT_FOUND.0,
                },
            ],
            ..Default::default()
        };

        assert_eq!(
            mem::size_of::<sys::SPropProblemArray>(),
            CbNewSPropProblemArray(1)
        );
        let prop_problem_array: *const sys::SPropProblemArray = prop_problem_array.as_ptr();
        let prop_problem_array = unsafe { prop_problem_array.as_ref() }.unwrap();
        assert_eq!(
            CbNewSPropProblemArray(2),
            CbSPropProblemArray(prop_problem_array)
        );
        assert_eq!(prop_problem_array.cProblem, 2);
        assert_eq!(
            prop_problem_array.aProblem,
            [sys::SPropProblem {
                ulIndex: 0,
                ulPropTag: sys::PR_ENTRYID,
                scode: sys::MAPI_E_NOT_FOUND.0,
            }],
            "can only see the first entry in the sys type"
        );
    }

    #[test]
    fn sized_flat_lists() {
        assert_eq!(mem::size_of::<sys::FLATENTRY>(), CbNewFLATENTRY(1));
        assert_eq!(mem::size_of::<sys::FLATENTRYLIST>(), CbNewFLATENTRYLIST(1));
        assert_eq!(mem::size_of::<sys::MTSID>(), CbNewMTSID(1));
        assert_eq!(mem::size_of::<sys::FLATMTSIDLIST>(), CbNewFLATMTSIDLIST(1));
    }

    #[test]
    fn sized_adr_list() {
        SizedADRLIST!(AdrList[2]);

        assert_eq!(mem::size_of::<AdrList>(), CbNewADRLIST(2));
        let adr_list = AdrList {
            aEntries: [
                sys::ADRENTRY {
                    ulReserved1: 0,
                    cValues: 0,
                    rgPropVals: ptr::null_mut(),
                },
                sys::ADRENTRY {
                    ulReserved1: 0,
                    cValues: 0,
                    rgPropVals: ptr::null_mut(),
                },
            ],
            ..Default::default()
        };

        assert_eq!(mem::size_of::<sys::ADRLIST>(), CbNewADRLIST(1));
        let adr_list: *const sys::ADRLIST = adr_list.as_ptr();
        let adr_list = unsafe { adr_list.as_ref() }.unwrap();
        assert_eq!(CbNewADRLIST(2), CbADRLIST(adr_list));
        assert_eq!(adr_list.cEntries, 2);
        assert_eq!(
            adr_list.aEntries,
            [sys::ADRENTRY {
                ulReserved1: 0,
                cValues: 0,
                rgPropVals: ptr::null_mut(),
            }],
            "can only see the first entry in the sys type"
        );
    }

    #[test]
    fn sized_row_set() {
        SizedSRowSet!(RowSet[2]);

        assert_eq!(mem::size_of::<RowSet>(), CbNewSRowSet(2));
        let row_set = RowSet {
            aRow: [
                sys::SRow {
                    ulAdrEntryPad: 0,
                    cValues: 0,
                    lpProps: ptr::null_mut(),
                },
                sys::SRow {
                    ulAdrEntryPad: 0,
                    cValues: 0,
                    lpProps: ptr::null_mut(),
                },
            ],
            ..Default::default()
        };

        assert_eq!(mem::size_of::<sys::SRowSet>(), CbNewSRowSet(1));
        let row_set: *const sys::SRowSet = row_set.as_ptr();
        let row_set = unsafe { row_set.as_ref() }.unwrap();
        assert_eq!(CbNewSRowSet(2), CbSRowSet(row_set));
        assert_eq!(row_set.cRows, 2);
        assert_eq!(
            row_set.aRow,
            [sys::SRow {
                ulAdrEntryPad: 0,
                cValues: 0,
                lpProps: ptr::null_mut(),
            }],
            "can only see the first entry in the sys type"
        );
    }

    #[test]
    fn sized_sort_order_set() {
        SizedSSortOrderSet!(SortOrderSet[3]);

        assert_eq!(mem::size_of::<SortOrderSet>(), CbNewSSortOrderSet(3));
        let sort_order_set = SortOrderSet {
            cCategories: 1,
            cExpanded: 1,
            aSort: [
                sys::SSortOrder {
                    ulPropTag: sys::PR_CONVERSATION_TOPIC_W,
                    ulOrder: sys::TABLE_SORT_DESCEND,
                },
                sys::SSortOrder {
                    ulPropTag: sys::PR_MESSAGE_DELIVERY_TIME,
                    ulOrder: sys::TABLE_SORT_CATEG_MAX,
                },
                sys::SSortOrder {
                    ulPropTag: sys::PR_CONVERSATION_INDEX,
                    ulOrder: sys::TABLE_SORT_ASCEND,
                },
            ],
            ..Default::default()
        };

        assert_eq!(mem::size_of::<sys::SSortOrderSet>(), CbNewSSortOrderSet(1));
        let sort_order_set: *const sys::SSortOrderSet = sort_order_set.as_ptr();
        let sort_order_set = unsafe { sort_order_set.as_ref() }.unwrap();
        assert_eq!(CbNewSSortOrderSet(3), CbSSortOrderSet(sort_order_set));
        assert_eq!(sort_order_set.cSorts, 3);
        assert_eq!(sort_order_set.cCategories, 1);
        assert_eq!(sort_order_set.cExpanded, 1);
        assert_eq!(
            sort_order_set.aSort,
            [sys::SSortOrder {
                ulPropTag: sys::PR_CONVERSATION_TOPIC_W,
                ulOrder: sys::TABLE_SORT_DESCEND,
            }],
            "can only see the first entry in the sys type"
        );
    }

    #[test]
    fn sized_display_table_label_a() {
        const LABEL: &str = "Display Table Label";

        SizedDtblLabel! { DisplayTableLabelA[u8; LABEL.len()] }

        let mut display_table_label = DisplayTableLabelA::default();
        let label: Vec<_> = LABEL.bytes().collect();
        assert_eq!(LABEL.len(), label.len());
        display_table_label
            .label_name()
            .copy_from_slice(label.as_slice());
        unsafe {
            assert_eq!(
                PCSTR::from_raw(display_table_label.lpszLabelName.as_ptr())
                    .to_string()
                    .expect("invalid string"),
                LABEL
            );
        }

        let display_table_label: *const sys::DTBLLABEL = display_table_label.as_ptr();
        let display_table_label = unsafe { display_table_label.as_ref() }.unwrap();
        assert_eq!(
            display_table_label.ulbLpszLabelName,
            mem::size_of::<sys::DTBLLABEL>() as u32
        );
        assert_eq!(display_table_label.ulFlags, 0);
    }

    #[test]
    fn sized_display_table_label_w() {
        const LABEL: &str = "Display Table Label";

        SizedDtblLabel! { DisplayTableLabelW[u16; LABEL.len()] }

        let mut display_table_label = DisplayTableLabelW::default();
        let label: Vec<_> = LABEL.encode_utf16().collect();
        assert_eq!(LABEL.len(), label.len());
        display_table_label
            .label_name()
            .copy_from_slice(label.as_slice());
        unsafe {
            assert_eq!(
                PCWSTR::from_raw(display_table_label.lpszLabelName.as_ptr())
                    .to_string()
                    .expect("invalid string"),
                LABEL
            );
        }

        let display_table_label: *const sys::DTBLLABEL = display_table_label.as_ptr();
        let display_table_label = unsafe { display_table_label.as_ref() }.unwrap();
        assert_eq!(
            display_table_label.ulbLpszLabelName,
            mem::size_of::<sys::DTBLLABEL>() as u32
        );
        assert_eq!(display_table_label.ulFlags, sys::MAPI_UNICODE);
    }

    #[test]
    fn sized_display_table_edit_a() {
        const ALLOWED: &str = "Allowed Characters";

        SizedDtblEdit! { DisplayTableEditA[u8; ALLOWED.len()] }

        let mut display_table_edit = DisplayTableEditA {
            ulNumCharsAllowed: ALLOWED.len() as u32,
            ulPropTag: sys::PR_DISPLAY_NAME_A,
            ..Default::default()
        };
        let allowed: Vec<_> = ALLOWED.bytes().collect();
        assert_eq!(ALLOWED.len(), allowed.len());
        display_table_edit
            .chars_allowed()
            .copy_from_slice(allowed.as_slice());
        unsafe {
            assert_eq!(
                PCSTR::from_raw(display_table_edit.lpszCharsAllowed.as_ptr())
                    .to_string()
                    .expect("invalid string"),
                ALLOWED
            );
        }

        let display_table_edit: *const sys::DTBLEDIT = display_table_edit.as_ptr();
        let display_table_edit = unsafe { display_table_edit.as_ref() }.unwrap();
        assert_eq!(
            display_table_edit.ulbLpszCharsAllowed,
            mem::size_of::<sys::DTBLEDIT>() as u32
        );
        assert_eq!(display_table_edit.ulFlags, 0);
        assert_eq!(display_table_edit.ulNumCharsAllowed, ALLOWED.len() as u32);
        assert_eq!(display_table_edit.ulPropTag, sys::PR_DISPLAY_NAME_A);
    }

    #[test]
    fn sized_display_table_edit_w() {
        const ALLOWED: &str = "Allowed Characters";

        SizedDtblEdit! { DisplayTableEditW[u16; ALLOWED.len()] }

        let mut display_table_edit = DisplayTableEditW {
            ulNumCharsAllowed: ALLOWED.len() as u32,
            ulPropTag: sys::PR_DISPLAY_NAME_W,
            ..Default::default()
        };
        let allowed: Vec<_> = ALLOWED.encode_utf16().collect();
        assert_eq!(ALLOWED.len(), allowed.len());
        display_table_edit
            .chars_allowed()
            .copy_from_slice(allowed.as_slice());
        unsafe {
            assert_eq!(
                PCWSTR::from_raw(display_table_edit.lpszCharsAllowed.as_ptr())
                    .to_string()
                    .expect("invalid string"),
                ALLOWED
            );
        }

        let display_table_edit: *const sys::DTBLEDIT = display_table_edit.as_ptr();
        let display_table_edit = unsafe { display_table_edit.as_ref() }.unwrap();
        assert_eq!(
            display_table_edit.ulbLpszCharsAllowed,
            mem::size_of::<sys::DTBLEDIT>() as u32
        );
        assert_eq!(display_table_edit.ulFlags, sys::MAPI_UNICODE);
        assert_eq!(display_table_edit.ulNumCharsAllowed, ALLOWED.len() as u32);
        assert_eq!(display_table_edit.ulPropTag, sys::PR_DISPLAY_NAME_W);
    }

    #[test]
    fn sized_display_table_combo_box_a() {
        const ALLOWED: &str = "Allowed Characters";

        SizedDtblComboBox! { DisplayTableComboBoxA[u8; ALLOWED.len()] }

        let mut display_table_combo_box = DisplayTableComboBoxA {
            ulNumCharsAllowed: ALLOWED.len() as u32,
            ulPRPropertyName: sys::PR_DISPLAY_NAME_A,
            ulPRTableName: sys::PR_MESSAGE_DELIVERY_TIME,
            ..Default::default()
        };
        let allowed: Vec<_> = ALLOWED.bytes().collect();
        assert_eq!(ALLOWED.len(), allowed.len());
        display_table_combo_box
            .chars_allowed()
            .copy_from_slice(allowed.as_slice());
        unsafe {
            assert_eq!(
                PCSTR::from_raw(display_table_combo_box.lpszCharsAllowed.as_ptr())
                    .to_string()
                    .expect("invalid string"),
                ALLOWED
            );
        }

        let display_table_combo_box: *const sys::DTBLCOMBOBOX = display_table_combo_box.as_ptr();
        let display_table_combo_box = unsafe { display_table_combo_box.as_ref() }.unwrap();
        assert_eq!(
            display_table_combo_box.ulbLpszCharsAllowed,
            mem::size_of::<sys::DTBLCOMBOBOX>() as u32
        );
        assert_eq!(display_table_combo_box.ulFlags, 0);
        assert_eq!(
            display_table_combo_box.ulNumCharsAllowed,
            ALLOWED.len() as u32
        );
        assert_eq!(
            display_table_combo_box.ulPRPropertyName,
            sys::PR_DISPLAY_NAME_A
        );
        assert_eq!(
            display_table_combo_box.ulPRTableName,
            sys::PR_MESSAGE_DELIVERY_TIME
        );
    }

    #[test]
    fn sized_display_table_combo_box_w() {
        const ALLOWED: &str = "Allowed Characters";

        SizedDtblComboBox! { DisplayTableComboBoxW[u16; ALLOWED.len()] }

        let mut display_table_combo_box = DisplayTableComboBoxW {
            ulNumCharsAllowed: ALLOWED.len() as u32,
            ulPRPropertyName: sys::PR_DISPLAY_NAME_W,
            ulPRTableName: sys::PR_MESSAGE_DELIVERY_TIME,
            ..Default::default()
        };
        let allowed: Vec<_> = ALLOWED.encode_utf16().collect();
        assert_eq!(ALLOWED.len(), allowed.len());
        display_table_combo_box
            .chars_allowed()
            .copy_from_slice(allowed.as_slice());
        unsafe {
            assert_eq!(
                PCWSTR::from_raw(display_table_combo_box.lpszCharsAllowed.as_ptr())
                    .to_string()
                    .expect("invalid string"),
                ALLOWED
            );
        }

        let display_table_combo_box: *const sys::DTBLCOMBOBOX = display_table_combo_box.as_ptr();
        let display_table_combo_box = unsafe { display_table_combo_box.as_ref() }.unwrap();
        assert_eq!(
            display_table_combo_box.ulbLpszCharsAllowed,
            mem::size_of::<sys::DTBLCOMBOBOX>() as u32
        );
        assert_eq!(display_table_combo_box.ulFlags, sys::MAPI_UNICODE);
        assert_eq!(
            display_table_combo_box.ulNumCharsAllowed,
            ALLOWED.len() as u32
        );
        assert_eq!(
            display_table_combo_box.ulPRPropertyName,
            sys::PR_DISPLAY_NAME_W
        );
        assert_eq!(
            display_table_combo_box.ulPRTableName,
            sys::PR_MESSAGE_DELIVERY_TIME
        );
    }

    #[test]
    fn sized_display_table_check_box_a() {
        const LABEL: &str = "Checkbox Label";

        SizedDtblCheckBox! { DisplayTableCheckBoxA[u8; LABEL.len()] }

        let mut display_table_check_box = DisplayTableCheckBoxA {
            ulPRPropertyName: sys::PR_DISPLAY_NAME_A,
            ..Default::default()
        };
        let label: Vec<_> = LABEL.bytes().collect();
        assert_eq!(LABEL.len(), label.len());
        display_table_check_box
            .label()
            .copy_from_slice(label.as_slice());
        unsafe {
            assert_eq!(
                PCSTR::from_raw(display_table_check_box.lpszLabel.as_ptr())
                    .to_string()
                    .expect("invalid string"),
                LABEL
            );
        }

        let display_table_check_box: *const sys::DTBLCHECKBOX = display_table_check_box.as_ptr();
        let display_table_check_box = unsafe { display_table_check_box.as_ref() }.unwrap();
        assert_eq!(
            display_table_check_box.ulbLpszLabel,
            mem::size_of::<sys::DTBLCHECKBOX>() as u32
        );
        assert_eq!(display_table_check_box.ulFlags, 0);
        assert_eq!(
            display_table_check_box.ulPRPropertyName,
            sys::PR_DISPLAY_NAME_A
        );
    }

    #[test]
    fn sized_display_table_check_box_w() {
        const LABEL: &str = "Checkbox Label";

        SizedDtblCheckBox! { DisplayTableCheckBoxW[u16; LABEL.len()] }

        let mut display_table_check_box = DisplayTableCheckBoxW {
            ulPRPropertyName: sys::PR_DISPLAY_NAME_W,
            ..Default::default()
        };
        let label: Vec<_> = LABEL.encode_utf16().collect();
        assert_eq!(LABEL.len(), label.len());
        display_table_check_box
            .label()
            .copy_from_slice(label.as_slice());
        unsafe {
            assert_eq!(
                PCWSTR::from_raw(display_table_check_box.lpszLabel.as_ptr())
                    .to_string()
                    .expect("invalid string"),
                LABEL
            );
        }

        let display_table_check_box: *const sys::DTBLCHECKBOX = display_table_check_box.as_ptr();
        let display_table_check_box = unsafe { display_table_check_box.as_ref() }.unwrap();
        assert_eq!(
            display_table_check_box.ulbLpszLabel,
            mem::size_of::<sys::DTBLCHECKBOX>() as u32
        );
        assert_eq!(display_table_check_box.ulFlags, sys::MAPI_UNICODE);
        assert_eq!(
            display_table_check_box.ulPRPropertyName,
            sys::PR_DISPLAY_NAME_W
        );
    }

    #[test]
    fn sized_display_table_group_box_a() {
        const LABEL: &str = "Groupbox Label";

        SizedDtblGroupBox! { DisplayTableGroupBoxA[u8; LABEL.len()] }

        let mut display_table_group_box = DisplayTableGroupBoxA::default();
        let label: Vec<_> = LABEL.bytes().collect();
        assert_eq!(LABEL.len(), label.len());
        display_table_group_box
            .label()
            .copy_from_slice(label.as_slice());
        unsafe {
            assert_eq!(
                PCSTR::from_raw(display_table_group_box.lpszLabel.as_ptr())
                    .to_string()
                    .expect("invalid string"),
                LABEL
            );
        }

        let display_table_group_box: *const sys::DTBLGROUPBOX = display_table_group_box.as_ptr();
        let display_table_group_box = unsafe { display_table_group_box.as_ref() }.unwrap();
        assert_eq!(
            display_table_group_box.ulbLpszLabel,
            mem::size_of::<sys::DTBLGROUPBOX>() as u32
        );
        assert_eq!(display_table_group_box.ulFlags, 0);
    }

    #[test]
    fn sized_display_table_group_box_w() {
        const LABEL: &str = "Groupbox Label";

        SizedDtblGroupBox! { DisplayTableGroupBoxW[u16; LABEL.len()] }

        let mut display_table_group_box = DisplayTableGroupBoxW::default();
        let label: Vec<_> = LABEL.encode_utf16().collect();
        assert_eq!(LABEL.len(), label.len());
        display_table_group_box
            .label()
            .copy_from_slice(label.as_slice());
        unsafe {
            assert_eq!(
                PCWSTR::from_raw(display_table_group_box.lpszLabel.as_ptr())
                    .to_string()
                    .expect("invalid string"),
                LABEL
            );
        }

        let display_table_group_box: *const sys::DTBLGROUPBOX = display_table_group_box.as_ptr();
        let display_table_group_box = unsafe { display_table_group_box.as_ref() }.unwrap();
        assert_eq!(
            display_table_group_box.ulbLpszLabel,
            mem::size_of::<sys::DTBLGROUPBOX>() as u32
        );
        assert_eq!(display_table_group_box.ulFlags, sys::MAPI_UNICODE);
    }

    #[test]
    fn sized_display_table_button_a() {
        const LABEL: &str = "Button Label";

        SizedDtblButton! { DisplayTableButtonA[u8; LABEL.len()] }

        let mut display_table_button = DisplayTableButtonA {
            ulPRControl: sys::PR_DISPLAY_NAME_A,
            ..Default::default()
        };
        let label: Vec<_> = LABEL.bytes().collect();
        assert_eq!(LABEL.len(), label.len());
        display_table_button
            .label()
            .copy_from_slice(label.as_slice());
        unsafe {
            assert_eq!(
                PCSTR::from_raw(display_table_button.lpszLabel.as_ptr())
                    .to_string()
                    .expect("invalid string"),
                LABEL
            );
        }

        let display_table_button: *const sys::DTBLBUTTON = display_table_button.as_ptr();
        let display_table_button = unsafe { display_table_button.as_ref() }.unwrap();
        assert_eq!(
            display_table_button.ulbLpszLabel,
            mem::size_of::<sys::DTBLBUTTON>() as u32
        );
        assert_eq!(display_table_button.ulFlags, 0);
        assert_eq!(display_table_button.ulPRControl, sys::PR_DISPLAY_NAME_A);
    }

    #[test]
    fn sized_display_table_button_w() {
        const LABEL: &str = "Button Label";

        SizedDtblButton! { DisplayTableButtonW[u16; LABEL.len()] }

        let mut display_table_button = DisplayTableButtonW {
            ulPRControl: sys::PR_DISPLAY_NAME_W,
            ..Default::default()
        };
        let label: Vec<_> = LABEL.encode_utf16().collect();
        assert_eq!(LABEL.len(), label.len());
        display_table_button
            .label()
            .copy_from_slice(label.as_slice());
        unsafe {
            assert_eq!(
                PCWSTR::from_raw(display_table_button.lpszLabel.as_ptr())
                    .to_string()
                    .expect("invalid string"),
                LABEL
            );
        }

        let display_table_button: *const sys::DTBLBUTTON = display_table_button.as_ptr();
        let display_table_button = unsafe { display_table_button.as_ref() }.unwrap();
        assert_eq!(
            display_table_button.ulbLpszLabel,
            mem::size_of::<sys::DTBLBUTTON>() as u32
        );
        assert_eq!(display_table_button.ulFlags, sys::MAPI_UNICODE);
        assert_eq!(display_table_button.ulPRControl, sys::PR_DISPLAY_NAME_W);
    }

    #[test]
    fn sized_display_table_page_a() {
        const LABEL: &str = "Page Label";
        const COMPONENT: &str = "Page Component";

        SizedDtblPage! { DisplayTablePageA[u8; LABEL.len(); COMPONENT.len()] }

        let mut display_table_page = DisplayTablePageA {
            ulContext: 10,
            ..Default::default()
        };
        let label: Vec<_> = LABEL.bytes().collect();
        assert_eq!(LABEL.len(), label.len());
        display_table_page.label().copy_from_slice(label.as_slice());
        let component: Vec<_> = COMPONENT.bytes().collect();
        assert_eq!(COMPONENT.len(), component.len());
        display_table_page
            .component()
            .copy_from_slice(component.as_slice());
        unsafe {
            assert_eq!(
                PCSTR::from_raw(display_table_page.lpszLabel.as_ptr())
                    .to_string()
                    .expect("invalid string"),
                LABEL
            );
            assert_eq!(
                PCSTR::from_raw(display_table_page.lpszComponent.as_ptr())
                    .to_string()
                    .expect("invalid string"),
                COMPONENT
            );
        }

        let display_table_page: *const sys::DTBLPAGE = display_table_page.as_ptr();
        let display_table_page = unsafe { display_table_page.as_ref() }.unwrap();
        assert_eq!(
            display_table_page.ulbLpszLabel,
            mem::size_of::<sys::DTBLPAGE>() as u32
        );
        assert_eq!(display_table_page.ulFlags, 0);
        assert_eq!(
            display_table_page.ulbLpszComponent,
            (mem::size_of::<sys::DTBLPAGE>() + mem::size_of::<[u8; LABEL.len() + 1]>()) as u32
        );
        assert_eq!(display_table_page.ulContext, 10);
    }

    #[test]
    fn sized_display_table_page_w() {
        const LABEL: &str = "Page Label";
        const COMPONENT: &str = "Page Component";

        SizedDtblPage! { DisplayTablePageW[u16; LABEL.len(); COMPONENT.len()] }

        let mut display_table_page = DisplayTablePageW {
            ulContext: 10,
            ..Default::default()
        };
        let label: Vec<_> = LABEL.encode_utf16().collect();
        assert_eq!(LABEL.len(), label.len());
        display_table_page.label().copy_from_slice(label.as_slice());
        let component: Vec<_> = COMPONENT.encode_utf16().collect();
        assert_eq!(COMPONENT.len(), component.len());
        display_table_page
            .component()
            .copy_from_slice(component.as_slice());
        unsafe {
            assert_eq!(
                PCWSTR::from_raw(display_table_page.lpszLabel.as_ptr())
                    .to_string()
                    .expect("invalid string"),
                LABEL
            );
            assert_eq!(
                PCWSTR::from_raw(display_table_page.lpszComponent.as_ptr())
                    .to_string()
                    .expect("invalid string"),
                COMPONENT
            );
        }

        let display_table_page: *const sys::DTBLPAGE = display_table_page.as_ptr();
        let display_table_page = unsafe { display_table_page.as_ref() }.unwrap();
        assert_eq!(
            display_table_page.ulbLpszLabel,
            mem::size_of::<sys::DTBLPAGE>() as u32
        );
        assert_eq!(display_table_page.ulFlags, sys::MAPI_UNICODE);
        assert_eq!(
            display_table_page.ulbLpszComponent,
            (mem::size_of::<sys::DTBLPAGE>() + mem::size_of::<[u16; LABEL.len() + 1]>()) as u32
        );
        assert_eq!(display_table_page.ulContext, 10);
    }

    #[test]
    fn sized_display_table_radio_button_a() {
        const LABEL: &str = "Radiobutton Label";

        SizedDtblRadioButton! { DisplayTableRadioButtonA[u8; LABEL.len()] }

        let mut display_table_radio_button = DisplayTableRadioButtonA {
            ulcButtons: 10,
            ulPropTag: sys::PR_DISPLAY_NAME_A,
            lReturnValue: -1,
            ..Default::default()
        };
        let label: Vec<_> = LABEL.bytes().collect();
        assert_eq!(LABEL.len(), label.len());
        display_table_radio_button
            .label()
            .copy_from_slice(label.as_slice());
        unsafe {
            assert_eq!(
                PCSTR::from_raw(display_table_radio_button.lpszLabel.as_ptr())
                    .to_string()
                    .expect("invalid string"),
                LABEL
            );
        }

        let display_table_radio_button: *const sys::DTBLRADIOBUTTON =
            display_table_radio_button.as_ptr();
        assert_eq!(
            unsafe { display_table_radio_button.as_ref() }
                .unwrap()
                .ulFlags,
            0
        );
        let display_table_radio_button = unsafe { display_table_radio_button.as_ref() }.unwrap();
        assert_eq!(
            display_table_radio_button.ulbLpszLabel,
            mem::size_of::<sys::DTBLRADIOBUTTON>() as u32
        );
        assert_eq!(display_table_radio_button.ulFlags, 0);
        assert_eq!(display_table_radio_button.ulcButtons, 10);
        assert_eq!(display_table_radio_button.ulPropTag, sys::PR_DISPLAY_NAME_A);
        assert_eq!(display_table_radio_button.lReturnValue, -1);
    }

    #[test]
    fn sized_display_table_radio_button_w() {
        const LABEL: &str = "Radiobutton Label";

        SizedDtblRadioButton! { DisplayTableRadioButtonW[u16; LABEL.len()] }

        let mut display_table_radio_button = DisplayTableRadioButtonW {
            ulcButtons: 10,
            ulPropTag: sys::PR_DISPLAY_NAME_W,
            lReturnValue: -1,
            ..Default::default()
        };
        let label: Vec<_> = LABEL.encode_utf16().collect();
        assert_eq!(LABEL.len(), label.len());
        display_table_radio_button
            .label()
            .copy_from_slice(label.as_slice());
        unsafe {
            assert_eq!(
                PCWSTR::from_raw(display_table_radio_button.lpszLabel.as_ptr())
                    .to_string()
                    .expect("invalid string"),
                LABEL
            );
        }

        let display_table_radio_button: *const sys::DTBLRADIOBUTTON =
            display_table_radio_button.as_ptr();
        let display_table_radio_button = unsafe { display_table_radio_button.as_ref() }.unwrap();
        assert_eq!(
            display_table_radio_button.ulbLpszLabel,
            mem::size_of::<sys::DTBLRADIOBUTTON>() as u32
        );
        assert_eq!(display_table_radio_button.ulFlags, sys::MAPI_UNICODE);
        assert_eq!(display_table_radio_button.ulcButtons, 10);
        assert_eq!(display_table_radio_button.ulPropTag, sys::PR_DISPLAY_NAME_W);
        assert_eq!(display_table_radio_button.lReturnValue, -1);
    }
}
