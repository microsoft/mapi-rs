// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

//! Define [`RowSet`].

use crate::{Row, sys};
use core::{ptr, slice};

/// Container for a [`sys::SRowSet`] structure, such as the rows returned from
/// [`sys::IMAPITable::QueryRows`].
///
/// The `*mut sys::SRowSet` should be freed with a call to [`sys::FreeProws`] in the destructor,
/// but [`sys::SRowSet`] allows embedded [`sys::SPropValue`] pointers to be peeled off and freed
/// separately, as long as the pointer in the [`sys::SRowSet`] is replaced with `null`. The
/// [`sys::FreeProws`] function will free any non-`null` property pointers in the [`sys::SRowSet`],
/// but silently skip the ones that are `null`.
pub struct RowSet {
    rows: *mut sys::SRowSet,
}

impl RowSet {
    /// Get an out-param pointer for the [`sys::SRowSet`] pointer.
    pub fn as_mut_ptr(&mut self) -> *mut *mut sys::SRowSet {
        &mut self.rows
    }

    /// Test for a `null` [`sys::SRowSet`] pointer or a pointer to 0 rows.
    pub fn is_empty(&self) -> bool {
        unsafe {
            self.rows
                .as_ref()
                .map(|rows| rows.cRows == 0)
                .unwrap_or(true)
        }
    }

    /// Get the count of rows contained in the [`sys::SRowSet`].
    pub fn len(&self) -> usize {
        unsafe {
            self.rows
                .as_ref()
                .map(|rows| rows.cRows as usize)
                .unwrap_or_default()
        }
    }
}

impl Default for RowSet {
    /// The initial state for [`RowSet`] should have a `null` `*mut sys::SRowSet` pointer. An
    /// allocation should be added to it later using the [`RowSet::as_mut_ptr`] method to fill in
    /// an out-param from one of the [`sys`] functions or interface methods which retrieve a
    /// [`sys::SRowSet`] structure.
    fn default() -> Self {
        Self {
            rows: ptr::null_mut(),
        }
    }
}

impl IntoIterator for RowSet {
    type Item = Row;
    type IntoIter = <Vec<Self::Item> as IntoIterator>::IntoIter;

    /// Transfer ownership of the embedded [`sys::SPropValue`] pointers to an [`Iterator`] of
    /// [`Row`].
    fn into_iter(self) -> Self::IntoIter {
        unsafe {
            if let Some(rows) = self.rows.as_mut() {
                let count = rows.cRows as usize;
                let data: &mut [sys::SRow] =
                    slice::from_raw_parts_mut(rows.aRow.as_mut_ptr(), count);
                data.iter_mut().map(Row::new).collect()
            } else {
                vec![]
            }
        }
        .into_iter()
    }
}

impl Drop for RowSet {
    /// Call [`sys::FreeProws`] to free the `*mut sys::SRowSet`. This will also free any
    /// [`sys::SPropValue`] pointers that have not been transfered to an instance of [`Row`].
    fn drop(&mut self) {
        if !self.rows.is_null() {
            unsafe {
                sys::FreeProws(self.rows);
            }
        }
    }
}
