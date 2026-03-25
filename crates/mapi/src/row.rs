// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

//! Define [`Row`].

use crate::{PropValue, sys};
use core::{mem, slice};
use std::ptr;

/// Container for the members of a [`sys::SRow`] structure. The [`sys::SPropValue`] pointer should
/// be freed in the destructor with a call to [`sys::MAPIFreeBuffer`].
///
/// Typically, the memory for the [`sys::SRow`] itself is still owned by an [`sys::SRowSet`]
/// allocation, but the [`sys::SRow::lpProps`] member is a separate allocation. [`Row`] copies the
/// [`sys::SRow::cValues`] member and takes ownership of the [`sys::SRow::lpProps`] pointer away
/// from the [`sys::SRow`], leaving both [`sys::SRow`] members empty in the source structure.
pub struct Row {
    count: usize,
    props: *mut sys::SPropValue,
}

impl Row {
    /// Take ownership of the [`sys::SRow`] members.
    pub fn new(row: &mut sys::SRow) -> Self {
        Self {
            count: mem::replace(&mut row.cValues, 0) as usize,
            props: mem::replace(&mut row.lpProps, ptr::null_mut()),
        }
    }

    /// Test for a count of 0 properties or a null [`sys::SPropValue`] pointer.
    pub fn is_empty(&self) -> bool {
        self.count == 0 || self.props.is_null()
    }

    /// Get the number of [`sys::SPropValue`] column values in the [`Row`].
    pub fn len(&self) -> usize {
        if self.props.is_null() { 0 } else { self.count }
    }

    /// Iterate over the [`sys::SPropValue`] column values in the [`Row`].
    pub fn iter(&self) -> impl Iterator<Item = PropValue<'_>> {
        if self.props.is_null() {
            vec![]
        } else {
            unsafe {
                let data: &[sys::SPropValue] = slice::from_raw_parts(self.props, self.count);
                data.iter().map(PropValue::from).collect()
            }
        }
        .into_iter()
    }
}

impl Drop for Row {
    /// Free the [`sys::SPropValue`] pointer with [`sys::MAPIFreeBuffer`].
    fn drop(&mut self) {
        if !self.props.is_null() {
            unsafe {
                sys::MAPIFreeBuffer(self.props as *mut _);
            }
        }
    }
}
