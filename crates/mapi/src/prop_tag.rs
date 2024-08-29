//! Define [`PropTag`] and [`PropType`].

use crate::sys;

pub const PROP_ID_MASK: u32 = 0xFFFF_0000;
pub const PROP_TYPE_MASK: u32 = 0xFFFF;

/// Simple wrapper for a MAPI `PROP_TAG`.
#[repr(transparent)]
#[derive(Clone, Copy)]
pub struct PropTag(pub u32);

impl PropTag {
    /// Combine the `PROP_TYPE` and `PROP_ID` to form a [`PropTag`]. Equivalent to the MAPI
    /// `PROP_TAG` macro.
    pub const fn new(prop_type: PropType, prop_id: u16) -> Self {
        Self(((prop_id as u32) << 16) | (prop_type.0 as u32))
    }

    /// Extract the `PROP_ID` portion of the [`PropTag`]. Equivalent to the MAPI `PROP_ID` macro.
    pub const fn prop_id(&self) -> u16 {
        ((self.0 & PROP_ID_MASK) >> 16) as u16
    }

    /// Extract the `PROP_TYPE` portion of the [`PropTag`]. Equivalent to the MAPI `PROP_TYPE`
    /// macro.
    pub const fn prop_type(&self) -> PropType {
        PropType::new((self.0 & PROP_TYPE_MASK) as u16)
    }

    /// Replace the `PROP_TYPE` portion of the [`PropTag`]. Equalivalent to the MAPI
    /// `CHANGE_PROP_TYPE` macro.
    pub const fn change_prop_type(self, prop_type: PropType) -> Self {
        Self::new(prop_type, self.prop_id())
    }
}

impl From<PropTag> for u32 {
    /// Get a constant `PROP_TAG` value from a [`PropTag`].
    fn from(value: PropTag) -> Self {
        value.0
    }
}

/// Simple wrapper for a MAPI `PROP_TYPE`.
#[repr(transparent)]
#[derive(Clone, Copy)]
pub struct PropType(u16);

impl PropType {
    /// Map invalid property types to [`sys::PT_UNSPECIFIED`].
    pub const fn new(prop_type: u16) -> Self {
        Self(match (prop_type as u32) & !sys::MV_INSTANCE {
            sys::PT_NULL
            | sys::PT_SHORT
            | sys::PT_LONG
            | sys::PT_PTR
            | sys::PT_FLOAT
            | sys::PT_DOUBLE
            | sys::PT_BOOLEAN
            | sys::PT_CURRENCY
            | sys::PT_APPTIME
            | sys::PT_SYSTIME
            | sys::PT_STRING8
            | sys::PT_BINARY
            | sys::PT_UNICODE
            | sys::PT_CLSID
            | sys::PT_LONGLONG
            | sys::PT_MV_SHORT
            | sys::PT_MV_LONG
            | sys::PT_MV_FLOAT
            | sys::PT_MV_DOUBLE
            | sys::PT_MV_CURRENCY
            | sys::PT_MV_APPTIME
            | sys::PT_MV_SYSTIME
            | sys::PT_MV_BINARY
            | sys::PT_MV_STRING8
            | sys::PT_MV_UNICODE
            | sys::PT_MV_CLSID
            | sys::PT_MV_LONGLONG
            | sys::PT_ERROR
            | sys::PT_OBJECT => prop_type,
            _ => sys::PT_UNSPECIFIED as u16,
        })
    }

    /// Set `PROP_TYPE` flags.
    pub const fn add_flags(self, mask: u32) -> Self {
        let mask = (mask & PROP_TYPE_MASK) as u16;
        Self(self.0 | mask)
    }

    /// Clear `PROP_TYPE` flags.
    pub const fn remove_flags(self, mask: u32) -> Self {
        let mask = (mask & PROP_TYPE_MASK) as u16;
        Self(self.0 & !mask)
    }
}

impl From<PropType> for u32 {
    /// Get a constant `PROP_TYPE` value from a [`PropType`].
    fn from(value: PropType) -> Self {
        value.0 as u32
    }
}
