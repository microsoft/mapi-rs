// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

//! Define [`PropValue`] and [`PropValueData`].

use crate::{sys, PropTag};
use core::{ffi, ptr, slice};
use windows::Win32::{
    Foundation::{E_INVALIDARG, E_POINTER, FILETIME},
    System::Com::CY,
};
use windows_core::*;

/// Wrapper for a [`sys::SPropValue`] structure which allows pattern matching on [`PropValueData`].
pub struct PropValue<'a> {
    pub tag: PropTag,
    pub value: PropValueData<'a>,
}

/// Enum with values from the original [`sys::SPropValue::Value`] union.
pub enum PropValueData<'a> {
    /// [`sys::PT_NULL`]
    Null,

    /// [`sys::PT_I2`] or [`sys::PT_SHORT`]
    Short(i16),

    /// [`sys::PT_I4`] or [`sys::PT_LONG`]
    Long(i32),

    /// [`sys::PT_PTR`] or [`sys::PT_FILE_HANDLE`]
    Pointer(*mut ffi::c_void),

    /// [`sys::PT_R4`] or [`sys::PT_FLOAT`]
    Float(f32),

    /// [`sys::PT_R8`] or [`sys::PT_DOUBLE`]
    Double(f64),

    /// [`sys::PT_BOOLEAN`]
    Boolean(u16),

    /// [`sys::PT_CURRENCY`]
    Currency(i64),

    /// [`sys::PT_APPTIME`]
    AppTime(f64),

    /// [`sys::PT_SYSTIME`]
    FileTime(FILETIME),

    /// [`sys::PT_STRING8`]
    AnsiString(PCSTR),

    /// [`sys::PT_BINARY`]
    Binary(&'a [u8]),

    /// [`sys::PT_UNICODE`]
    Unicode(PCWSTR),

    /// [`sys::PT_CLSID`]
    Guid(GUID),

    /// [`sys::PT_I8`] or [`sys::PT_LONGLONG`]
    LargeInteger(i64),

    /// [`sys::PT_MV_SHORT`]
    ShortArray(&'a [i16]),

    /// [`sys::PT_MV_LONG`]
    LongArray(&'a [i32]),

    /// [`sys::PT_MV_FLOAT`]
    FloatArray(&'a [f32]),

    /// [`sys::PT_MV_DOUBLE`]
    DoubleArray(Vec<f64>),

    /// [`sys::PT_MV_CURRENCY`]
    CurrencyArray(Vec<CY>),

    /// [`sys::PT_MV_APPTIME`]
    AppTimeArray(Vec<f64>),

    /// [`sys::PT_MV_SYSTIME`]
    FileTimeArray(Vec<FILETIME>),

    /// [`sys::PT_MV_BINARY`]
    BinaryArray(Vec<sys::SBinary>),

    /// [`sys::PT_MV_STRING8`]
    AnsiStringArray(Vec<PCSTR>),

    /// [`sys::PT_MV_UNICODE`]
    UnicodeArray(Vec<PCWSTR>),

    /// [`sys::PT_MV_CLSID`]
    GuidArray(Vec<GUID>),

    /// [`sys::PT_MV_LONGLONG`]
    LargeIntegerArray(Vec<i64>),

    /// [`sys::PT_ERROR`]
    Error(HRESULT),

    /// [`sys::PT_OBJECT`]
    Object(i32),
}

impl<'a> From<&'a sys::SPropValue> for PropValue<'a> {
    /// Convert a [`sys::SPropValue`] reference into a friendlier [`PropValue`] type, which often
    /// supports safe access to the [`sys::SPropValue::Value`] union.
    fn from(value: &sys::SPropValue) -> Self {
        let tag = PropTag(value.ulPropTag);
        let prop_type = tag.prop_type().remove_flags(sys::MV_INSTANCE).into();
        let data = unsafe {
            match prop_type {
                sys::PT_NULL => PropValueData::Null,
                sys::PT_SHORT => PropValueData::Short(value.Value.i),
                sys::PT_LONG => PropValueData::Long(value.Value.l),
                sys::PT_PTR => PropValueData::Pointer(value.Value.lpv),
                sys::PT_FLOAT => PropValueData::Float(value.Value.flt),
                sys::PT_DOUBLE => PropValueData::Double(value.Value.dbl),
                sys::PT_BOOLEAN => PropValueData::Boolean(value.Value.b),
                sys::PT_CURRENCY => PropValueData::Currency(value.Value.cur.int64),
                sys::PT_APPTIME => PropValueData::AppTime(value.Value.at),
                sys::PT_SYSTIME => PropValueData::FileTime(value.Value.ft),
                sys::PT_STRING8 => {
                    if value.Value.lpszA.is_null() {
                        PropValueData::Error(E_POINTER)
                    } else {
                        PropValueData::AnsiString(PCSTR::from_raw(value.Value.lpszA.as_ptr()))
                    }
                }
                sys::PT_BINARY => {
                    if value.Value.bin.lpb.is_null() {
                        PropValueData::Error(E_POINTER)
                    } else {
                        PropValueData::Binary(slice::from_raw_parts(
                            value.Value.bin.lpb,
                            value.Value.bin.cb as usize,
                        ))
                    }
                }
                sys::PT_UNICODE => {
                    if value.Value.lpszW.is_null() {
                        PropValueData::Error(E_POINTER)
                    } else {
                        PropValueData::Unicode(PCWSTR::from_raw(value.Value.lpszW.as_ptr()))
                    }
                }
                sys::PT_CLSID => {
                    if value.Value.lpguid.is_null() {
                        PropValueData::Error(E_POINTER)
                    } else {
                        PropValueData::Guid(ptr::read_unaligned(value.Value.lpguid))
                    }
                }
                sys::PT_LONGLONG => PropValueData::LargeInteger(value.Value.li),
                sys::PT_MV_SHORT => {
                    if value.Value.MVi.lpi.is_null() {
                        PropValueData::Error(E_POINTER)
                    } else {
                        PropValueData::ShortArray(slice::from_raw_parts(
                            value.Value.MVi.lpi,
                            value.Value.MVi.cValues as usize,
                        ))
                    }
                }
                sys::PT_MV_LONG => {
                    if value.Value.MVl.lpl.is_null() {
                        PropValueData::Error(E_POINTER)
                    } else {
                        PropValueData::LongArray(slice::from_raw_parts(
                            value.Value.MVl.lpl,
                            value.Value.MVl.cValues as usize,
                        ))
                    }
                }
                sys::PT_MV_FLOAT => {
                    if value.Value.MVflt.lpflt.is_null() {
                        PropValueData::Error(E_POINTER)
                    } else {
                        PropValueData::FloatArray(slice::from_raw_parts(
                            value.Value.MVflt.lpflt,
                            value.Value.MVflt.cValues as usize,
                        ))
                    }
                }
                sys::PT_MV_DOUBLE => {
                    if value.Value.MVdbl.lpdbl.is_null() {
                        PropValueData::Error(E_POINTER)
                    } else {
                        let count = value.Value.MVdbl.cValues as usize;
                        let first = value.Value.MVdbl.lpdbl;
                        let mut values = Vec::with_capacity(count);
                        for idx in 0..count {
                            values.push(ptr::read_unaligned(first.add(idx)))
                        }
                        PropValueData::DoubleArray(values)
                    }
                }
                sys::PT_MV_CURRENCY => {
                    if value.Value.MVcur.lpcur.is_null() {
                        PropValueData::Error(E_POINTER)
                    } else {
                        let count = value.Value.MVcur.cValues as usize;
                        let first = value.Value.MVcur.lpcur;
                        let mut values = Vec::with_capacity(count);
                        for idx in 0..count {
                            values.push(ptr::read_unaligned(first.add(idx)))
                        }
                        PropValueData::CurrencyArray(values)
                    }
                }
                sys::PT_MV_APPTIME => {
                    if value.Value.MVat.lpat.is_null() {
                        PropValueData::Error(E_POINTER)
                    } else {
                        let count = value.Value.MVat.cValues as usize;
                        let first = value.Value.MVat.lpat;
                        let mut values = Vec::with_capacity(count);
                        for idx in 0..count {
                            values.push(ptr::read_unaligned(first.add(idx)))
                        }
                        PropValueData::AppTimeArray(values)
                    }
                }
                sys::PT_MV_SYSTIME => {
                    if value.Value.MVft.lpft.is_null() {
                        PropValueData::Error(E_POINTER)
                    } else {
                        let count = value.Value.MVft.cValues as usize;
                        let first = value.Value.MVft.lpft;
                        let mut values = Vec::with_capacity(count);
                        for idx in 0..count {
                            values.push(ptr::read_unaligned(first.add(idx)))
                        }
                        PropValueData::FileTimeArray(values)
                    }
                }
                sys::PT_MV_BINARY => {
                    if value.Value.MVbin.lpbin.is_null() {
                        PropValueData::Error(E_POINTER)
                    } else {
                        let count = value.Value.MVbin.cValues as usize;
                        let first = value.Value.MVbin.lpbin;
                        let mut values = Vec::with_capacity(count);
                        for idx in 0..count {
                            values.push(ptr::read_unaligned(first.add(idx)))
                        }
                        PropValueData::BinaryArray(values)
                    }
                }
                sys::PT_MV_STRING8 => {
                    if value.Value.MVszA.lppszA.is_null() {
                        PropValueData::Error(E_POINTER)
                    } else {
                        let count = value.Value.MVszA.cValues as usize;
                        let first = value.Value.MVszA.lppszA;
                        let mut values = Vec::with_capacity(count);
                        for idx in 0..count {
                            values.push(PCSTR(ptr::read_unaligned(first.add(idx)).0))
                        }
                        PropValueData::AnsiStringArray(values)
                    }
                }
                sys::PT_MV_UNICODE => {
                    if value.Value.MVszW.lppszW.is_null() {
                        PropValueData::Error(E_POINTER)
                    } else {
                        let count = value.Value.MVszW.cValues as usize;
                        let first = value.Value.MVszW.lppszW;
                        let mut values = Vec::with_capacity(count);
                        for idx in 0..count {
                            values.push(PCWSTR(ptr::read_unaligned(first.add(idx)).0))
                        }
                        PropValueData::UnicodeArray(values)
                    }
                }
                sys::PT_MV_CLSID => {
                    if value.Value.MVguid.lpguid.is_null() {
                        PropValueData::Error(E_POINTER)
                    } else {
                        let count = value.Value.MVguid.cValues as usize;
                        let first = value.Value.MVguid.lpguid;
                        let mut values = Vec::with_capacity(count);
                        for idx in 0..count {
                            values.push(ptr::read_unaligned(first.add(idx)))
                        }
                        PropValueData::GuidArray(values)
                    }
                }
                sys::PT_MV_LONGLONG => {
                    if value.Value.MVli.lpli.is_null() {
                        PropValueData::Error(E_POINTER)
                    } else {
                        let count = value.Value.MVli.cValues as usize;
                        let first = value.Value.MVli.lpli;
                        let mut values = Vec::with_capacity(count);
                        for idx in 0..count {
                            values.push(ptr::read_unaligned(first.add(idx)))
                        }
                        PropValueData::LargeIntegerArray(values)
                    }
                }
                sys::PT_ERROR => PropValueData::Error(HRESULT(value.Value.err)),
                sys::PT_OBJECT => PropValueData::Object(value.Value.x),
                _ => PropValueData::Error(E_INVALIDARG),
            }
        };
        PropValue { tag, value: data }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    use crate::{sys, PropTag, PropType};
    use core::{mem, ptr};
    use windows_core::{s, w};

    #[test]
    fn test_null() {
        let value = sys::SPropValue {
            ulPropTag: sys::PR_NULL,
            ..Default::default()
        };
        let value = PropValue::from(&value);
        assert_eq!(u32::from(value.tag.prop_type()), sys::PT_NULL);
    }

    #[test]
    fn test_short() {
        let mut value = sys::SPropValue {
            ulPropTag: u32::from(
                PropTag(sys::PR_NULL).change_prop_type(PropType::new(sys::PT_I2 as u16)),
            ),
            ..Default::default()
        };
        value.Value.i = 1;
        let value = PropValue::from(&value);
        assert_eq!(u32::from(value.tag.prop_type()), sys::PT_I2);
        assert!(matches!(value.value, PropValueData::Short(1)));
    }

    #[test]
    fn test_long() {
        let mut value = sys::SPropValue {
            ulPropTag: u32::from(
                PropTag(sys::PR_NULL).change_prop_type(PropType::new(sys::PT_I4 as u16)),
            ),
            ..Default::default()
        };
        value.Value.l = 2;
        let value = PropValue::from(&value);
        assert_eq!(u32::from(value.tag.prop_type()), sys::PT_I4);
        assert!(matches!(value.value, PropValueData::Long(2)));
    }

    #[test]
    fn test_pointer() {
        let mut value = sys::SPropValue {
            ulPropTag: u32::from(
                PropTag(sys::PR_NULL).change_prop_type(PropType::new(sys::PT_PTR as u16)),
            ),
            ..Default::default()
        };
        value.Value.lpv = ptr::null_mut();
        let value = PropValue::from(&value);
        assert_eq!(u32::from(value.tag.prop_type()), sys::PT_PTR);
        assert!(matches!(
            value.value,
            PropValueData::Pointer(ptr) if ptr.is_null()
        ));
    }

    #[test]
    fn test_float() {
        let mut value = sys::SPropValue {
            ulPropTag: u32::from(
                PropTag(sys::PR_NULL).change_prop_type(PropType::new(sys::PT_R4 as u16)),
            ),
            ..Default::default()
        };
        value.Value.flt = 3.0;
        let value = PropValue::from(&value);
        assert_eq!(u32::from(value.tag.prop_type()), sys::PT_R4);
        assert!(matches!(value.value, PropValueData::Float(3.0)));
    }

    #[test]
    fn test_double() {
        let mut value = sys::SPropValue {
            ulPropTag: u32::from(
                PropTag(sys::PR_NULL).change_prop_type(PropType::new(sys::PT_R8 as u16)),
            ),
            ..Default::default()
        };
        value.Value.dbl = 4.0;
        let value = PropValue::from(&value);
        assert_eq!(u32::from(value.tag.prop_type()), sys::PT_R8);
        assert!(matches!(value.value, PropValueData::Double(4.0)));
    }

    #[test]
    fn test_boolean() {
        let mut value = sys::SPropValue {
            ulPropTag: u32::from(
                PropTag(sys::PR_NULL).change_prop_type(PropType::new(sys::PT_BOOLEAN as u16)),
            ),
            ..Default::default()
        };
        value.Value.b = 5;
        let value = PropValue::from(&value);
        assert_eq!(u32::from(value.tag.prop_type()), sys::PT_BOOLEAN);
        assert!(matches!(value.value, PropValueData::Boolean(5)));
    }

    #[test]
    fn test_currency() {
        let mut value = sys::SPropValue {
            ulPropTag: u32::from(
                PropTag(sys::PR_NULL).change_prop_type(PropType::new(sys::PT_CURRENCY as u16)),
            ),
            ..Default::default()
        };
        value.Value.cur.int64 = 6;
        let value = PropValue::from(&value);
        assert_eq!(u32::from(value.tag.prop_type()), sys::PT_CURRENCY);
        assert!(matches!(value.value, PropValueData::Currency(6)));
    }

    #[test]
    fn test_app_time() {
        let mut value = sys::SPropValue {
            ulPropTag: u32::from(
                PropTag(sys::PR_NULL).change_prop_type(PropType::new(sys::PT_APPTIME as u16)),
            ),
            ..Default::default()
        };
        value.Value.at = 7.0;
        let value = PropValue::from(&value);
        assert_eq!(u32::from(value.tag.prop_type()), sys::PT_APPTIME);
        assert!(matches!(value.value, PropValueData::AppTime(7.0)));
    }

    #[test]
    fn test_file_time() {
        let mut value = sys::SPropValue {
            ulPropTag: u32::from(
                PropTag(sys::PR_NULL).change_prop_type(PropType::new(sys::PT_SYSTIME as u16)),
            ),
            ..Default::default()
        };
        value.Value.ft.dwLowDateTime = 8;
        let value = PropValue::from(&value);
        assert_eq!(u32::from(value.tag.prop_type()), sys::PT_SYSTIME);
        assert!(matches!(
            value.value,
            PropValueData::FileTime(FILETIME {
                dwHighDateTime: 0,
                dwLowDateTime: 8
            })
        ));
    }

    #[test]
    fn test_ansi_string() {
        let expected = s!("nine");
        let mut value = sys::SPropValue {
            ulPropTag: u32::from(
                PropTag(sys::PR_NULL).change_prop_type(PropType::new(sys::PT_STRING8 as u16)),
            ),
            ..Default::default()
        };
        value.Value.lpszA.0 = expected.0 as *mut _;
        let value = PropValue::from(&value);
        assert_eq!(u32::from(value.tag.prop_type()), sys::PT_STRING8);
        assert!(matches!(
            value.value,
            PropValueData::AnsiString(actual) if actual.0 == expected.0
        ));
    }

    #[test]
    fn test_binary() {
        let expected = 10;
        let mut value = sys::SPropValue {
            ulPropTag: u32::from(
                PropTag(sys::PR_NULL).change_prop_type(PropType::new(sys::PT_BINARY as u16)),
            ),
            ..Default::default()
        };
        value.Value.bin.cb = mem::size_of_val(&expected) as u32;
        value.Value.bin.lpb = &expected as *const i32 as *mut i32 as *mut _;
        let value = PropValue::from(&value);
        assert_eq!(u32::from(value.tag.prop_type()), sys::PT_BINARY);
        assert!(matches!(
            value.value,
            PropValueData::Binary(actual)
                if actual.len() == mem::size_of_val(&expected)
                    && actual.as_ptr() as *const i32 == &expected as *const i32
        ));
    }

    #[test]
    fn test_unicode() {
        let expected = w!("eleven");
        let mut value = sys::SPropValue {
            ulPropTag: u32::from(
                PropTag(sys::PR_NULL).change_prop_type(PropType::new(sys::PT_UNICODE as u16)),
            ),
            ..Default::default()
        };
        value.Value.lpszW.0 = expected.0 as *mut _;
        let value = PropValue::from(&value);
        assert_eq!(u32::from(value.tag.prop_type()), sys::PT_UNICODE);
        assert!(matches!(
            value.value,
            PropValueData::Unicode(actual) if actual.0 == expected.0
        ));
    }

    #[test]
    fn test_guid() {
        let expected = GUID {
            data1: 12,
            ..Default::default()
        };
        let mut value = sys::SPropValue {
            ulPropTag: u32::from(
                PropTag(sys::PR_NULL).change_prop_type(PropType::new(sys::PT_CLSID as u16)),
            ),
            ..Default::default()
        };
        value.Value.lpguid = &expected as *const _ as *mut _;
        let value = PropValue::from(&value);
        assert_eq!(u32::from(value.tag.prop_type()), sys::PT_CLSID);
        assert!(matches!(
            value.value,
            PropValueData::Guid(GUID { data1: 12, .. })
        ));
    }

    #[test]
    fn test_large_integer() {
        let mut value = sys::SPropValue {
            ulPropTag: u32::from(
                PropTag(sys::PR_NULL).change_prop_type(PropType::new(sys::PT_I8 as u16)),
            ),
            ..Default::default()
        };
        value.Value.li = 13;
        let value = PropValue::from(&value);
        assert_eq!(u32::from(value.tag.prop_type()), sys::PT_I8);
        assert!(matches!(value.value, PropValueData::LargeInteger(13)));
    }

    #[test]
    fn test_short_array() {
        let expected = [14_i16, 15];
        let mut value = sys::SPropValue {
            ulPropTag: u32::from(
                PropTag(sys::PR_NULL).change_prop_type(PropType::new(sys::PT_MV_SHORT as u16)),
            ),
            ..Default::default()
        };
        value.Value.MVi.cValues = expected.len() as u32;
        value.Value.MVi.lpi = expected.as_ptr() as *mut _;
        let value = PropValue::from(&value);
        assert_eq!(u32::from(value.tag.prop_type()), sys::PT_MV_SHORT);
        assert!(matches!(value.value, PropValueData::ShortArray([14, 15])));
    }

    #[test]
    fn test_long_array() {
        let expected = [15_i32, 16];
        let mut value = sys::SPropValue {
            ulPropTag: u32::from(
                PropTag(sys::PR_NULL).change_prop_type(PropType::new(sys::PT_MV_LONG as u16)),
            ),
            ..Default::default()
        };
        value.Value.MVl.cValues = expected.len() as u32;
        value.Value.MVl.lpl = expected.as_ptr() as *mut _;
        let value = PropValue::from(&value);
        assert_eq!(u32::from(value.tag.prop_type()), sys::PT_MV_LONG);
        assert!(matches!(value.value, PropValueData::LongArray([15, 16])));
    }

    #[test]
    fn test_float_array() {
        let expected = [16.0_f32, 17.0];
        let mut value = sys::SPropValue {
            ulPropTag: u32::from(
                PropTag(sys::PR_NULL).change_prop_type(PropType::new(sys::PT_MV_FLOAT as u16)),
            ),
            ..Default::default()
        };
        value.Value.MVflt.cValues = expected.len() as u32;
        value.Value.MVflt.lpflt = expected.as_ptr() as *mut _;
        let value = PropValue::from(&value);
        assert_eq!(u32::from(value.tag.prop_type()), sys::PT_MV_FLOAT);
        assert!(matches!(
            value.value,
            PropValueData::FloatArray([16.0, 17.0])
        ));
    }

    #[test]
    fn test_double_array() {
        let expected = [17.0_f64, 18.0];
        let mut value = sys::SPropValue {
            ulPropTag: u32::from(
                PropTag(sys::PR_NULL).change_prop_type(PropType::new(sys::PT_MV_DOUBLE as u16)),
            ),
            ..Default::default()
        };
        value.Value.MVdbl.cValues = expected.len() as u32;
        value.Value.MVdbl.lpdbl = expected.as_ptr() as *mut _;
        let value = PropValue::from(&value);
        assert_eq!(u32::from(value.tag.prop_type()), sys::PT_MV_DOUBLE);
        let PropValueData::DoubleArray(values) = value.value else {
            panic!("wrong type")
        };
        assert!(matches!(values.as_slice(), [17.0, 18.0]));
    }

    #[test]
    fn test_currency_array() {
        let expected = [CY { int64: 18 }, CY { int64: 19 }];
        let mut value = sys::SPropValue {
            ulPropTag: u32::from(
                PropTag(sys::PR_NULL).change_prop_type(PropType::new(sys::PT_MV_CURRENCY as u16)),
            ),
            ..Default::default()
        };
        value.Value.MVcur.cValues = expected.len() as u32;
        value.Value.MVcur.lpcur = expected.as_ptr() as *mut _;
        let value = PropValue::from(&value);
        assert_eq!(u32::from(value.tag.prop_type()), sys::PT_MV_CURRENCY);
        let PropValueData::CurrencyArray(values) = value.value else {
            panic!("wrong type")
        };
        unsafe {
            assert!(matches!(
                values.as_slice(),
                [CY { int64: 18 }, CY { int64: 19 }]
            ));
        }
    }

    #[test]
    fn test_app_time_array() {
        let expected = [19.0_f64, 20.0];
        let mut value = sys::SPropValue {
            ulPropTag: u32::from(
                PropTag(sys::PR_NULL).change_prop_type(PropType::new(sys::PT_MV_APPTIME as u16)),
            ),
            ..Default::default()
        };
        value.Value.MVat.cValues = expected.len() as u32;
        value.Value.MVat.lpat = expected.as_ptr() as *mut _;
        let value = PropValue::from(&value);
        assert_eq!(u32::from(value.tag.prop_type()), sys::PT_MV_APPTIME);
        let PropValueData::AppTimeArray(values) = value.value else {
            panic!("wrong type")
        };
        assert!(matches!(values.as_slice(), [19.0, 20.0]));
    }

    #[test]
    fn test_file_time_array() {
        let expected = [
            FILETIME {
                dwHighDateTime: 20,
                dwLowDateTime: 21,
            },
            FILETIME {
                dwHighDateTime: 22,
                dwLowDateTime: 23,
            },
        ];
        let mut value = sys::SPropValue {
            ulPropTag: u32::from(
                PropTag(sys::PR_NULL).change_prop_type(PropType::new(sys::PT_MV_SYSTIME as u16)),
            ),
            ..Default::default()
        };
        value.Value.MVft.cValues = expected.len() as u32;
        value.Value.MVft.lpft = expected.as_ptr() as *mut _;
        let value = PropValue::from(&value);
        assert_eq!(u32::from(value.tag.prop_type()), sys::PT_MV_SYSTIME);
        let PropValueData::FileTimeArray(values) = value.value else {
            panic!("wrong type")
        };
        assert!(matches!(
            values.as_slice(),
            [
                FILETIME {
                    dwHighDateTime: 20,
                    dwLowDateTime: 21,
                },
                FILETIME {
                    dwHighDateTime: 22,
                    dwLowDateTime: 23,
                }
            ]
        ));
    }

    #[test]
    fn test_binary_array() {
        let expected1 = [24_u8, 25_u8];
        let expected1 = sys::SBinary {
            cb: expected1.len() as u32,
            lpb: expected1.as_ptr() as *mut _,
        };
        let expected2 = [26_u8, 27_u8];
        let expected2 = sys::SBinary {
            cb: expected2.len() as u32,
            lpb: expected2.as_ptr() as *mut _,
        };
        let expected = [expected1, expected2];
        let mut value = sys::SPropValue {
            ulPropTag: u32::from(
                PropTag(sys::PR_NULL).change_prop_type(PropType::new(sys::PT_MV_BINARY as u16)),
            ),
            ..Default::default()
        };
        value.Value.MVbin.cValues = expected.len() as u32;
        value.Value.MVbin.lpbin = expected.as_ptr() as *mut _;
        let value = PropValue::from(&value);
        assert_eq!(u32::from(value.tag.prop_type()), sys::PT_MV_BINARY);
        let PropValueData::BinaryArray(values) = value.value else {
            panic!("wrong type")
        };
        assert!(matches!(
            values.as_slice(),
            [actual1, actual2]
                if actual1.cb == expected[0].cb && actual1.lpb == expected[0].lpb
                    && actual2.cb == expected[1].cb && actual2.lpb == expected[1].lpb
        ));
    }

    #[test]
    fn test_ansi_string_array() {
        let expected = [s!("twenty-eight"), s!("twenty-nine")];
        let mut value = sys::SPropValue {
            ulPropTag: u32::from(
                PropTag(sys::PR_NULL).change_prop_type(PropType::new(sys::PT_MV_STRING8 as u16)),
            ),
            ..Default::default()
        };
        value.Value.MVszA.cValues = expected.len() as u32;
        value.Value.MVszA.lppszA = expected.as_ptr() as *mut _;
        let value = PropValue::from(&value);
        assert_eq!(u32::from(value.tag.prop_type()), sys::PT_MV_STRING8);
        let PropValueData::AnsiStringArray(values) = value.value else {
            panic!("wrong type")
        };
        assert!(matches!(
            values.as_slice(),
            [actual1, actual2]
                if actual1.0 == expected[0].0 && actual2.0 == expected[1].0
        ));
    }

    #[test]
    fn test_unicode_string_array() {
        let expected = [w!("thirty"), w!("thirty-one")];
        let mut value = sys::SPropValue {
            ulPropTag: u32::from(
                PropTag(sys::PR_NULL).change_prop_type(PropType::new(sys::PT_MV_UNICODE as u16)),
            ),
            ..Default::default()
        };
        value.Value.MVszW.cValues = expected.len() as u32;
        value.Value.MVszW.lppszW = expected.as_ptr() as *mut _;
        let value = PropValue::from(&value);
        assert_eq!(u32::from(value.tag.prop_type()), sys::PT_MV_UNICODE);
        let PropValueData::UnicodeArray(values) = value.value else {
            panic!("wrong type")
        };
        assert!(matches!(
            values.as_slice(),
            [actual1, actual2]
                if actual1.0 == expected[0].0 && actual2.0 == expected[1].0
        ));
    }

    #[test]
    fn test_guid_array() {
        let expected = [
            GUID {
                data1: 32,
                ..Default::default()
            },
            GUID {
                data2: 33,
                ..Default::default()
            },
            GUID {
                data3: 34,
                ..Default::default()
            },
            GUID {
                data4: [35, 0, 0, 0, 0, 0, 0, 0],
                ..Default::default()
            },
        ];
        let mut value = sys::SPropValue {
            ulPropTag: u32::from(
                PropTag(sys::PR_NULL).change_prop_type(PropType::new(sys::PT_MV_CLSID as u16)),
            ),
            ..Default::default()
        };
        value.Value.MVguid.cValues = expected.len() as u32;
        value.Value.MVguid.lpguid = expected.as_ptr() as *mut _;
        let value = PropValue::from(&value);
        assert_eq!(u32::from(value.tag.prop_type()), sys::PT_MV_CLSID);
        let PropValueData::GuidArray(values) = value.value else {
            panic!("wrong type")
        };
        assert!(matches!(
            values.as_slice(),
            [
                GUID { data1: 32, .. },
                GUID { data2: 33, .. },
                GUID { data3: 34, .. },
                GUID {
                    data4: [35, ..],
                    ..
                }
            ]
        ));
    }

    #[test]
    fn test_large_integer_array() {
        let expected = [36_i64, 37];
        let mut value = sys::SPropValue {
            ulPropTag: u32::from(
                PropTag(sys::PR_NULL).change_prop_type(PropType::new(sys::PT_MV_LONGLONG as u16)),
            ),
            ..Default::default()
        };
        value.Value.MVli.cValues = expected.len() as u32;
        value.Value.MVli.lpli = expected.as_ptr() as *mut _;
        let value = PropValue::from(&value);
        assert_eq!(u32::from(value.tag.prop_type()), sys::PT_MV_LONGLONG);
        let PropValueData::LargeIntegerArray(values) = value.value else {
            panic!("wrong type")
        };
        assert!(matches!(values.as_slice(), [36, 37]));
    }

    #[test]
    fn test_error() {
        let mut value = sys::SPropValue {
            ulPropTag: u32::from(
                PropTag(sys::PR_NULL).change_prop_type(PropType::new(sys::PT_ERROR as u16)),
            ),
            ..Default::default()
        };
        value.Value.err = 38;
        let value = PropValue::from(&value);
        assert_eq!(u32::from(value.tag.prop_type()), sys::PT_ERROR);
        assert!(matches!(value.value, PropValueData::Error(HRESULT(38))));
    }

    #[test]
    fn test_object() {
        let mut value = sys::SPropValue {
            ulPropTag: u32::from(
                PropTag(sys::PR_NULL).change_prop_type(PropType::new(sys::PT_OBJECT as u16)),
            ),
            ..Default::default()
        };
        value.Value.x = 39;
        let value = PropValue::from(&value);
        assert_eq!(u32::from(value.tag.prop_type()), sys::PT_OBJECT);
        assert!(matches!(value.value, PropValueData::Object(39)));
    }
}
