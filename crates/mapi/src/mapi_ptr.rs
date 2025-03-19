// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

//! Define [`MAPIUninit`], [`MAPIBuffer`], and [`MAPIOutParam`].
//!
//! Smart pointer types for memory allocated with [`sys::MAPIAllocateBuffer`], which must be freed
//! with [`sys::MAPIFreeBuffer`], or [`sys::MAPIAllocateMore`], which is chained to another
//! allocation and must not outlive that allocation or be separately freed.

use crate::sys;
use core::{
    ffi,
    marker::PhantomData,
    mem::{self, MaybeUninit},
    ptr, slice,
};
use windows::Win32::Foundation::E_OUTOFMEMORY;
use windows_core::{Error, HRESULT};

/// Errors which can be returned from this module.
#[derive(Debug)]
pub enum MAPIAllocError {
    /// The underlying [`sys::MAPIAllocateBuffer`] and [`sys::MAPIAllocateMore`] take a `u32`
    /// parameter specifying the size of the buffer. If you exceed the capacity of a `u32`, it will
    /// be impossible to make the necessary allocation.
    SizeOverflow(usize),

    /// MAPI APIs like to work with input and output buffers using `*const u8` and `*mut u8` rather
    /// than strongly typed pointers. In C++, this requires a lot of `reinterpret_cast` operations.
    /// For the accessors on this type, we'll allow transmuting the buffer to the desired type, as
    /// long as it fits in the allocation. If you don't allocate enough space for the number of
    /// elements you are accessing, it will return this error.
    OutOfBoundsAccess,

    /// There are no [documented](https://learn.microsoft.com/en-us/office/client-developer/outlook/mapi/mapiallocatebuffer)
    /// conditions where [`sys::MAPIAllocateBuffer`] or [`sys::MAPIAllocateMore`] will return an
    /// error, but if they do fail, this will propagate the [`Error`] result. If the allocation
    /// function returns `null` with no other error, it will treat that as [`E_OUTOFMEMORY`].
    AllocationFailed(Error),
}

enum Buffer<T>
where
    T: Sized,
{
    Uninit(*mut MaybeUninit<T>),
    Ready(*mut T),
}

enum Allocation<'a, T>
where
    T: Sized,
{
    Root {
        buffer: Buffer<T>,
        byte_count: usize,
    },
    More {
        buffer: Buffer<T>,
        byte_count: usize,
        root: *mut ffi::c_void,
        phantom: PhantomData<&'a T>,
    },
}

impl<'a, T> Allocation<'a, T>
where
    T: Sized,
{
    fn new(count: usize) -> Result<Self, MAPIAllocError> {
        let byte_count = count * mem::size_of::<T>();
        Ok(Self::Root {
            buffer: unsafe {
                let mut alloc = ptr::null_mut();
                HRESULT::from_win32(sys::MAPIAllocateBuffer(
                    u32::try_from(byte_count)
                        .map_err(|_| MAPIAllocError::SizeOverflow(byte_count))?,
                    &mut alloc,
                ) as u32)
                .ok()
                .map_err(MAPIAllocError::AllocationFailed)?;
                if alloc.is_null() {
                    return Err(MAPIAllocError::AllocationFailed(Error::from_hresult(
                        E_OUTOFMEMORY,
                    )));
                }
                Buffer::Uninit(alloc as *mut _)
            },
            byte_count,
        })
    }

    fn chain<P>(&self, count: usize) -> Result<Allocation<'a, P>, MAPIAllocError>
    where
        P: Sized,
    {
        let root = match self {
            Self::Root { buffer, .. } => match buffer {
                Buffer::Uninit(alloc) => *alloc as *mut _,
                Buffer::Ready(alloc) => *alloc as *mut _,
            },
            Self::More { root, .. } => *root,
        };
        let byte_count = count * mem::size_of::<T>();
        Ok(Allocation::More {
            buffer: unsafe {
                let mut alloc = ptr::null_mut();
                HRESULT::from_win32(sys::MAPIAllocateMore(
                    u32::try_from(byte_count)
                        .map_err(|_| MAPIAllocError::SizeOverflow(byte_count))?,
                    root,
                    &mut alloc,
                ) as u32)
                .ok()
                .map_err(MAPIAllocError::AllocationFailed)?;
                if alloc.is_null() {
                    return Err(MAPIAllocError::AllocationFailed(Error::from_hresult(
                        E_OUTOFMEMORY,
                    )));
                }
                Buffer::Uninit(alloc as *mut _)
            },
            byte_count,
            root,
            phantom: PhantomData,
        })
    }

    fn into<P>(self) -> Result<Allocation<'a, P>, MAPIAllocError> {
        let result = match self {
            Self::Root {
                buffer: Buffer::Ready(_),
                ..
            }
            | Self::More {
                buffer: Buffer::Ready(_),
                ..
            } => unreachable!(),
            Self::Root {
                buffer: Buffer::Uninit(alloc),
                byte_count,
            } if byte_count >= mem::size_of::<T>() => Ok(Allocation::Root {
                buffer: Buffer::Uninit(alloc as *mut _),
                byte_count,
            }),
            Self::More {
                buffer: Buffer::Uninit(alloc),
                byte_count,
                root,
                ..
            } if byte_count >= mem::size_of::<T>() => Ok(Allocation::More {
                buffer: Buffer::Uninit(alloc as *mut _),
                byte_count,
                root,
                phantom: PhantomData,
            }),
            _ => Err(MAPIAllocError::OutOfBoundsAccess),
        };
        if result.is_ok() {
            mem::forget(self);
        }
        result
    }

    fn iter(&self) -> AllocationIter<'a, T> {
        match self {
            Self::Root {
                buffer: Buffer::Uninit(alloc),
                byte_count,
            } => AllocationIter {
                alloc: *alloc,
                byte_count: *byte_count,
                element_size: mem::size_of::<T>(),
                root: *alloc as *mut _,
                phantom: PhantomData,
            },
            Self::More {
                buffer: Buffer::Uninit(alloc),
                byte_count,
                root,
                ..
            } => AllocationIter {
                alloc: *alloc,
                byte_count: *byte_count,
                element_size: mem::size_of::<T>(),
                root: *root,
                phantom: PhantomData,
            },
            _ => unreachable!(),
        }
    }

    fn uninit(&mut self) -> Result<&mut MaybeUninit<T>, MAPIAllocError> {
        match self {
            Self::Root {
                buffer: Buffer::Ready(_),
                ..
            }
            | Self::More {
                buffer: Buffer::Ready(_),
                ..
            } => unreachable!(),
            Self::Root {
                buffer: Buffer::Uninit(alloc),
                byte_count,
            } if mem::size_of::<T>() <= *byte_count => Ok(unsafe { &mut *(*alloc) }),
            Self::More {
                buffer: Buffer::Uninit(alloc),
                byte_count,
                ..
            } if mem::size_of::<T>() <= *byte_count => Ok(unsafe { &mut *(*alloc) }),
            _ => Err(MAPIAllocError::OutOfBoundsAccess),
        }
    }

    unsafe fn assume_init(self) -> Self {
        let result = match self {
            Self::Root {
                buffer: Buffer::Uninit(alloc),
                byte_count,
            } => Self::Root {
                buffer: Buffer::Ready(alloc as *mut _),
                byte_count,
            },
            Self::More {
                buffer: Buffer::Uninit(alloc),
                byte_count,
                root,
                ..
            } => Self::More {
                buffer: Buffer::Ready(alloc as *mut _),
                byte_count,
                root,
                phantom: PhantomData,
            },
            _ => unreachable!(),
        };
        mem::forget(self);
        result
    }

    fn as_mut(&mut self) -> Result<&mut T, MAPIAllocError> {
        match self {
            Self::Root {
                buffer: Buffer::Uninit(_),
                ..
            }
            | Self::More {
                buffer: Buffer::Uninit(_),
                ..
            } => unreachable!(),
            Self::Root {
                buffer: Buffer::Ready(alloc),
                byte_count,
            } if mem::size_of::<T>() <= *byte_count => Ok(unsafe { &mut *(*alloc) }),
            Self::More {
                buffer: Buffer::Ready(alloc),
                byte_count,
                ..
            } if mem::size_of::<T>() <= *byte_count => Ok(unsafe { &mut *(*alloc) }),
            _ => Err(MAPIAllocError::OutOfBoundsAccess),
        }
    }
}

impl<T> Drop for Allocation<'_, T> {
    fn drop(&mut self) {
        if let Self::Root { buffer, .. } = self {
            let alloc = match mem::replace(buffer, Buffer::Uninit(ptr::null_mut())) {
                Buffer::Uninit(alloc) => alloc as *mut T,
                Buffer::Ready(alloc) => alloc,
            };
            if !alloc.is_null() {
                #[cfg(test)]
                unreachable!();
                #[cfg(not(test))]
                unsafe {
                    sys::MAPIFreeBuffer(alloc as *mut _);
                }
            }
        }
    }
}

struct AllocationIter<'a, T>
where
    T: Sized,
{
    alloc: *mut MaybeUninit<T>,
    byte_count: usize,
    root: *mut ffi::c_void,
    element_size: usize,
    phantom: PhantomData<&'a T>,
}

impl<'a, T> Iterator for AllocationIter<'a, T>
where
    T: Sized,
{
    type Item = Allocation<'a, T>;

    fn next(&mut self) -> Option<Self::Item> {
        if self.byte_count < self.element_size {
            return None;
        }

        let item = Allocation::More {
            buffer: Buffer::Uninit(self.alloc),
            byte_count: self.element_size,
            root: self.root,
            phantom: PhantomData,
        };

        self.byte_count -= self.element_size;
        self.alloc = unsafe { self.alloc.add(1) };

        Some(item)
    }
}

/// Wrapper type for an allocation with either [`sys::MAPIAllocateBuffer`] or
/// [`sys::MAPIAllocateMore`] which has not been initialized yet.
pub struct MAPIUninit<'a, T>(Allocation<'a, T>)
where
    T: Sized;

impl<'a, T> MAPIUninit<'a, T> {
    /// Create a new allocation with enough room for `count` elements of type `T` with a call to
    /// [`sys::MAPIAllocateBuffer`]. The buffer is freed as soon as the [`MAPIUninit`] or
    /// [`MAPIBuffer`] is dropped.
    ///
    /// If you call [`MAPIUninit::chain`] to create any more allocations with
    /// [`sys::MAPIAllocateMore`], their lifetimes are constrained to the lifetime of this
    /// allocation and they will all be freed together in a single call to [`sys::MAPIFreeBuffer`].
    pub fn new(count: usize) -> Result<Self, MAPIAllocError> {
        Ok(Self(Allocation::new(count)?))
    }

    /// Create a new allocation with enough room for `count` elements of type `P` with a call to
    /// [`sys::MAPIAllocateMore`]. The result is a separate allocation that is not freed until
    /// `self` is dropped at the beginning of the chain.
    ///
    /// You may call [`MAPIUninit::chain`] on the result as well, they will both share a root
    /// allocation created with [`MAPIUninit::new`].
    pub fn chain<P>(&self, count: usize) -> Result<MAPIUninit<'a, P>, MAPIAllocError> {
        Ok(MAPIUninit::<'a, P>(self.0.chain::<P>(count)?))
    }

    /// Convert an uninitialized allocation to another type. You can use this, for example, to
    /// perform an allocation with extra space in a `&mut [u8]` buffer, and then cast that to a
    /// specific type. This is useful with the `CbNewXXX` functions in [`crate::sized_types`].
    pub fn into<P>(self) -> Result<MAPIUninit<'a, P>, MAPIAllocError> {
        Ok(MAPIUninit::<'a, P>(self.0.into::<P>()?))
    }

    /// Get an iterator over the uninitialized elements.
    pub fn iter(&self) -> MAPIUninitIter<'a, T> {
        MAPIUninitIter(self.0.iter())
    }

    /// Get an uninitialized out-parameter with enough room for a single element of type `T`.
    pub fn uninit(&mut self) -> Result<&mut MaybeUninit<T>, MAPIAllocError> {
        self.0.uninit()
    }

    /// Once the buffer is known to be completely filled in, convert this [`MAPIUninit`] to a
    /// fully initialized [`MAPIBuffer`].
    ///
    /// # Safety
    ///
    /// Like [`MaybeUninit`], the caller must ensure that the buffer is completely initialized
    /// before calling [`MAPIUninit::assume_init`]. It is undefined behavior to leave it
    /// uninitialized once we start accessing it.
    pub unsafe fn assume_init(self) -> MAPIBuffer<'a, T> {
        MAPIBuffer(unsafe { self.0.assume_init() })
    }
}

/// Iterator over the uninitialized elements in a [`MAPIUninit`] allocation.
pub struct MAPIUninitIter<'a, T>(AllocationIter<'a, T>)
where
    T: Sized;

impl<'a, T> Iterator for MAPIUninitIter<'a, T>
where
    T: Sized,
{
    type Item = MAPIUninit<'a, T>;

    fn next(&mut self) -> Option<Self::Item> {
        self.0.next().map(MAPIUninit)
    }
}

/// Wrapper type for an allocation in [`MAPIUninit`] which has been fully initialized.
pub struct MAPIBuffer<'a, T>(Allocation<'a, T>)
where
    T: Sized;

impl<'a, T> MAPIBuffer<'a, T> {
    /// Create a new allocation with enough room for `count` elements of type `P` with a call to
    /// [`sys::MAPIAllocateMore`]. The result is a separate allocation that is not freed until
    /// `self` is dropped at the beginning of the chain.
    ///
    /// You may call [`MAPIBuffer::chain`] on the result as well, they will both share a root
    /// allocation created with [`MAPIUninit::new`].
    pub fn chain<P>(&self, count: usize) -> Result<MAPIUninit<'a, P>, MAPIAllocError> {
        Ok(MAPIUninit::<'a, P>(self.0.chain::<P>(count)?))
    }

    /// Access a single element of type `T` once it has been initialized with
    /// [`MAPIUninit::assume_init`].
    pub fn as_mut(&mut self) -> Result<&mut T, MAPIAllocError> {
        self.0.as_mut()
    }
}

/// Hold an out-pointer for MAPI APIs which perform their own buffer allocations. This version does
/// not perform any validation of the buffer size, so the typed accessors are inherently unsafe.
pub struct MAPIOutParam<T>(*mut T)
where
    T: Sized;

impl<T> MAPIOutParam<T>
where
    T: Sized,
{
    /// Get a `*mut *mut T` suitable for use with a MAPI API that fills in an out-pointer
    /// with a newly allocated buffer.
    pub fn as_mut_ptr(&mut self) -> *mut *mut T {
        &mut self.0
    }

    /// Access a single element of type `T`.
    ///
    /// # Safety
    ///
    /// This version does not perform any validation of the buffer size, so the typed accessors are
    /// inherently unsafe. The only thing it handles is a `null` check.
    pub unsafe fn as_mut(&mut self) -> Option<&mut T> {
        unsafe { self.0.as_mut() }
    }

    /// Access a slice with `count` elements of type `T`.
    ///
    /// # Safety
    ///
    /// This version does not perform any validation of the buffer size, so the typed accessors are
    /// inherently unsafe. The only thing it handles is a `null` check.
    pub unsafe fn as_mut_slice(&mut self, count: usize) -> Option<&mut [T]> {
        if self.0.is_null() {
            None
        } else {
            Some(unsafe { slice::from_raw_parts_mut(self.0, count) })
        }
    }
}

impl<T> Default for MAPIOutParam<T>
where
    T: Sized,
{
    fn default() -> Self {
        Self(ptr::null_mut())
    }
}

impl<T> Drop for MAPIOutParam<T>
where
    T: Sized,
{
    fn drop(&mut self) {
        if !self.0.is_null() {
            #[cfg(test)]
            unreachable!();
            #[cfg(not(test))]
            unsafe {
                sys::MAPIFreeBuffer(self.0 as *mut _);
            }
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::*;

    use mem::ManuallyDrop;

    SizedSPropTagArray! { TestTags[2] }

    const TEST_TAGS: TestTags = TestTags {
        cValues: 2,
        aulPropTag: [sys::PR_INSTANCE_KEY, sys::PR_SUBJECT_W],
    };

    #[test]
    fn buffer_uninit() {
        let mut buffer: MaybeUninit<TestTags> = MaybeUninit::uninit();
        let mut mapi_buffer = ManuallyDrop::new(MAPIUninit(Allocation::Root {
            buffer: Buffer::Uninit(&mut buffer),
            byte_count: mem::size_of_val(&buffer),
        }));
        assert!(mapi_buffer.uninit().is_ok());
    }

    #[test]
    fn buffer_into() {
        let mut buffer: [MaybeUninit<u8>; mem::size_of::<TestTags>()] =
            [MaybeUninit::uninit(); CbNewSPropTagArray(2)];
        let mut mapi_buffer = ManuallyDrop::new(MAPIUninit(Allocation::Root {
            buffer: Buffer::Uninit(buffer.as_mut_ptr()),
            byte_count: buffer.len(),
        }));
        assert!(mapi_buffer.uninit().is_ok());
        let mut mapi_buffer = ManuallyDrop::new(
            ManuallyDrop::into_inner(mapi_buffer)
                .into::<TestTags>()
                .expect("into failed"),
        );
        assert!(mapi_buffer.uninit().is_ok());
    }

    #[test]
    fn buffer_iter() {
        let mut buffer: [MaybeUninit<u32>; 2] = [MaybeUninit::uninit(); 2];
        let mapi_buffer = ManuallyDrop::new(MAPIUninit(Allocation::Root {
            buffer: Buffer::Uninit(buffer.as_mut_ptr()),
            byte_count: buffer.len() * mem::size_of::<u32>(),
        }));
        let mut next = mapi_buffer.iter();
        assert!(match next.next() {
            Some(MAPIUninit(Allocation::More {
                buffer: Buffer::Uninit(alloc),
                byte_count,
                root,
                ..
            })) => {
                assert_eq!(alloc, buffer.as_mut_ptr() as *mut _);
                assert_eq!(root, buffer.as_mut_ptr() as *mut _);
                assert_eq!(byte_count, mem::size_of::<u32>());
                true
            }
            _ => false,
        });
        assert!(match next.next() {
            Some(MAPIUninit(Allocation::More {
                buffer: Buffer::Uninit(alloc),
                byte_count,
                root,
                ..
            })) => {
                assert_eq!(alloc, unsafe { buffer.as_mut_ptr().add(1) } as *mut _);
                assert_eq!(root, buffer.as_mut_ptr() as *mut _);
                assert_eq!(byte_count, mem::size_of::<u32>());
                true
            }
            _ => false,
        });
        assert!(next.next().is_none());
    }

    #[test]
    fn buffer_assume_init() {
        let mut buffer = MaybeUninit::uninit();
        let mapi_buffer = ManuallyDrop::new(MAPIUninit(Allocation::Root {
            buffer: Buffer::Uninit(&mut buffer),
            byte_count: mem::size_of_val(&buffer),
        }));
        buffer.write(TEST_TAGS);
        let mut mapi_buffer =
            ManuallyDrop::new(unsafe { ManuallyDrop::into_inner(mapi_buffer).assume_init() });
        let test_tags = mapi_buffer.as_mut().expect("as_mut failed");
        assert_eq!(TEST_TAGS.cValues, test_tags.cValues);
        assert_eq!(TEST_TAGS.aulPropTag, test_tags.aulPropTag);
    }
}
