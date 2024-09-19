# Changelog
All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [0.15.1](https://github.com/wravery/outlook-mapi-rs/compare/outlook-mapi-v0.15.0...outlook-mapi-v0.15.1) - 2024-09-19

### Other
- update Cargo.toml dependencies

## [0.15.0](https://github.com/microsoft/mapi-rs/compare/outlook-mapi-v0.14.5...outlook-mapi-v0.15.0) - 2024-09-05

### Other
- *(doc)* update markdown docs
- initial port to microsoft/mapi-rs

## [0.14.5](https://github.com/wravery/mapi-rs/compare/outlook-mapi-v0.14.4...outlook-mapi-v0.14.5) - 2024-08-01

### Other
- *(deps)* update outlook-mapi deps

## [0.14.4](https://github.com/wravery/mapi-rs/compare/outlook-mapi-v0.14.3...outlook-mapi-v0.14.4) - 2024-07-23

### Other
- *(deps)* Update outlook-mapi-sys

## [0.14.3](https://github.com/wravery/mapi-rs/compare/outlook-mapi-v0.14.2...outlook-mapi-v0.14.3) - 2024-07-23

### Other
- *(deps)* Update windows-rs to 0.58

## [0.14.2](https://github.com/wravery/mapi-rs/compare/outlook-mapi-v0.14.1...outlook-mapi-v0.14.2) - 2024-07-16

### Added
- Expose a new function to check for the presence of an Outlook MAPI installation

### Other
- *(deps)* Bump outlook-mapi-sys dependency

## [0.14.1](https://github.com/wravery/mapi-rs/compare/outlook-mapi-v0.14.0...outlook-mapi-v0.14.1) - 2024-06-12

### Fixed
- Update change log for outlook-mapi 0.14.0

### Other
- *(deps)* Update windows-rs to 0.57

## [0.14.0](https://github.com/wravery/mapi-rs/compare/outlook-mapi-v0.13.3...outlook-mapi-v0.14.0) - 2024-05-17

### Fixed
- PropValue creates unaligned pointer references for multi-value types
- Bump the major version of outlook-mapi for incompatible enum variants

### Other
- *(test)* Add tests for PropValue from SPropValue
- Bump MSRV according to new clippy warnings

## [0.13.3](https://github.com/wravery/mapi-rs/compare/outlook-mapi-v0.13.2...outlook-mapi-v0.13.3) - 2024-05-10

### Fixed
- PropTag and PropType should impl Copy

## [0.13.2](https://github.com/wravery/mapi-rs/compare/outlook-mapi-v0.13.1...outlook-mapi-v0.13.2) - 2024-04-12

### Other
- Update windows-rs to 0.56

## [0.13.1](https://github.com/wravery/mapi-rs/compare/outlook-mapi-v0.13.0...outlook-mapi-v0.13.1) - 2024-04-05

### Fixed
- Drop `Logon::session` before `Logon::_initialized` for proper MAPI shutdown

## [0.13.0](https://github.com/wravery/mapi-rs/compare/outlook-mapi-v0.12.1...outlook-mapi-v0.13.0) - 2024-03-21

### Fixed
- `split_off` is less useful and more error prone than `iter`
- Cleanup extra lifetime constraint on `chain(&'a self)`

## [0.12.1](https://github.com/wravery/mapi-rs/compare/outlook-mapi-v0.12.0...outlook-mapi-v0.12.1) - 2024-03-21

### Added
- Implement `MAPIUninit::iter()`

## [0.12.0](https://github.com/wravery/mapi-rs/compare/outlook-mapi-v0.11.3...outlook-mapi-v0.12.0) - 2024-03-21

### Added
- Separate uninit and init states into MAPIUninit and MAPIBuffer for usability

## [0.11.3](https://github.com/wravery/mapi-rs/compare/outlook-mapi-v0.11.2...outlook-mapi-v0.11.3) - 2024-03-20

### Added
- Add MAPIBuffer::get accessor to index into slice offsets without unwrapping

### Other
- minor unit test cleanup
- Merge branch 'main' of https://github.com/wravery/mapi-rs

## [0.11.2](https://github.com/wravery/mapi-rs/compare/outlook-mapi-v0.11.1...outlook-mapi-v0.11.2) - 2024-03-19

### Added
- Allow casting uninitialized MAPIBuffer to another type with `into<P>()`
- Just hide impl macros in rustdoc instead of separate outlook-mapi-macros crate

### Fixed
- Prevent double free in `into<P>(self)` and add mapi_ptr unit tests

## [0.11.1](https://github.com/wravery/mapi-rs/compare/outlook-mapi-v0.11.0...outlook-mapi-v0.11.1) - 2024-03-11

### Fixed
- Simplify MAPIBuffer implementation

## [0.11.0](https://github.com/wravery/mapi-rs/compare/outlook-mapi-v0.10.3...outlook-mapi-v0.11.0) - 2024-03-10

### Added
- Make MAPIBuffer and MAPIOutParam strongly typed

### Other
- Merge branch 'main' of https://github.com/wravery/mapi-rs

## [0.10.3](https://github.com/wravery/mapi-rs/compare/outlook-mapi-v0.10.2...outlook-mapi-v0.10.3) - 2024-03-10

### Other
- Merge branch 'main' of https://github.com/wravery/mapi-rs
- *(test)* Get the categories named prop ID from each store in the sample

## [0.10.2](https://github.com/wravery/mapi-rs/compare/outlook-mapi-v0.10.1...outlook-mapi-v0.10.2) - 2024-03-08

### Added
- Add MAPIBuffer and MAPIOutParam types

## [0.10.1](https://github.com/wravery/mapi-rs/compare/outlook-mapi-v0.10.0...outlook-mapi-v0.10.1) - 2024-03-01

### Other
- update windows-rs dependencies

## [0.10.0](https://github.com/wravery/mapi-rs/compare/outlook-mapi-v0.9.2...outlook-mapi-v0.10.0) - 2024-03-01

### Fixed
- Add stronger PropType validation and distinguish between PT_NULL and PT_OBJECT
- make PropTag #[repr(transparent)]

## [0.9.2](https://github.com/wravery/mapi-rs/compare/outlook-mapi-v0.9.1...outlook-mapi-v0.9.2) - 2024-02-29

### Added
- add PropTag utility

## [0.9.1](https://github.com/wravery/mapi-rs/compare/outlook-mapi-v0.9.0...outlook-mapi-v0.9.1) - 2024-02-29
- Added CHANGELOG.md
- Simplify doc samples by moving most asserts to unit tests
