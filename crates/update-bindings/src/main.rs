extern crate windows_bindgen;

fn main() -> Result<()> {
    if mapi_bindgen::update_mapi_sys()? {
        println!("Microsoft.rs changed");
    }

    Ok(())
}

#[derive(Debug, Error)]
pub enum Error {
    #[error("Missing Parent")]
    MissingParent(std::path::PathBuf),
    #[error(transparent)]
    Io(#[from] std::io::Error),
    #[error(transparent)]
    Regex(#[from] regex::Error),
}

pub type Result<T> = std::result::Result<T, Error>;

#[macro_use]
extern crate thiserror;

mod mapi_path {
    use std::{convert::From, path::PathBuf};

    pub fn get_out_dir() -> PathBuf {
        PathBuf::from(env!("OUT_DIR"))
    }

    pub fn get_manifest_dir() -> PathBuf {
        PathBuf::from(env!("CARGO_MANIFEST_DIR"))
    }

    pub fn get_mapi_sys_dir() -> super::Result<PathBuf> {
        let manifest_dir = get_manifest_dir();
        let mut mapi_sys_dir = get_manifest_dir().parent().map_or_else(
            || Err(super::Error::MissingParent(manifest_dir)),
            |parent| Ok(PathBuf::from(parent)),
        )?;
        mapi_sys_dir.push("mapi-sys");
        Ok(mapi_sys_dir)
    }
}

mod mapi_bindgen {
    use std::{
        fs,
        io::{Read, Write},
        path::{Path, PathBuf},
    };

    use regex::Regex;

    use windows_bindgen::bindgen;

    use super::mapi_path::*;

    pub fn update_mapi_sys() -> super::Result<bool> {
        let source_path = generate_mapi_sys()?;
        format_mapi_sys(&source_path)?;
        let source = read_mapi_sys(&source_path)?;

        let mut dest_path = get_mapi_sys_dir()?;
        dest_path.push("src");
        dest_path.push("Microsoft.rs");
        let dest = read_mapi_sys(&dest_path)?;

        if source != dest {
            fs::copy(&source_path, &dest_path)?;
            Ok(true)
        } else {
            Ok(false)
        }
    }

    fn generate_mapi_sys() -> super::Result<PathBuf> {
        const WINMD_FILE: &str = "Microsoft.Office.Outlook.MAPI.Win32.winmd";

        let mut winmd_path = get_manifest_dir();
        winmd_path.push("winmd");
        winmd_path.push(WINMD_FILE);
        let mut source_path = get_out_dir();
        source_path.push("Microsoft.rs");
        println!(
            "{}",
            bindgen([
                "--in",
                winmd_path.to_str().expect("invalid winmd path"),
                "--out",
                source_path.to_str().expect("invalid Microsoft.rs path"),
                "--filter",
                "Microsoft.Office.Outlook.MAPI.Win32",
                "--config",
                "implement",
            ])
            .expect("bindgen failed")
        );

        let mut outlook_mapi_sys = Default::default();
        fs::File::open(source_path.clone())?.read_to_string(&mut outlook_mapi_sys)?;

        let mut source_file = fs::File::create(source_path.clone())?;

        source_file.write_all(patch_mapi_sys(outlook_mapi_sys)?.as_bytes())?;
        Ok(source_path)
    }

    fn patch_mapi_sys(outlook_mapi_sys: String) -> super::Result<String> {
        let pattern = Regex::new(r#"#\s*\[\s*link\s*\("#)?;
        let replacement = r#"#[delay_load("#;
        Ok(pattern
            .replace_all(&outlook_mapi_sys, replacement)
            .to_string())
    }

    fn format_mapi_sys(source_path: &Path) -> super::Result<()> {
        let mut cmd = ::std::process::Command::new("rustfmt");
        cmd.arg(source_path);
        cmd.output()?;
        Ok(())
    }

    fn read_mapi_sys(source_path: &Path) -> super::Result<String> {
        let mut source_file = fs::File::open(source_path)?;
        let mut updated = String::default();
        source_file.read_to_string(&mut updated)?;
        Ok(updated)
    }
}
