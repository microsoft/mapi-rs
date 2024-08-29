extern crate windows_bindgen;

fn main() -> Result<()> {
    if mapi_bindgen::update_mapi_sys(mapi_winmd::generate_winmd()?)? {
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
    #[error("Failed to run dotnet CLI.\n{0}")]
    DotNetCli(String),
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

mod mapi_winmd {
    use std::{
        path::PathBuf,
        process::{Command, Output},
    };

    use cmake;
    use regex::RegexBuilder;

    use super::mapi_path::*;

    pub fn generate_winmd() -> super::Result<PathBuf> {
        let header_path = scrub_mapi_headers()?;
        install_clang_sharp()?;
        generate_winmd_from_scrubbed(header_path)?;
        
        let mut winmd_path = get_manifest_dir();
        winmd_path.push("winmd");
        winmd_path.push("bin");
        Ok(winmd_path)
    }

    const CMAKE_TRIPLET: &str = "x86_64-pc-windows-msvc";

    fn scrub_mapi_headers() -> super::Result<PathBuf> {
        let mut mapi_scrubbed = get_manifest_dir();
        mapi_scrubbed.push("winmd");
        mapi_scrubbed.push("mapi-scrubbed");

        Ok(cmake::Config::new(mapi_scrubbed)
            .profile("RelWithDebInfo")
            .target(CMAKE_TRIPLET)
            .host(CMAKE_TRIPLET)
            .generator("Ninja")
            .out_dir(get_out_dir())
            .build())
    }

    fn invoke_dotnet(args: &[&str]) -> super::Result<Output> {
        Command::new("dotnet")
            .args(args)
            .output()
            .map_err(|_| super::Error::DotNetCli(String::from("dotnet.exe not found")))
    }

    const CLANG_SHARP_NAME: &str = r"ClangSharpPInvokeGenerator";
    const CLANG_SHARP_VERSION: &str = r"17.0.1";

    fn install_clang_sharp() -> super::Result<()> {
        let output = invoke_dotnet(&["tool", "list", "-g"])?;
        let output = String::from_utf8_lossy(&output.stdout);

        let version_pattern = CLANG_SHARP_VERSION.replace('.', r"\.");
        let version_pattern =
            RegexBuilder::new(format!(r"{CLANG_SHARP_NAME}\s+{version_pattern}").as_str())
                .case_insensitive(true)
                .build()
                .expect("invalid regex");

        if !version_pattern.is_match(&output) {
            invoke_dotnet(&[
                "tool",
                "update",
                CLANG_SHARP_NAME,
                "--version",
                CLANG_SHARP_VERSION,
                "-g",
            ])?;
            println!("Installed {CLANG_SHARP_NAME} v{CLANG_SHARP_VERSION}");
        }

        Ok(())
    }

    fn generate_winmd_from_scrubbed(header_path: PathBuf) -> super::Result<()> {
        let mut winmd_path = get_manifest_dir();
        winmd_path.push("winmd");
        let winmd_path = winmd_path.display().to_string();
        let header_path = header_path.display().to_string();
        let mapi_scrubbed = format!(r"--property:MapiScrubbedDir={header_path}");

        let args = &["build", winmd_path.as_str(), mapi_scrubbed.as_str()];
        let output = invoke_dotnet(args)?;
        let output = String::from_utf8_lossy(&output.stdout);
        let args = args.join(" ");
        println!("dotnet {args}:\n{output}");

        Ok(())
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

    pub fn update_mapi_sys(winmd_path: PathBuf) -> super::Result<bool> {
        let source_path = generate_mapi_sys(winmd_path)?;
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

    fn generate_mapi_sys(mut winmd_path: PathBuf) -> super::Result<PathBuf> {
        const WINMD_FILE: &str = "Microsoft.Office.Outlook.MAPI.Win32.winmd";

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
