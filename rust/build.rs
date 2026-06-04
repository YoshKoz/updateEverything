use std::{env, path::PathBuf};

fn main() {
    if env::var("CARGO_CFG_TARGET_OS").as_deref() != Ok("windows") {
        return;
    }

    let manifest = PathBuf::from(env::var_os("CARGO_MANIFEST_DIR").unwrap()).join("app.manifest");
    println!("cargo:rustc-link-arg-bin=updateeverything=/MANIFEST:EMBED");
    println!(
        "cargo:rustc-link-arg-bin=updateeverything=/MANIFESTINPUT:{}",
        manifest.display()
    );
}
