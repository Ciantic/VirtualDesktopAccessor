extern crate cbindgen;

use cbindgen::Config;
use std::env;
use std::path::Path;

fn main() {
    let crate_env = env::var("CARGO_MANIFEST_DIR").unwrap();
    let crate_path = Path::new(&crate_env);

    let target_env = env::var("OUT_DIR").unwrap();
    let target_path = Path::new(&target_env);
    let header_path = target_path
        .join("..")
        .join("..")
        .join("..")
        .join("VirtualDesktopAccessor.h");

    let header_path = header_path.to_str().unwrap();

    let config = Config::from_root_or_default(crate_path);
    cbindgen::Builder::new()
        .with_crate(crate_path.to_str().unwrap())
        .with_config(config)
        .generate()
        .expect("Unable to generate bindings")
        .write_to_file(header_path);
}
