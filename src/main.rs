pub mod consts;
pub mod parse_xlsx;

use crate::parse_xlsx::parse_xlsx;
use std::fs::File;
use std::io::prelude::*;

fn main() -> std::io::Result<()> {
    let s = parse_xlsx();
    let mut file = File::create("./markdown/FLandIOT.md")?;
    file.write_all(s.as_bytes())?;
    Ok(())
}
