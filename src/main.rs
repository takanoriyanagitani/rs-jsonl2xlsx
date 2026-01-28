use std::process::ExitCode;

use clap::Parser;

use rs_jsonl2xlsx::BUF_CAP_DEFAULT;
use rs_jsonl2xlsx::SHEET_NAME_DEFAULT;
use rs_jsonl2xlsx::XErr;
use rs_jsonl2xlsx::stdin2jsons2sheet2stdout;

#[derive(Parser, Debug)]
#[command(version, about, long_about = None)]
struct Args {
    #[arg(long, default_value = SHEET_NAME_DEFAULT)]
    sheet_name: String,
}

fn sub() -> Result<(), XErr> {
    let args = Args::parse();
    let mut buf: Vec<u8> = Vec::with_capacity(BUF_CAP_DEFAULT);
    stdin2jsons2sheet2stdout(args.sheet_name, &mut buf)?;
    Ok(())
}

fn main() -> ExitCode {
    sub().map(|_| ExitCode::SUCCESS).unwrap_or_else(|e| {
        eprintln!("{e:#?}");
        ExitCode::FAILURE
    })
}
