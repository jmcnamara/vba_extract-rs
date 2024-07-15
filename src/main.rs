// vba_extract - A simple utility to extract a vbaProject.bin binary from an
// Excel xlsm file for insertion into an `rust_xlsxwriter` file.
//
// If the macro is digitally signed the utility will also extract a
// vbaProjectSignature.bin file.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

use std::fs::File;
use std::io::copy;
use std::path::Path;

use clap::Parser;
use zip::ZipArchive;

// Clap struct to define the CLI flags and output.
#[derive(Parser, Debug)]
#[command(version, about)]
#[command(
    long_about = "Utility to extract a `vbaProject.bin` binary from an Excel xlsm macro file
for insertion into an `rust_xlsxwriter` file. If the macros are digitally
signed, it also extracts a `vbaProjectSignature.bin` file."
)]
struct Args {
    /// Input Excel xlsm filename.
    filename_xlsx: String,

    /// Output vba macro filename.
    #[arg(short, long, default_value = "vbaProject.bin")]
    output_macro_filename: String,

    /// Output vba signature filename (if present in the parent file).
    #[arg(short = 's', long, default_value = "vbaProjectSignature.bin")]
    output_sig_filename: String,
}

//
fn main() {
    // Parse the command line with Clap.
    let args = Args::parse();
    let xlsm_filename = args.filename_xlsx;

    // Open the Excel xlsm file.
    let xlsm_file = match File::open(&xlsm_filename) {
        Ok(file) => file,
        Err(err) => {
            eprintln!("Couldn't open file '{xlsm_filename}'. Error: {err}",);
            return;
        }
    };

    // The Excel xlsx/xlsm file format is in a zip container.
    let mut zip_archive = match ZipArchive::new(xlsm_file) {
        Ok(file) => file,
        Err(err) => {
            eprintln!("File '{xlsm_filename}' isn't a valid xlsm/zip file. Error: {err}");
            return;
        }
    };

    // Extract the vbaProject.bin binary file which contains the VBA macros.
    extract_bin_file(
        &mut zip_archive,
        "xl/vbaProject.bin",
        &xlsm_filename,
        &args.output_macro_filename,
    );

    // Extract the optional vbaProjectSignature.bin binary.
    extract_bin_file(
        &mut zip_archive,
        "xl/vbaProjectSignature.bin",
        &xlsm_filename,
        &args.output_sig_filename,
    );
}

// Extract a binary file from an Excel xlsm zip archive.
fn extract_bin_file(
    zip_archive: &mut ZipArchive<File>,
    binary_file: &str,
    xlsm_filename: &str,
    output_filename: &str,
) {
    let Ok(mut file) = zip_archive.by_name(binary_file) else {
        // Only raise a warning if vbaProject.bin is missing. The signature file
        // is optional.
        if binary_file == "xl/vbaProject.bin" {
            eprintln!("File '{xlsm_filename}' doesn't contain a '{binary_file}' file.");
        }
        return;
    };

    let output_path = Path::new(&output_filename);
    let mut output_file = match File::create(output_path) {
        Ok(file) => file,
        Err(err) => {
            eprintln!("Couldn't open '{output_filename}'. Error: {err}");
            return;
        }
    };

    // Copy the binary file to the OS.
    copy(&mut file, &mut output_file).unwrap();

    println!("Extracted {output_filename}.");
}
