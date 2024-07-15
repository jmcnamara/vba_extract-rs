# vba_extract

The `vba_extract` command line application is a simple utility to extract a
`vbaProject.bin` binary from an Excel xlsm file for insertion into an
`rust_xlsxwriter` file.

If the macro is digitally signed the utility will also extract a
`vbaProjectSignature.bin` file.

## Usage

```bash
Usage: vba_extract [OPTIONS] <FILENAME_XLSX>

Arguments:
  <FILENAME_XLSX>
          Input Excel xlsm filename

Options:
  -o, --output-macro-filename <OUTPUT_MACRO_FILENAME>
          Output vba macro filename

          [default: vbaProject.bin]

  -s, --output-sig-filename <OUTPUT_SIG_FILENAME>
          Output vba signature filename (if present in the parent file)

          [default: vbaProjectSignature.bin]

  -h, --help
          Print help (see a summary with '-h')

  -V, --version
          Print version
```

## Installation

```bash
cargo install vba_extract
```
