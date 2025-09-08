# logtoExcel

`logtoExcel` is a C++20 command line tool that parses PhotoMesh and Reality Mesh
logs and writes an Excel workbook summarizing the exports.

## Building

This project uses CMake and vcpkg for dependencies.

```bash
cmake -S . -B build
cmake --build build
```

## Usage

```bash
logtoExcel --photomesh pm1.log pm2.log --realitymesh rm1.log -o Report.xlsx
```

By default each run also appends rows into `EXCEL OUTPUTS/All_Exports.tsv`
(skipping duplicates by log path) and rebuilds the single-sheet
`EXCEL OUTPUTS/All_Exports.xlsx`. The per-run multi-sheet workbook is still
generated unless `--no-report` is used. Use `--single-only` to update only the
master workbook, or `--outputs-dir <folder>` to choose a custom output folder.

If no input logs are provided the program prints a short usage message and
returns exit code 2. The default output name is
`Photomesh_RealityMesh_LogReport.xlsx`.

The per-run workbook contains sheets `PhotoMesh_Exports`,
`RealityMesh_Exports`, `Summary`, `HowTo` and `Data_Dictionary`.

## Testing

After building, run the tests with `ctest`.

