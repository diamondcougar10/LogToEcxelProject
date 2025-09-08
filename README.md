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

If no input logs are provided the program prints a short usage message and
returns exit code 2. The default output name is
`Photomesh_RealityMesh_LogReport.xlsx`.

The resulting workbook contains sheets `PhotoMesh_Exports`,
`RealityMesh_Exports`, `Summary`, `HowTo` and `Data_Dictionary`.

## Testing

After building, run the tests with `ctest`.

