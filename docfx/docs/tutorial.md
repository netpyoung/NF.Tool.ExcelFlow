# Tutorial


0. Install

``` sh
$ dotnet tool install --global dotnet-excel-flow
```

``` sh
$ dotnet dotnet-excel-flow
USAGE:
    NF.Tool.ExcelFlow.CLI [OPTIONS] <COMMAND>

EXAMPLES:
    NF.Tool.ExcelFlow.CLI codegen --input hello.xlsx --output outdir
    NF.Tool.ExcelFlow.CLI codegen --input hello.xlsx --output outdir --check-compile
    NF.Tool.ExcelFlow.CLI sqlite --input hello.xlsx --output out.db

OPTIONS:
    -h, --help    Prints help information

COMMANDS:
    codegen    Generate code from exel
    sqlite     Export sqlite database
```
