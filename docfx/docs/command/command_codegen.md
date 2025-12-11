# Codegen

- Generate csharp code

``` sh
$ dotnet excel-flow codegen --help
DESCRIPTION:
Generate code from excel

USAGE:
    NF.Tool.ExcelFlow.CLI codegen [OPTIONS]

EXAMPLES:
    NF.Tool.ExcelFlow.CLI codegen --input hello.xlsx --output outdir
    NF.Tool.ExcelFlow.CLI codegen --input hello.xlsx --output outdir --check-compile

OPTIONS:
    -h, --help                         Prints help information
        --config <CONFIG_PATH>         Pass a custom config file at <CONFIG_PATH>.
                                       Default: ExcelFlow.config.toml
    -i, --input <XLSX_PATH>            Input paths
        --template <TEMPLATE_PATH>     Template path
        --namespace <NAMESPACE>        Namespace
        --part <PART>                  both | client | server
        --output <OUTPUT_DIRECTORY>    Output directory
        --check-compile                Check code compilable
```