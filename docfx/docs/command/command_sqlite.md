# Sqlite

- Generate sqlite database

``` sh
$ dotnet excel-flow codegen --help
DESCRIPTION:
Export sqlite database

USAGE:
    NF.Tool.ExcelFlow.CLI sqlite [OPTIONS]

EXAMPLES:
    NF.Tool.ExcelFlow.CLI sqlite --input hello.xlsx --output out.db

OPTIONS:
    -h, --help                        Prints help information
        --config <CONFIG_PATH>        Pass a custom config file at <CONFIG_PATH>.
                                      Default: ExcelFlow.config.toml
    -i, --input <XLSX_PATH>           Input paths
        --template <TEMPLATE_PATH>    Template path
        --namespace <NAMESPACE>       Namespace
        --part <PART>                 both | client | server
        --output <SQLITE_DB_PATH>     Output database path
```