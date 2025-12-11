# NF.Tool.ExelFlow

[![GitHub](https://img.shields.io/badge/GitHub-%23121011.svg?logo=github&logoColor=white)](https://github.com/netpyoung/NF.Tool.ExelFlow)
[![.NET Test Workflow](https://github.com/netpyoung/NF.Tool.ExelFlow/actions/workflows/dotnet-test.yml/badge.svg)](https://github.com/netpyoung/NF.Tool.ExelFlow/actions/workflows/dotnet-test.yml)
[![Document](https://img.shields.io/badge/document-docfx-blue)](https://netpyoung.github.io/NF.Tool.ExelFlow/)
[![License](https://img.shields.io/badge/license-MIT-C06524)](https://github.com/netpyoung/NF.Tool.ExelFlow/blob/main/LICENSE.md)
[![NuGet](https://img.shields.io/nuget/v/dotnet-excel-flow.svg?style=flat&label=NuGet%3A%20dotnet-excel-flow)](https://www.nuget.org/packages/dotnet-excel-flow/)

## Overview

Generate code(C#) & database(sqlite) from Exel.

``` mermaid
sequenceDiagram

    xlsx ->> excel-flow: dotnet excel-flow

    activate excel-flow
    excel-flow ->> excel-flow: T4 template
    excel-flow ->> csharp.cs : codegen
    excel-flow ->> excel-flow: Assembly - roslyn
    excel-flow ->> sqlite.db : db
    deactivate excel-flow
```

## Document

- [Documentation](https://netpyoung.github.io/NF.Tool.ExelFlow/)

## Dependencies

- [xoofx/Tomlyn library](https://github.com/xoofx/Tomlyn) for [Toml format](https://toml.io/en/) Config file.
- [mono/t4 library](https://github.com/mono/t4) for [T4 template](https://learn.microsoft.com/en-us/visualstudio/modeling/code-generation-and-t4-text-templates).
- [Spectre.Console.Cli](https://spectreconsole.net/cli/) for parse args.
- [SmartFormat](https://github.com/axuno/SmartFormat) for format string.
- [sqlite-net-pcl](https://github.com/praeclarum/sqlite-net) for sqlite
- [NPOI](https://github.com/nissl-lab/npoi) for exel
- [Microsoft.CodeAnalysis.CSharp](https://github.com/dotnet/roslyn) for Assembly

## License

This project is licensed under the MIT License. See the [LICENSE](https://github.com/netpyoung/NF.Tool.ExelFlow/blob/main/LICENSE.md) file for details.