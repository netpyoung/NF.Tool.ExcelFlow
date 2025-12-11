using NF.Tool.ExcelFlow.CLI.Impl;
using Spectre.Console.Cli;
using System.ComponentModel;

namespace NF.Tool.ExcelFlow.CLI.Commands;

internal class ExcelFlowConfigSettings : CommandSettings
{
    [Description(Const.DESCRIPTION_CONFIG)]
    [CommandOption("--config")]
    public string Config { get; set; } = string.Empty;

    [Description("Template path")]
    [CommandOption("--template <PATH>")]
    public string TemplatePath { get; set; } = string.Empty;

    [Description("Namespace")]
    [CommandOption("--namespace <NAMESPACE>")]
    public string Namespace { get; set; } = string.Empty;

    // TODO(pyoung) - part - client / server / both
}
