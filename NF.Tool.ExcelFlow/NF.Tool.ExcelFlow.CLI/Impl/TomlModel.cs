using NF.Tool.ExcelFlow.Common.Model;

namespace NF.Tool.ExcelFlow.CLI.Impl;

public sealed class TomlModel
{
    public ExcelFlowConfig ExcelFlow { get; set; } = new ExcelFlowConfig();
}
