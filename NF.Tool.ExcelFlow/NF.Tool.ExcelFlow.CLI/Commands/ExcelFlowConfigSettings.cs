using NF.Tool.ExcelFlow.CLI.Impl;
using NF.Tool.ExcelFlow.Common.Model;
using Spectre.Console.Cli;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;

namespace NF.Tool.ExcelFlow.CLI.Commands;

internal class ExcelFlowConfigSettings : CommandSettings, IExcelFlowConfig
{
    [Description(Const.DESCRIPTION_CONFIG)]
    [CommandOption("--config <CONFIG_PATH>")]
    public string Config { get; set; } = string.Empty;

    [Description("Input paths")]
    [CommandOption("-i|--input <XLSX_PATH>")]
    public string[] InputPaths { get; set; } = [];

    [Description("Template path")]
    [CommandOption("--template <TEMPLATE_PATH>")]
    public string TemplatePath { get; set; } = string.Empty;

    [Description("Namespace")]
    [CommandOption("--namespace <NAMESPACE>")]
    public string Namespace { get; set; } = string.Empty;

    [Description("both | client | server")]
    [CommandOption("--part <PART>")]
    [TypeConverter(typeof(E_PART_Converter))]
    public E_PART? PartOrNull { get; set; } = null;
}

public sealed class E_PART_Converter : TypeConverter
{
    private readonly Dictionary<string, E_PART> _lookup;

    public E_PART_Converter()
    {
        _lookup = new Dictionary<string, E_PART>(StringComparer.OrdinalIgnoreCase) {
            { "both", E_PART.Both},
            { "client", E_PART.Client},
            { "server", E_PART.Server},
        };
    }

    public override object? ConvertFrom(ITypeDescriptorContext? context, CultureInfo? culture, object value)
    {
        if (value == null)
        {
            return null;
        }

        if (value is not string stringValue)
        {
            throw new ExcelFlowException("Can't convert value to E_PART.");
        }

        if (!_lookup.TryGetValue(stringValue, out E_PART ret))
        {
            throw new ExcelFlowException($"The value '{value}' is not a valid E_PART.");
        }
        return ret;
    }
}