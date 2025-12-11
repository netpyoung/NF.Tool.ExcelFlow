using NF.Tool.ExcelFlow.CLI.Impl;
using NF.Tool.ExcelFlow.Common.ExcelToModel;
using NF.Tool.ExcelFlow.Common.Model;
using NF.Tool.ExcelFlow.Common.ModelToCode;
using Spectre.Console;
using Spectre.Console.Cli;
using System;
using System.ComponentModel;
using System.Threading;
using System.Threading.Tasks;

namespace NF.Tool.ExcelFlow.CLI.Commands;

[Description("Generate code from excel")]
internal sealed class Command_Codegen : AsyncCommand<Command_Codegen.Settings>
{
    internal sealed class Settings : ExcelFlowConfigSettings, IExcelFlowConfig_Codegen
    {
        [Description("Output directory")]
        [CommandOption("--output <OUTPUT_DIRECTORY>")]
        public string Output { get; set; } = string.Empty;

        [Description("Check code compilable")]
        [CommandOption("--check-compile")]
        public bool? IsCheckCompilableOrNull { get; set; } = null;

        public bool IsCheckCompilable { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
    }

    public override ValidationResult Validate(CommandContext context, Settings settings)
    {
        return Util.CheckInputPaths(settings.InputPaths);
    }

    public override async Task<int> ExecuteAsync(CommandContext context, Settings settings, CancellationToken cancellationToken)
    {
        string? errOrNull = ConfigHelper.GetConfig(string.Empty, settings.Config, out string baseDirectory, out ExcelFlowConfig config);
        if (errOrNull != null)
        {
            Console.Error.WriteLine(errOrNull);
            return (int)E_EXIT_CODE.FAIL_GET_CONFIG;
        }

        errOrNull = ConfigHelper.Override(config, settings);
        if (errOrNull != null)
        {
            Console.Error.WriteLine(errOrNull);
            return (int)E_EXIT_CODE.FAIL_CONFIG_VALUE;
        }

        string[] excelPaths = Util.ExpandXlsxPaths(config.InputPaths);
        (ModelRoot[] models, errOrNull) = await ExcelToModelBaker.Bake(excelPaths);
        if (errOrNull != null)
        {
            Console.Error.WriteLine(errOrNull);
            return (int)E_EXIT_CODE.FAIL_BAKE_MODEL;
        }

        if (string.IsNullOrEmpty(config.TemplatePath))
        {
            config.TemplatePath = Util.ExtractResourceToTempFilePath(Const.DEFAULT_TEMPLATE_T4_FILENAME);
        }

        foreach (E_PART part in new E_PART[] { E_PART.CLIENT, E_PART.SERVER })
        {
            if (!config.Part.HasFlag(part))
            {
                continue;
            }

            errOrNull = await ModelToCodeBaker.Bake(part, config, models, config.Codegen.IsCheckCompilable, config.Codegen.Output);
            if (errOrNull != null)
            {
                Console.Error.WriteLine(errOrNull);
                return (int)E_EXIT_CODE.FAIL_BAKE_CODE;
            }
        }
        return 0;
    }
}