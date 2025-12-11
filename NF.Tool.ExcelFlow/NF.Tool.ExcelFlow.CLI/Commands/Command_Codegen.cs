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

[Description("Generate code from exel")]
internal sealed class Command_Codegen : AsyncCommand<Command_Codegen.Settings>
{
    internal sealed class Settings : ExcelFlowConfigSettings
    {
        [Description("Input paths")]
        [CommandOption("-i|--input <PATH>")]
        public string[] InputPaths { get; set; } = [];

        [Description("Output directory")]
        [CommandOption("--output <DIRECTORY>")]
        public string OutputDirectory { get; set; } = string.Empty;

        [Description("Check code compilable")]
        [CommandOption("--check-compile")]
        public bool IsCheckCompilable { get; set; } = false;
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

        ConfigHelper.Override(config, settings);

        string[] excelPaths = Util.ExpandXlsxPaths(settings.InputPaths);
        (ModelRoot[] models, errOrNull) = await ExcelToModelBaker.Bake(excelPaths);
        if (errOrNull != null)
        {
            Console.Error.WriteLine(errOrNull);
            return (int)E_EXIT_CODE.FAIL_BAKE_MODEL;
        }

        if (string.IsNullOrEmpty(config.TemplatePath))
        {
            config.TemplatePath = Util.ExtractResourceToTempFilePath("Template.t4");
        }

        errOrNull = await ModelToCodeBaker.Bake(config, models, settings.IsCheckCompilable, settings.OutputDirectory);
        if (errOrNull != null)
        {
            Console.Error.WriteLine(errOrNull);
            return (int)E_EXIT_CODE.FAIL_BAKE_CODE;
        }
        return 0;
    }
}