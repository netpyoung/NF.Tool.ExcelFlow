using NF.Tool.ExcelFlow.CLI.Impl;
using NF.Tool.ExcelFlow.Common.ExcelToDB;
using NF.Tool.ExcelFlow.Common.ExcelToModel;
using NF.Tool.ExcelFlow.Common.Model;
using NF.Tool.ExcelFlow.Common.Model.Intermediary;
using NF.Tool.ExcelFlow.Common.ModelToCode;
using Spectre.Console;
using Spectre.Console.Cli;
using System;
using System.ComponentModel;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;

namespace NF.Tool.ExcelFlow.CLI.Commands;

[Description("Export sqlite database")]
internal sealed class Command_Sqlite : AsyncCommand<Command_Sqlite.Settings>
{
    internal sealed class Settings : ExcelFlowConfigSettings
    {
        [Description("Input paths")]
        [CommandOption("-i|--input <PATH>")]
        public string[] InputPaths { get; set; } = [];

        [Description("Output database path")]
        [CommandOption("--output <PATH>")]
        public string OutputDatabasePath { get; set; } = string.Empty;
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

        (Assembly? assembly, errOrNull) = await ModelToCodeBaker.GetAssembly(config, models);
        if (errOrNull != null)
        {
            Console.Error.WriteLine(errOrNull);
            return (int)E_EXIT_CODE.FAIL_ASSEMBLY;
        }

        (ClassAndRows[] updateList, errOrNull) = await ExcelToModelBaker.CollectRows(models, assembly!);
        if (errOrNull != null)
        {
            Console.Error.WriteLine(errOrNull);
            return (int)E_EXIT_CODE.FAIL_COLLECT_ROWS;
        }

        await ExcelToSqliteBaker.Bake(updateList, settings.OutputDatabasePath);

        return (int)E_EXIT_CODE.NONE;
    }

}