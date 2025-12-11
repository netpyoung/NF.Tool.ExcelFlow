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
using System.IO;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;

namespace NF.Tool.ExcelFlow.CLI.Commands;

[Description("Export sqlite database")]
internal sealed class Command_Sqlite : AsyncCommand<Command_Sqlite.Settings>
{
    internal sealed class Settings : ExcelFlowConfigSettings, IExcelFlowConfig_Sqlite
    {
        [Description("Output database path")]
        [CommandOption("--output <SQLITE_DB_PATH>")]
        public string Output { get; set; } = string.Empty;
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

        string[] excelPaths = Util.ExpandXlsxPaths(settings.InputPaths);
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
            if (!config.PartOrNull!.Value.HasFlag(part))
            {
                continue;
            }

            (Assembly? assembly, errOrNull) = await ModelToCodeBaker.GetAssembly(part, config, models);
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

            string dbPath = OutputDatabasePath(part, config);
            Directory.CreateDirectory(Path.GetDirectoryName(dbPath)!);
            await ExcelToSqliteBaker.Bake(updateList, dbPath);
        }
        return (int)E_EXIT_CODE.NONE;
    }

    string OutputDatabasePath(E_PART part, ExcelFlowConfig config)
    {
        if (config.PartOrNull != E_PART.BOTH)
        {
            return config.Db.Output;
        }

        string originFpath = Path.GetFullPath(config.Db.Output);
        if (string.IsNullOrEmpty(Path.GetExtension(originFpath)))
        {
            string dir = originFpath;
            string filename = $"{part.ToString().ToLower()}.db";
            string newPath = Path.Combine(dir, filename);
            return newPath;
        }
        else
        {
            string dir = Path.GetDirectoryName(originFpath)!;
            string filename = Path.GetFileName(originFpath);
            string newPath = Path.Combine(dir, part.ToString().ToLower(), filename);
            return newPath;
        }
    }
}