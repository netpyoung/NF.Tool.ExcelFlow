using NF.Tool.ExcelFlow.CLI.Impl;
using Spectre.Console;
using Spectre.Console.Cli;
using System.ComponentModel;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace NF.Tool.ExcelFlow.CLI.Commands;

[Description("Init excel-flow setup")]
internal sealed class Command_Init : AsyncCommand<Command_Init.Settings>
{
    internal sealed class Settings : CommandSettings
    {
        [Description("Specify config file name.")]
        [CommandOption("--file")]
        [DefaultValue(Const.DEFAULT_CONFIG_FILENAME)]
        public string FileName { get; set; } = string.Empty;
    }

    public override async Task<int> ExecuteAsync(CommandContext context, Settings settings, CancellationToken cancellationToken)
    {
        StringBuilder errSb = new StringBuilder();
        string newConfigFilePath = settings.FileName;

        if (File.Exists(newConfigFilePath))
        {
            _ = errSb.AppendLine($"FileName [yellow]{newConfigFilePath}[/] already exists.");
        }

        string templateFileName = Const.DEFAULT_TEMPLATE_T4_FILENAME;
        string templatePath = $"{Const.DEFUALT_EXCEL_FLOW_DIR}/{templateFileName}";
        if (File.Exists(templatePath))
        {
            _ = errSb.AppendLine($"FileName [yellow]{templatePath}[/] already exists.");
        }

        string sampleFileName = Const.DEFAULT_SAMPLE_XLSX;
        string samplePath = $"{Const.DEFUALT_EXCEL_FLOW_DIR}/{sampleFileName}";
        if (File.Exists(samplePath))
        {
            _ = errSb.AppendLine($"FileName [yellow]{samplePath}[/] already exists.");
        }

        string errStr = errSb.ToString();
        if (!string.IsNullOrEmpty(errStr))
        {
            AnsiConsole.Markup(errStr);
            return 1;
        }

        string configFileTempPath = Util.ExtractResourceToTempFilePath(Const.DEFAULT_CONFIG_FILENAME);
        File.Move(configFileTempPath, newConfigFilePath);
        _ = Directory.CreateDirectory(Const.DEFUALT_EXCEL_FLOW_DIR);

        string templateFileTempPath = Util.ExtractResourceToTempFilePath(templateFileName);
        File.Move(templateFileTempPath, templatePath);
        string sampleFileTempPath = Util.ExtractResourceToTempFilePath(sampleFileName);
        File.Move(sampleFileTempPath, samplePath);

        {
            // display layout
            AnsiConsole.WriteLine("Initialized");
            Tree root = new Tree("./");
            _ = root.AddNode(Const.DEFAULT_CONFIG_FILENAME);

            TreeNode excelFlowD = root.AddNode($"[blue]{Const.DEFUALT_EXCEL_FLOW_DIR}/ [/]");
            _ = excelFlowD.AddNode(templateFileName);
            _ = excelFlowD.AddNode(sampleFileName);
            AnsiConsole.Write(root);
        }

        return 0;
    }
}