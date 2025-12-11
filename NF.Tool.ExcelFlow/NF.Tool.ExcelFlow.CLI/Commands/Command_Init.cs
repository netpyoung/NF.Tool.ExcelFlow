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
        string templatePath = $"ExcelFlow.d/{templateFileName}";
        if (File.Exists(templatePath))
        {
            _ = errSb.AppendLine($"FileName [yellow]{templatePath}[/] already exists.");
        }

        string errStr = errSb.ToString();
        if (!string.IsNullOrEmpty(errStr))
        {
            AnsiConsole.Markup(errStr);
            return 1;
        }

        string configFileTempPath = Util.ExtractResourceToTempFilePath(Const.DEFAULT_CONFIG_FILENAME);
        File.Move(configFileTempPath, newConfigFilePath);
        _ = Directory.CreateDirectory("ExcelFlow.d");

        string templateFileTempPath = Util.ExtractResourceToTempFilePath(templateFileName);
        File.Move(templateFileTempPath, templatePath);

        {
            // display layout
            AnsiConsole.WriteLine("Initialized");
            Tree root = new Tree("./");
            _ = root.AddNode($"{Const.DEFAULT_CONFIG_FILENAME}");

            TreeNode changelogD = root.AddNode("[blue]ChangeLog.d/[/]");
            _ = changelogD.AddNode($"{templateFileName}");
            AnsiConsole.Write(root);
        }

        return 0;
    }
}