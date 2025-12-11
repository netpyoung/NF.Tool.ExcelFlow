using NF.Tool.ExcelFlow.CLI.Commands;
using NF.Tool.ExcelFlow.Common.Model;
using System.IO;
using System.Text;
using Tomlyn;
using Tomlyn.Syntax;

namespace NF.Tool.ExcelFlow.CLI.Impl;

internal static class ConfigHelper
{
    public static string? GetConfig(string directory, string configPath, out string outBaseDirectory, out ExcelFlowConfig outConfig)
    {
        if (string.IsNullOrEmpty(configPath))
        {
            return TraverseToParentForConfig(directory, out outBaseDirectory, out outConfig);
        }

        string configFpath = Path.GetFullPath(configPath);
        if (!string.IsNullOrEmpty(directory))
        {
            outBaseDirectory = Path.GetFullPath(directory);
        }
        else
        {
            outBaseDirectory = Path.GetDirectoryName(configFpath)!;
        }

        if (!File.Exists(configFpath))
        {
            outConfig = ExcelFlowConfig.Default();
            return null;
        }

        (outConfig, string? exOrNull) = LoadConfigFromFile(configFpath);
        return exOrNull;
    }

    private static string? TraverseToParentForConfig(string path, out string outDirectoryFpath, out ExcelFlowConfig outConfig)
    {
        string startDirectoryFpath;
        if (!string.IsNullOrEmpty(path))
        {
            startDirectoryFpath = Path.GetFullPath(path);
        }
        else
        {
            startDirectoryFpath = Directory.GetCurrentDirectory();
        }

        outDirectoryFpath = startDirectoryFpath;
        while (true)
        {
            string configFpath = Path.Combine(outDirectoryFpath, Const.DEFAULT_CONFIG_FILENAME);
            if (File.Exists(configFpath))
            {
                (outConfig, string? exOrNull) = LoadConfigFromFile(configFpath);
                return exOrNull;
            }

            DirectoryInfo? parentOrNull = Directory.GetParent(outDirectoryFpath);
            if (parentOrNull == null)
            {
                outConfig = ExcelFlowConfig.Default();
                return null;
            }
            outDirectoryFpath = parentOrNull.FullName;
        }
    }

    private static (ExcelFlowConfig config, string? exOrNull) LoadConfigFromFile(string configFpath)
    {
        string configText = File.ReadAllText(configFpath);
        TomlModelOptions option = new TomlModelOptions
        {
            ConvertFieldName = StringIdentity,
            ConvertPropertyName = StringIdentity
        };

        bool isSuccess = Toml.TryToModel(configText, out TomlModel? modelOrNull, out DiagnosticsBag? diagostics, options: option);
        if (!isSuccess)
        {
            StringBuilder sb = new StringBuilder();
            _ = sb.AppendLine($"TOML Parsing Error: configFpath={configFpath}");
            foreach (DiagnosticMessage x in diagostics!)
            {
                _ = sb.AppendLine(x.ToString());
            }
            return (ExcelFlowConfig.Default(), sb.ToString());
        }

        TomlModel model = modelOrNull!;
        ExcelFlowConfig config = model.ExcelFlow;
        return (config, null);
    }

    private static string StringIdentity(string x)
    {
        return x;
    }

    internal static string? Override(ExcelFlowConfig config, ExcelFlowConfigSettings settings)
    {
        StringBuilder sb = new StringBuilder();
        if (settings.InputPaths.Length != 0)
        {
            config.InputPaths = settings.InputPaths;
        }
        if (config.InputPaths.Length == 0)
        {
            sb.AppendLine(Err.MessageWithInfo("Fail: Need inputPaths"));
        }

        if (!string.IsNullOrEmpty(settings.TemplatePath))
        {
            config.TemplatePath = settings.TemplatePath;
        }

        if (!string.IsNullOrEmpty(settings.Namespace))
        {
            config.Namespace = settings.Namespace;
        }
        if (string.IsNullOrEmpty(config.Namespace))
        {
            sb.AppendLine(Err.MessageWithInfo("Fail: Need Namespace"));
        }

        if (settings.PartOrNull != null)
        {
            config.PartOrNull = settings.PartOrNull;
        }

        if (settings is Command_Codegen.Settings settings_codegen)
        {
            config.Codegen.Output = settings_codegen.Output;
            config.Codegen.IsCheckCompilable = settings_codegen.IsCheckCompilable;
        }
        else if (settings is Command_Sqlite.Settings settings_db)
        {
            config.Db.Output = settings_db.Output;
        }

        string err = sb.ToString();
        if (!string.IsNullOrEmpty(err))
        {
            return err;
        }
        return null;
    }
}