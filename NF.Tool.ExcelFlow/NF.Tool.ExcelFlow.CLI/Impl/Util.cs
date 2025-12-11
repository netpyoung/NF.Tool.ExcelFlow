using NF.Tool.ExcelFlow.Common.Model;
using Spectre.Console;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace NF.Tool.ExcelFlow.CLI.Impl;

internal static class Util
{
    public static string ExtractResourceToTempFilePath(string resourceName)
    {
        Assembly assembly = Assembly.GetExecutingAssembly();
        string tempFilePath = $"{Path.GetTempPath()}/{resourceName}";
        using (Stream resourceStream = assembly.GetManifestResourceStream(resourceName)!)
        {
            Debug.Assert(resourceStream != null);
            using (FileStream fileStream = new FileStream(tempFilePath, FileMode.Create, FileAccess.Write))
            {
                resourceStream.CopyTo(fileStream);
            }
        }
        return tempFilePath;
    }

    public static ValidationResult CheckInputPaths(string[] inputPaths)
    {
        if (inputPaths.Length == 0)
        {
            return ValidationResult.Error("Fail: Need inputPaths");
        }

        IGrouping<string, string>[] groups = inputPaths.GroupBy(x => x).Where(x => x.Count() > 1).ToArray();
        if (groups.Length != 0)
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("Fail: Duplicate Path");
            sb.AppendLine();
            foreach (IGrouping<string, string> group in groups)
            {
                string k = group.Key;
                sb.AppendLine($"Path: {k}");
                foreach (string v in group)
                {
                    sb.AppendLine($"    {v}");
                }
                sb.AppendLine();
            }
            return ValidationResult.Error(sb.ToString());
        }

        foreach (string path in inputPaths)
        {
            if (Directory.Exists(path))
            {
                continue;
            }

            if (File.Exists(path))
            {
                if (path.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }
            }

            return ValidationResult.Error($"Invalid path: '{path}' (not file nor directory)");
        }

        return ValidationResult.Success();
    }

    public static string[] ExpandXlsxPaths(IEnumerable<string> inputPaths)
    {
        ConcurrentBag<string> bag = [];
        Parallel.ForEach(inputPaths, path =>
        {
            if (File.Exists(path))
            {
                if (Path.GetExtension(path).Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
                {
                    if (!Path.GetFileName(path).StartsWith('~'))
                    {
                        bag.Add(Path.GetFullPath(path));
                    }
                }

                return;
            }

            if (Directory.Exists(path))
            {
                string[] files = Directory.GetFiles(path, "*.xlsx", SearchOption.AllDirectories);
                foreach (string f in files)
                {
                    if (!Path.GetFileName(f).StartsWith('~'))
                    {
                        bag.Add(Path.GetFullPath(f));
                    }
                }
                return;
            }
        });

        string[] excelPaths = bag.ToArray().Distinct().ToArray();
        CheckExcelPaths(inputPaths, excelPaths);

        return excelPaths;
    }

    private static void CheckExcelPaths(IEnumerable<string> inputPaths, string[] excelPaths)
    {
        if (excelPaths.Length == 0)
        {
            throw new ExcelFlowException($"Fail: No xlsx files found in {string.Join(", ", inputPaths)}");
        }

        IGrouping<string, string>[] groups = excelPaths.GroupBy(k => Path.GetFileName(k), v => v).Where(x => x.Count() > 1).ToArray();
        if (groups.Length != 0)
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("Fail: Duplicate FileName");
            sb.AppendLine();
            foreach (IGrouping<string, string> group in groups)
            {
                string k = group.Key;
                sb.AppendLine($"FileName: {k}");
                foreach (string v in group)
                {
                    sb.AppendLine($"    {v}");
                }
                sb.AppendLine();
            }
            throw new ExcelFlowException(sb.ToString());
        }
    }
}
