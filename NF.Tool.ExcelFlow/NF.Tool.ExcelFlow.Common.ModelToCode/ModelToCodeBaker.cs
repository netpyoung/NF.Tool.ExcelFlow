using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.Emit;
using Microsoft.VisualStudio.TextTemplating;
using Mono.TextTemplating;
using NF.Tool.ExcelFlow.Common.Model;
using System;
using System.CodeDom.Compiler;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.Loader;
using System.Text;
using System.Threading.Tasks;

namespace NF.Tool.ExcelFlow.Common.ModelToCode;

public static class ModelToCodeBaker
{
    public sealed record ResultGenerateCode(ModelRoot ModelRoot, string FileName, string Code, string ErrorString);
    public sealed record ParsingChunk(ResultGenerateCode ResultGenerateCode, SyntaxTree SyntaxTree, Diagnostic[] Diagnostics);

    public static async Task<string?> Bake(E_PART part, ExcelFlowConfig config, ModelRoot[] models, bool IsCheckCompilable, string outputDirectory)
    {
        Debug.Assert(!string.IsNullOrEmpty(config.TemplatePath));

        string? errOrNull;
        (ResultGenerateCode[] results, errOrNull) = await GenerateCodes(part, config, models);
        if (errOrNull != null)
        {
            return errOrNull;
        }

        if (IsCheckCompilable)
        {
            (Assembly? assembly, errOrNull) = await GetAssembly(results);
            if (errOrNull != null)
            {
                return errOrNull;
            }
        }

        if (!string.IsNullOrEmpty(outputDirectory))
        {
            string middleDir;
            switch (part)
            {
                case E_PART.CLIENT:
                    middleDir = "client";
                    break;
                case E_PART.SERVER:
                    middleDir = "server";
                    break;
                case E_PART.BOTH:
                default:
                    middleDir = string.Empty;
                    break;
            }

            (ResultGenerateCode r, string path)[] xs = results.Select(x => (r: x, path: Path.Combine(outputDirectory, middleDir, x.FileName))).ToArray();
            foreach ((ResultGenerateCode r, string path) x in xs)
            {
                string dir = Path.GetDirectoryName(x.path)!;
                Directory.CreateDirectory(dir);
            }

            await Parallel.ForEachAsync(xs, async (x, _) =>
            {
                ModelRoot m = x.r.ModelRoot;
                string content = x.r.Code;
                string path = x.path;
                await File.WriteAllTextAsync(path, content, _);
            });
        }
        return null;
    }

    public static async Task<(Assembly? val, string? errOrNull)> GetAssembly(E_PART part, ExcelFlowConfig config, ModelRoot[] models)
    {
        Debug.Assert(!string.IsNullOrEmpty(config.TemplatePath));

        string? errOrNull;
        (ResultGenerateCode[] results, errOrNull) = await GenerateCodes(part, config, models);
        if (errOrNull != null)
        {
            return (null, errOrNull);
        }

        (Assembly? assembly, errOrNull) = await GetAssembly(results);
        if (errOrNull != null)
        {
            return (null, errOrNull);
        }

        return (assembly, null);
    }

    private static async Task<(Assembly? val, string? errOrNull)> GetAssembly(ResultGenerateCode[] results)
    {
        ConcurrentBag<ParsingChunk> bag2 = [];

        await Parallel.ForEachAsync(results, async (result, ct) =>
        {
            SyntaxTree syntaxTree = await Task.Run(() => CSharpSyntaxTree.ParseText(result.Code), ct);
            bag2.Add(new ParsingChunk(result, syntaxTree, []));
        });

        ParsingChunk[] parseResults = bag2.ToArray();

        List<ParsingChunk> xs = parseResults
            .Select(x => x with { Diagnostics = x.SyntaxTree.GetDiagnostics().Where(x => x.Severity == DiagnosticSeverity.Error).ToArray() })
            .ToList();

        Diagnostic[] diagnostics = xs.SelectMany(x => x.Diagnostics).ToArray();
        if (diagnostics.Length > 0)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("Fail: GetAssembly - ParseText has error\n\n");
            foreach (ParsingChunk x in xs)
            {
                Diagnostic[] b = x.Diagnostics;
                if (b.Length != 0)
                {
                    sb.AppendLine($"{x.ResultGenerateCode.ModelRoot.ExcelFileName} {x.ResultGenerateCode.ModelRoot.SheetName} :");
                    foreach (Diagnostic y in b)
                    {
                        sb.Append("    ");
                        sb.AppendLine(y.ToString());
                    }
                }
            }
            string errString = sb.ToString();
            return (null, Err.MessageWithInfo(errString));
        }

        string tpa = (string)AppContext.GetData("TRUSTED_PLATFORM_ASSEMBLIES")!;
        string[] splitted = tpa.Split(Path.PathSeparator);
        PortableExecutableReference[] references = splitted.Select(x => MetadataReference.CreateFromFile(x)).ToArray();

        string assemblyName = $"assemblyName-{Guid.NewGuid()}";
        SyntaxTree[] syntaxTrees = parseResults.Select(x => x.SyntaxTree).ToArray();
        CSharpCompilationOptions options = new CSharpCompilationOptions(OutputKind.DynamicallyLinkedLibrary);
        CSharpCompilation compilation = CSharpCompilation.Create(assemblyName, syntaxTrees, references, options);

        using (MemoryStream peStream = new MemoryStream())
        using (MemoryStream pdbStream = new MemoryStream())
        {
            EmitResult result = compilation.Emit(peStream, pdbStream);
            if (!result.Success)
            {
                StringBuilder sb = new StringBuilder();
                sb.Append("Fail: GetAssembly - Compilation Emit failed\n\n");
                foreach (Diagnostic d in result.Diagnostics.Where(d => d.Severity == DiagnosticSeverity.Error))
                {
                    sb.Append("    ");
                    sb.AppendLine(d.ToString());
                }
                string errString = sb.ToString();
                return (null, Err.MessageWithInfo(errString));
            }

            peStream.Seek(0, SeekOrigin.Begin);
            pdbStream.Seek(0, SeekOrigin.Begin);
            Assembly assembly = AssemblyLoadContext.Default.LoadFromStream(peStream, pdbStream);
            return (assembly, null);
        }
    }

    private static async Task<(ResultGenerateCode[] val, string? errOrNull)> GenerateCodes(E_PART part, ExcelFlowConfig config, ModelRoot[] models)
    {
        string templatePath = Path.GetFullPath(config.TemplatePath);
        if (!File.Exists(templatePath))
        {
            return ([], Err.MessageWithInfo($"templatePath={templatePath} does not exists"));
        }

        string templateContent = File.ReadAllText(config.TemplatePath);

        ConcurrentBag<ResultGenerateCode> results = [];
        await Parallel.ForEachAsync(models, async (model, _) =>
        {
            ResultGenerateCode r = await GenerateCode(part, config, model, templatePath, templateContent);
            results.Add(r);
        });
        ResultGenerateCode[] ret = results.ToArray();

        ResultGenerateCode[] errs = ret.Where(x => !string.IsNullOrEmpty(x.ErrorString)).ToArray();
        if (errs.Length != 0)
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("Fail: Generate Codes\n");
            foreach (ResultGenerateCode x in errs)
            {
                sb.Append("    ");
                sb.AppendLine(x.ToString());
            }
            string errString = sb.ToString();
            return ([], Err.MessageWithInfo(errString));
        }
        return (ret, null);
    }

    private static async Task<ResultGenerateCode> GenerateCode(E_PART part, ExcelFlowConfig config, ModelRoot model, string templatePath, string templateContent)
    {
        TemplateGenerator host = new TemplateGenerator();
        host.UseInProcessCompiler();

        foreach (Type item in new Type[] { typeof(ModelRoot) })
        {
            string assemblyLocation = item.Assembly.Location;
            host.Refs.Add(assemblyLocation);
            host.Imports.Add(item.Namespace);
        }

        ITextTemplatingSession session = host.GetOrCreateSession();
        session["root"] = model;
        session["config"] = config;
        session["part"] = part;

        // ref: https://gist.github.com/rkttu/56c8a96b48a3534a55a671bb8774931f
        ParsedTemplate parsedTemplate = host.ParseTemplate(templatePath, templateContent);

        TemplateSettings templateSettings = TemplatingEngine.GetSettings(host, parsedTemplate);
        TemplatingEngine engine = new TemplatingEngine();
        (CompiledTemplate template, string[] references)? compilerResult = await engine.CompileTemplateAsync(
            parsedTemplate,
            null,
            host,
            templateSettings
        );

        CompilerErrorCollection errs = host.Errors;
        if (errs.Count != 0)
        {
            StringBuilder sb = new StringBuilder();
            foreach (CompilerError error in errs)
            {
                sb.AppendLine(error.ToString());
            }
            string errString = sb.ToString();
            return new ResultGenerateCode(model, string.Empty, string.Empty, errString);
        }


        bool compilerHasValue = compilerResult.HasValue;
        Debug.Assert(compilerHasValue);

        string renderedCode = compilerResult.Value!.template.Process();
        string fileNameWithoutExt = Path.GetFileNameWithoutExtension(model.ExcelFileName);
        string fileName = $"{fileNameWithoutExt}_{model.SheetName}.cs";
        return new ResultGenerateCode(model, fileName, renderedCode, string.Empty);
    }
}