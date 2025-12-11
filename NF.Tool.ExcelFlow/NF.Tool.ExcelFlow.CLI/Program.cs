using NF.Tool.ExcelFlow.CLI.Commands;
using NF.Tool.ExcelFlow.Common.Model;
using Spectre.Console;
using Spectre.Console.Cli;
using System;
using System.Threading.Tasks;

namespace NF.Tool.ExcelFlow.CLI;

internal sealed class Program
{
    internal static async Task<int> Main(string[] args)
    {
        CommandApp app = new CommandApp();

        app.Configure(config =>
        {
            _ = config.PropagateExceptions();
            _ = config.UseStrictParsing();
            _ = config.SetApplicationName("NF.Tool.ExcelFlow.CLI");

            _ = config.AddCommand<Command_Init>("init")
                  .WithExample("init");
            _ = config.AddCommand<Command_Codegen>("codegen")
                .WithExample("codegen", "--input", "hello.xlsx", "--output", "outdir")
                .WithExample("codegen", "--input", "hello.xlsx", "--output", "outdir", "--check-compile");
            _ = config.AddCommand<Command_Sqlite>("sqlite")
                .WithExample("sqlite", "--input", "hello.xlsx", "--output", "out.db");
#if DEBUG
            _ = config.PropagateExceptions();
            _ = config.ValidateExamples();
#endif // DEBUG
        });

        try
        {
            return await app.RunAsync(args);
        }
        catch (CommandAppException ex)
        {
            if (ex.Pretty is { } pretty)
            {
                AnsiConsole.Write(pretty);
            }
            else
            {
                AnsiConsole.MarkupInterpolated($"[red]Error:[/] {ex.Message}");
            }
            return (int)E_EXIT_CODE.COMMAND_APP_EXCEPTION;
        }
        catch (ExcelFlowException ex)
        {
            AnsiConsole.WriteException(ex, new ExceptionSettings()
            {
                Format = ExceptionFormats.ShortenEverything,
                Style = new()
                {
                    ParameterName = Color.Grey,
                    ParameterType = Color.Grey78,
                    LineNumber = Color.Grey78,
                },
            });
            return (int)E_EXIT_CODE.EXCEL_FLOW_EXCEPTION;
        }
        catch (Exception ex)
        {
            AnsiConsole.WriteException(ex, new ExceptionSettings()
            {
                Format = ExceptionFormats.ShortenEverything,
                Style = new()
                {
                    ParameterName = Color.Grey,
                    ParameterType = Color.Grey78,
                    LineNumber = Color.Grey78,
                },
            });
            return (int)E_EXIT_CODE.UNHANDLE;
        }
    }
}