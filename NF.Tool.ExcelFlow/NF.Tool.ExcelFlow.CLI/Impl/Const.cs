namespace NF.Tool.ExcelFlow.CLI.Impl;

internal static class Const
{
    public const string DEFAULT_CONFIG_FILENAME = "ExcelFlow.config.toml";
    public const string DEFAULT_TEMPLATE_T4_FILENAME = "Template.tt";
    public const string DEFAULT_NEWS_CONTENT = "Add your info here";
    public const string DEFAULT_EDIT_NEWS_CONTENT = @"

# Please write your news content. Lines starting with '#' will be ignored, and
# an empty message aborts.
";
    public const string DESCRIPTION_CONFIG = $"""
Pass a custom config file at <CONFIG_PATH>.
Default: {DEFAULT_CONFIG_FILENAME}
""";
}
