using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Text.RegularExpressions;

namespace NF.Tool.ExcelFlow.Common.Model.Infos;

[DebuggerDisplay("{ToString(),nq}")]
public sealed record RegionInfo(int SheetIndex, string SheetName, Rect Region, RegionInfo.E_REGION_TYPE RegionType, string RawTitle)
{
    public static RegionInfo New(int sheetIndex, string sheetName, [NotNull] string rawTitle, in Rect region, E_REGION_TYPE regionType)
    {
        string trimedTitle = Regex.Replace(rawTitle.Trim(), @"\|[^|]*\|", "").Trim();
        RegionInfo x = new RegionInfo(sheetIndex, sheetName, region, regionType, rawTitle);
        return x;
    }

    public enum E_REGION_TYPE
    {
        NONE,
        CONST,
        ENUM,
        CLASS,
    }

    public override string ToString()
    {
        return $"SheetName: {SheetName} / SheetIndex:{SheetIndex} / Region: {Region}";
    }
}