using System.Diagnostics;

namespace NF.Tool.ExcelFlow.Common.Model;

[DebuggerDisplay("<int2({x}, {y})>")]
[System.Diagnostics.CodeAnalysis.SuppressMessage("Style", "IDE1006:Naming Styles", Justification = "<Pending>")]
public readonly record struct int2(int x, int y)
{
    public static int2 Zero => new int2(0, 0);
    public static int2 MinusOne => new int2(-1, -1);

    public override string ToString()
    {
        return $"({x}, {y})";
    }

    public string ToA1()
    {
        return ToA1(x, y);
    }

    public static string ToA1(int x, int y)
    {
        string col = "";
        for (x++; x > 0; x = (x - 1) / 26)
        {
            col = (char)('A' + (x - 1) % 26) + col;
        }
        return col + (y + 1);
    }
}