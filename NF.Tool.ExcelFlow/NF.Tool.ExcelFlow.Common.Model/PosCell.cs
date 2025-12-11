using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;

namespace NF.Tool.ExcelFlow.Common.Model;

[DebuggerDisplay("{Pos,nq}-{Value,nq}")]
public class PosCell
{
    public int2 Pos { get; }
    public string Value { get; }

    public static PosCell Empty => new PosCell(int2.MinusOne, string.Empty);

    public PosCell(in int2 pos, [NotNull] string value)
    {
        Pos = pos;
        Value = value.Trim();
    }

    public override string ToString()
    {
        return Value;
    }
}