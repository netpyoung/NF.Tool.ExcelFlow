using System.Diagnostics.CodeAnalysis;

namespace NF.Tool.ExcelFlow.Common.Model;

public sealed class PartCell : PosCell
{
    public E_PART Part { get; }

    public PartCell(in int2 pos, [NotNull] string value, E_PART part) : base(pos, value)
    {
        Part = part;
    }

    public bool HasFlags(E_PART flags)
    {
        return Part.HasFlag(flags);
    }

    public override string ToString()
    {
        return Part.ToString();
    }
}
