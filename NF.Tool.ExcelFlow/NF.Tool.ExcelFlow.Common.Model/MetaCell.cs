using System.Diagnostics.CodeAnalysis;
using System.Linq;

namespace NF.Tool.ExcelFlow.Common.Model;

public sealed class MetaCell : PosCell
{
    public string Name { get; }
    public string[] Metas { get; }

    public MetaCell(in int2 pos, [NotNull] string value) : base(pos, value)
    {
        string v = Value.Trim();
        string[] xs = v.Replace("\r\n", "\n").Split("\n").Select(x => x.Trim()).ToArray();
        Metas = xs.SkipLast(1).ToArray();
        Name = xs.Last();
    }

    public override string ToString()
    {
        return Name;
    }
}