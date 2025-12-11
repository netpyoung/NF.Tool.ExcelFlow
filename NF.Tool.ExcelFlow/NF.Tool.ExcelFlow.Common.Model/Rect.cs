using System.Diagnostics;

namespace NF.Tool.ExcelFlow.Common.Model;

[DebuggerDisplay("{ToString(),nq}")]
public readonly record struct Rect(int2 TopLeft, int2 BottomRight)
{
    public Rect(int startX, int startY, int endX, int endY) : this(new int2(startX, startY), new int2(endX, endY))
    {
    }

    public override string ToString()
    {
        return $"[{TopLeft.ToA1()}, {BottomRight.ToA1()}]";
    }
}
