using System;

namespace NF.Tool.ExcelFlow.Common.Model.Attributes;

[AttributeUsage(AttributeTargets.Field, AllowMultiple = false)]
internal sealed class PositionAttribute : Attribute
{
    public int Index { get; }

    public PositionAttribute(int index)
    {
        Index = index;
    }
}