using System;

namespace NF.Tool.ExcelFlow.Common.Model.Attributes;

[AttributeUsage(AttributeTargets.Field, AllowMultiple = false)]
internal sealed class TableAttribute : Attribute
{
}