using System;

namespace NF.Tool.ExcelFlow.Common.Model;

public sealed class ExcelFlowException : Exception
{
    public ExcelFlowException(string message) : base(message)
    {
    }
}
