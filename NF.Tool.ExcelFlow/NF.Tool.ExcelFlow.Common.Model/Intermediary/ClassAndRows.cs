using System;
using System.Collections.Generic;

namespace NF.Tool.ExcelFlow.Common.Model.Intermediary;

public sealed record ClassAndRows(Type Class, List<object> Rows);