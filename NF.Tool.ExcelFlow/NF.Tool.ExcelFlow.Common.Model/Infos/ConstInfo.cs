using NF.Tool.ExcelFlow.Common.Model.Attributes;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Diagnostics;

namespace NF.Tool.ExcelFlow.Common.Model.Infos;

[DebuggerDisplay("{RegionInfo}")]
public sealed record ConstInfo(MetaCell ConstName, RegionInfo RegionInfo, List<ConstRow> Rows);
public sealed record ConstRow(PosCell Type, PosCell Name, PosCell Value, PosCell Desc, PosCell Attr);

[StepDown]
public enum E_HEADER_CONST
{
    [Required]
    [Position(0)]
    CONST,

    [Required]
    [Table]
    TABLE,
}

[StepRight]
public enum E_FIELD_CONST
{
    [Required]
    [Position(0)]
    TYPE = 0,
    [Required]
    [Position(1)]
    NAME,
    [Required]
    [Position(2)]
    VALUE,
    DESC,
    ATTR,
}