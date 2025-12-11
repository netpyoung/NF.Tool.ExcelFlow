using NF.Tool.ExcelFlow.Common.Model.Attributes;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace NF.Tool.ExcelFlow.Common.Model.Infos;

public sealed record EnumInfo(MetaCell EnumName, RegionInfo RegionInfo, List<EnumRow> Rows);
public sealed record EnumRow(PosCell Name, PosCell Value, PosCell Desc, PosCell Attr);

[StepDown]
public enum E_HEADER_ENUM
{
    [Required]
    [Position(0)]
    ENUM,
    ATTR,
    [Required]
    [Table]
    TABLE,
}

[StepRight]
public enum E_FIELD_ENUM
{
    [Required]
    [Position(0)]
    NAME,

    [Required]
    [Position(1)]
    VALUE,
    DESC,
    ATTR,
}