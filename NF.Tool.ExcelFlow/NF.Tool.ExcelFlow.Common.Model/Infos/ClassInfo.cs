using NF.Tool.ExcelFlow.Common.Model.Attributes;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Diagnostics;

namespace NF.Tool.ExcelFlow.Common.Model.Infos;

public sealed record ClassInfo(MetaCell ClassName, RegionInfo RegionInfo, List<HeaderColumnClass> HeaderColumns);

[DebuggerDisplay("{Name,nq}-{Type,nq}")]
public sealed record HeaderColumnClass(MetaCell Name, PosCell Type, PosCell Part, PosCell Attr, PosCell Desc)
{
    public override string ToString()
    {
        return Name.Value;
    }
}

[StepDown]
public enum E_HEADER_CLASS
{
    CLASS,
    [Required]
    TYPE,
    PART,
    DESC,
    ATTR,
    [Required]
    [Table]
    TABLE,
}

[Flags]
public enum E_PART
{
    Client = 1 << 0,
    Server = 1 << 1,
    Both = Client | Server
}