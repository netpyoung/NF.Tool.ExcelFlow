using System;

namespace NF.Tool.ExcelFlow.Common.Model;

[Flags]
public enum E_PART
{
    CLIENT = 1 << 0,
    SERVER = 1 << 1,
    BOTH = CLIENT | SERVER
}