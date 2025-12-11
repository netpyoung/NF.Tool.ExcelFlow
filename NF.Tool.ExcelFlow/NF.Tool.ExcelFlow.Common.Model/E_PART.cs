using System;

namespace NF.Tool.ExcelFlow.Common.Model;

[Flags]
public enum E_PART
{
    Client = 1 << 0,
    Server = 1 << 1,
    Both = Client | Server
}