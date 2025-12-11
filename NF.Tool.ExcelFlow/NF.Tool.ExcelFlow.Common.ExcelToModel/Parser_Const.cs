using NF.Tool.ExcelFlow.Common.Model;
using NF.Tool.ExcelFlow.Common.Model.Infos;
using NPOI.SS.UserModel;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace NF.Tool.ExcelFlow.Common.ExcelToModel;

internal static class Parser_Const
{
    public static (ConstInfo? infoOrNull, string? errOrNull) Parse(IWorkbook excel, RegionInfo region)
    {
        Debug.Assert(region.RegionType == RegionInfo.E_REGION_TYPE.CONST);

        ISheet sheet = excel.GetSheetAt(region.SheetIndex);
        string? errOrNull;
        (Dictionary<E_HEADER_CONST, PosCell>? dicHeader, errOrNull) = Util.CachingIndexY<E_HEADER_CONST>(sheet, region);
        if (errOrNull != null)
        {
            return (null, errOrNull);
        }
        (Dictionary<E_FIELD_CONST, PosCell> dicField, errOrNull) = Util.CachingIndexX<E_FIELD_CONST>(sheet, region, yHeaderField: dicHeader[E_HEADER_CONST.TABLE].Pos.y);
        if (errOrNull != null)
        {
            return (null, errOrNull);
        }

        (MetaCell constName, errOrNull) = GetConstName(sheet, region, dicHeader);
        if (errOrNull != null)
        {
            return (null, errOrNull);
        }

        List<ConstRow> rows = GetTableRows(sheet, region, dicField);

        return (new ConstInfo(constName, region, rows), null);
    }

    private static (MetaCell val, string? errOrNull) GetConstName(ISheet sheet, RegionInfo sheetRegion, Dictionary<E_HEADER_CONST, PosCell> dicHeader)
    {
        if (!dicHeader.TryGetValue(E_HEADER_CONST.CONST, out PosCell? mc2))
        {
            return (new MetaCell(new int2(-1, -1), sheetRegion.SheetName), null);
        }

        int x = sheetRegion.Region.TopLeft.x;
        int y = mc2!.Pos.y;
        int2 p = new int2(x, y);
        ICell? cellOrNull = sheet.CellOrNullXY(x, y);
        if (cellOrNull == null)
        {
            return (new MetaCell(new int2(-1, -1), sheetRegion.SheetName), Err.MessageWithInfo($"{p} - const not found"));
        }

        if (cellOrNull.CellType != CellType.String)
        {
            return (new MetaCell(new int2(-1, -1), sheetRegion.SheetName), Err.MessageWithInfo($"{p} - const name type error"));
        }

        string v = cellOrNull.StringCellValue;
        return (new MetaCell(p, v), null);
    }

    private static List<ConstRow> GetTableRows(ISheet sheet, RegionInfo sheetRegion, Dictionary<E_FIELD_CONST, PosCell> dicField)
    {
        int yMin = dicField.Values.First().Pos.y + 1;
        int yMax = sheetRegion.Region.BottomRight.y;
        Debug.Assert(yMax > yMin);

        List<ConstRow> ret = new List<ConstRow>(capacity: yMax - yMin);
        for (int y = yMin; y < yMax; ++y)
        {
            PosCell type = PosCell.Empty;
            PosCell name = PosCell.Empty;
            PosCell value = PosCell.Empty;
            PosCell desc = PosCell.Empty;
            PosCell attr = PosCell.Empty;

            foreach (E_FIELD_CONST e in ExEnum.GetValues<E_FIELD_CONST>())
            {
                if (!dicField.TryGetValue(e, out PosCell? vv))
                {
                    continue;
                }
                int x = vv.Pos.x;

                int2 pos = new int2(x, y);
                ICell? cellOrNull = sheet.CellOrNullXY(x, y);
                if (cellOrNull == null)
                {
                    continue;
                }

                if (cellOrNull.CellType != CellType.String)
                {
                    cellOrNull.SetCellType(CellType.String);
                }

                string v = cellOrNull.StringCellValue;
                PosCell mc = new PosCell(pos, v);

                switch (e)
                {
                    case E_FIELD_CONST.TYPE:
                        type = mc;
                        break;
                    case E_FIELD_CONST.NAME:
                        name = mc;
                        break;
                    case E_FIELD_CONST.VALUE:
                        if (type.Value == "string")
                        {
                            if (mc.Value.StartsWith('"') && mc.Value.EndsWith('"'))
                            {
                                value = mc;
                            }
                            else
                            {
                                value = new PosCell(mc.Pos, $"\"{mc.Value}\"");
                            }
                        }
                        else
                        {
                            value = mc;
                        }
                        break;
                    case E_FIELD_CONST.DESC:
                        desc = mc;
                        break;
                    case E_FIELD_CONST.ATTR:
                        attr = mc;
                        break;
                    default:
                        break;
                }
            }

            ConstRow row = new ConstRow(type, name, value, desc, attr);
            ret.Add(row);
        }
        return ret;
    }
}
