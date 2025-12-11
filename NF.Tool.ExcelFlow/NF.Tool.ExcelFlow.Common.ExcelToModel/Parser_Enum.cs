using NF.Tool.ExcelFlow.Common.Model;
using NF.Tool.ExcelFlow.Common.Model.Infos;
using NPOI.SS.UserModel;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace NF.Tool.ExcelFlow.Common.ExcelToModel;

internal static class Parser_Enum
{
    public static (EnumInfo? infoOrNull, string? errOrNull) TryParse(IWorkbook excel, RegionInfo region)
    {
        Debug.Assert(region.RegionType == RegionInfo.E_REGION_TYPE.ENUM);

        ISheet sheet = excel.GetSheetAt(region.SheetIndex);

        string? errOrNull;
        (Dictionary<E_HEADER_ENUM, PosCell>? dicHeader, errOrNull) = Util.CachingIndexY<E_HEADER_ENUM>(sheet, region);
        if (errOrNull != null)
        {
            return (null, errOrNull);
        }
        int yHeaderField = dicHeader[E_HEADER_ENUM.TABLE].Pos.y;

        (Dictionary<E_FIELD_ENUM, PosCell> dicField, errOrNull) = Util.CachingIndexX<E_FIELD_ENUM>(sheet, region, yHeaderField);
        if (errOrNull != null)
        {
            return (null, errOrNull);
        }

        List<EnumRow> rows = GetTableRows(sheet, region, dicField);
        (MetaCell enumName, errOrNull) = GetEnumName(sheet, region, dicHeader);
        if (errOrNull != null)
        {
            return (null, errOrNull);
        }
        return (new EnumInfo(enumName, region, rows), null);
    }

    private static (MetaCell val, string? errOrNull) GetEnumName(ISheet sheet, RegionInfo sheetRegion, Dictionary<E_HEADER_ENUM, PosCell> dicHeader)
    {
        if (!dicHeader.TryGetValue(E_HEADER_ENUM.ENUM, out PosCell? mc2))
        {
            return (new MetaCell(new int2(-1, -1), sheetRegion.SheetName), null);
        }

        int x = sheetRegion.Region.TopLeft.x;
        int y = mc2!.Pos.y;
        int2 p = new int2(x, y);
        ICell? cellOrNull = sheet.CellOrNullXY(x, y);
        if (cellOrNull == null)
        {
            return (new MetaCell(new int2(-1, -1), sheetRegion.SheetName), Err.MessageWithInfo($"{p} - enum not found"));
        }

        if (cellOrNull.CellType != CellType.String)
        {
            return (new MetaCell(new int2(-1, -1), sheetRegion.SheetName), Err.MessageWithInfo($"{p} - enum name type error"));
        }

        string v = cellOrNull.StringCellValue;
        return (new MetaCell(p, v), null);
    }

    private static List<EnumRow> GetTableRows(ISheet sheet, RegionInfo sheetRegion, Dictionary<E_FIELD_ENUM, PosCell> dicField)
    {
        int yMin = dicField.Values.First().Pos.y + 1;
        int yMax = sheetRegion.Region.BottomRight.y;
        Debug.Assert(yMax > yMin);

        List<EnumRow> ret = new List<EnumRow>(capacity: yMax - yMin);
        for (int y = yMin; y < yMax; ++y)
        {
            PosCell name = PosCell.Empty;
            PosCell value = PosCell.Empty;
            PosCell desc = PosCell.Empty;
            PosCell attr = PosCell.Empty;

            foreach (E_FIELD_ENUM e in ExEnum.GetValues<E_FIELD_ENUM>())
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
                    case E_FIELD_ENUM.NAME:
                        name = mc;
                        break;
                    case E_FIELD_ENUM.VALUE:
                        value = mc;
                        break;
                    case E_FIELD_ENUM.DESC:
                        desc = mc;
                        break;
                    case E_FIELD_ENUM.ATTR:
                        attr = mc;
                        break;
                    default:
                        break;
                }
            }

            EnumRow row = new EnumRow(name, value, desc, attr);
            ret.Add(row);
        }
        return ret;
    }
}
