using NF.Tool.ExcelFlow.Common.Model;
using NF.Tool.ExcelFlow.Common.Model.Infos;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;

namespace NF.Tool.ExcelFlow.Common.ExcelToModel;

internal static class Parser_Class
{
    public static (ClassInfo? infoOrNull, string? errOrNull) Parse(IWorkbook excel, RegionInfo sheetRegion)
    {
        Debug.Assert(sheetRegion.RegionType == RegionInfo.E_REGION_TYPE.CLASS);

        ISheet sheet = excel.GetSheetAt(sheetRegion.SheetIndex);

        string? errOrNull;
        (Dictionary<E_HEADER_CLASS, PosCell>? dicHeader, errOrNull) = Util.CachingIndexY<E_HEADER_CLASS>(sheet, sheetRegion);
        if (errOrNull != null)
        {
            return (null, errOrNull);
        }

        List<PosCell> columnNameMcList = GetColumnNames(sheet, sheetRegion, dicHeader);
        errOrNull = ValidateColumnNameMcList(columnNameMcList);
        if (errOrNull != null)
        {
            return (null, errOrNull);
        }

        (List<HeaderColumnClass> headerColumns, errOrNull) = GetHeaderColumns(sheet, dicHeader, columnNameMcList);
        if (errOrNull != null)
        {
            return (null, errOrNull);
        }

        (MetaCell className, errOrNull) = GetClassName(sheet, sheetRegion, dicHeader);
        if (errOrNull != null)
        {
            return (null, errOrNull);
        }

        ClassInfo dataClass = new ClassInfo(className, sheetRegion, headerColumns);
        return (dataClass, null);
    }

    private static (List<HeaderColumnClass> val, string? errOrNull) GetHeaderColumns(ISheet sheet, Dictionary<E_HEADER_CLASS, PosCell> dicHeader, List<PosCell> columnNameMcList)
    {
        List<HeaderColumnClass> headerColumns = new List<HeaderColumnClass>(capacity: columnNameMcList.Count);
        StringBuilder sbErr = new StringBuilder();
        foreach (PosCell pc in columnNameMcList)
        {
            MetaCell cName = new MetaCell(pc.Pos, pc.Value);
            PosCell cType = new PosCell(new int2(-1, -1), string.Empty);
            PartCell cPart = new PartCell(new int2(-1, -1), string.Empty, E_PART.BOTH);
            PosCell cAttr = new PosCell(new int2(-1, -1), string.Empty);
            PosCell cDesc = new PosCell(new int2(-1, -1), string.Empty);
            foreach ((E_HEADER_CLASS k, PosCell c) in dicHeader)
            {
                if (k == E_HEADER_CLASS.CLASS)
                {
                    continue;
                }
                if (k == E_HEADER_CLASS.TABLE)
                {
                    continue;
                }

                int x = pc.Pos.x;
                int y = c.Pos.y;
                int2 pos = new int2(x, y);
                ICell? cellOrNull = sheet.CellOrNullXY(x, y);
                if (cellOrNull == null)
                {
                    continue;
                }

                if (cellOrNull.CellType != CellType.String)
                {
                    continue;
                }

                string v = cellOrNull.StringCellValue;
                switch (k)
                {

                    case E_HEADER_CLASS.TYPE:
                        cType = new PosCell(pos, v);
                        break;
                    case E_HEADER_CLASS.PART:
                        if (!Enum.TryParse(v, ignoreCase: true, out E_PART part))
                        {
                            sbErr.AppendLine(Err.MessageWithInfo($"Fail: to parse part - {v} - {sheet.SheetName} - {pos.ToA1()}"));
                            break;
                        }

                        cPart = new PartCell(pos, v, part);
                        break;
                    case E_HEADER_CLASS.DESC:
                        cDesc = new PosCell(pos, v);
                        break;
                    case E_HEADER_CLASS.ATTR:
                        cAttr = new PosCell(pos, v);
                        break;
                    case E_HEADER_CLASS.CLASS:
                    case E_HEADER_CLASS.TABLE:
                    default:
                        break;
                }
            }

            HeaderColumnClass hc = new HeaderColumnClass(cName, cType, cPart, cAttr, cDesc);
            headerColumns.Add(hc);
        }
        string err = sbErr.ToString();
        if (!string.IsNullOrEmpty(err))
        {
            return ([], err);
        }

        return (headerColumns, null);
    }

    private static (MetaCell val, string? errOrNull) GetClassName(ISheet sheet, RegionInfo sheetRegion, Dictionary<E_HEADER_CLASS, PosCell> dicHeader)
    {
        if (!dicHeader.TryGetValue(E_HEADER_CLASS.CLASS, out PosCell? mc2))
        {
            return (new MetaCell(new int2(-1, -1), sheetRegion.SheetName), null);
        }

        int x = sheetRegion.Region.TopLeft.x;
        int y = mc2!.Pos.y;
        int2 pos = new int2(x, y);
        ICell? cellOrNull = sheet.CellOrNullXY(x, y);
        if (cellOrNull == null)
        {
            return (new MetaCell(new int2(-1, -1), sheetRegion.SheetName), Err.MessageWithInfo("class not found"));
        }

        if (cellOrNull.CellType != CellType.String)
        {
            return (new MetaCell(new int2(-1, -1), sheetRegion.SheetName), Err.MessageWithInfo("class name type error"));
        }

        string v = cellOrNull.StringCellValue;
        return (new MetaCell(pos, v), null);
    }

    private static List<PosCell> GetColumnNames(ISheet sheet, RegionInfo sheetRegion, Dictionary<E_HEADER_CLASS, PosCell> dicHeader)
    {
        List<PosCell> ret = new List<PosCell>(sheetRegion.Region.BottomRight.x - sheetRegion.Region.TopLeft.x);

        PosCell dv = dicHeader[E_HEADER_CLASS.TABLE];
        int y = dv.Pos.y;
        for (int x = sheetRegion.Region.TopLeft.x; x < sheetRegion.Region.BottomRight.x; ++x)
        {
            int2 pos = new int2(x, y);
            ICell? cellOrNull = sheet.CellOrNullXY(x, y);
            if (cellOrNull == null)
            {
                continue;
            }

            if (cellOrNull.CellType != CellType.String)
            {
                continue;
            }

            string v = cellOrNull.StringCellValue;
            if (v.StartsWith('_'))
            {
                continue;
            }

            ret.Add(new PosCell(pos, v));
        }

        return ret;
    }

    private static string? ValidateColumnNameMcList(List<PosCell> columnNameMcList)
    {
        List<IGrouping<string, PosCell>> xxxxxxx = columnNameMcList.GroupBy(x => x.Value).Where(x => x.Count() != 1).ToList();
        if (xxxxxxx.Count == 0)
        {
            return null;
        }

        StringBuilder sb = new StringBuilder();
        sb.Append("Duplicate: ");
        foreach (IGrouping<string, PosCell> group in xxxxxxx)
        {
            List<PosCell> lst = group.ToList();
            sb.Append(group.Key);
            foreach (PosCell mc in group)
            {
                sb.Append(mc.Pos.ToA1());
            }
            sb.AppendLine();
        }

        return sb.ToString();
    }
}
