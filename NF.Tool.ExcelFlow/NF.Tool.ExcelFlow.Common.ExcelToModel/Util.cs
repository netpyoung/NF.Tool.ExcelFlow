using NF.Tool.ExcelFlow.Common.Model;
using NF.Tool.ExcelFlow.Common.Model.Infos;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;

namespace NF.Tool.ExcelFlow.Common.ExcelToModel;

internal static class Util
{
    public static (Dictionary<T, PosCell> xdic, string? errOrNull) CachingIndexX<T>(ISheet sheet, RegionInfo sheetRegion, int yHeaderField) where T : struct, Enum
    {
        Debug.Assert(typeof(T).IsStepRight());
        StringBuilder sbErr = new StringBuilder();

        int y = yHeaderField;
        T[] es = ExEnum.GetValues<T>();
        Dictionary<T, PosCell> dic = new Dictionary<T, PosCell>(capacity: es.Length);

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
            if (!Enum.TryParse(v, ignoreCase: true, out T e))
            {
                continue;
            }

            if (dic.TryGetValue(e, out PosCell? meta))
            {
                sbErr.AppendLine(Err.MessageWithInfo($"duplicate {e}"));
                continue;
            }

            dic[e] = new PosCell(pos, v);
        }

        foreach (T e in es.Where(x => x.IsRequired()))
        {
            if (!dic.TryGetValue(e, out PosCell? mcOrNull))
            {
                sbErr.AppendLine(Err.MessageWithInfo($"{e} is required. but not found in {sheetRegion}"));
                continue;
            }

            int pos = mcOrNull.Pos.x - sheetRegion.Region.TopLeft.x;
            int epos = e.GetPosition();
            if (pos != epos)
            {
                sbErr.AppendLine(Err.MessageWithInfo($"{e}'s pos = {mcOrNull.Pos.ToA1()}  | epos = {epos}"));
                continue;
            }
        }

        string err = sbErr.ToString();
        if (!string.IsNullOrEmpty(err))
        {
            return ([], err);
        }

        return (dic, null);
    }

    public static (Dictionary<T, PosCell> ydic, string? errOrNull) CachingIndexY<T>(ISheet sheet, RegionInfo sheetRegion) where T : struct, Enum
    {
        Debug.Assert(typeof(T).IsStepDown());
        StringBuilder sbErr = new StringBuilder();

        T[] es = ExEnum.GetValues<T>();
        Dictionary<T, PosCell> dic = new Dictionary<T, PosCell>(capacity: es.Length);

        int x = sheetRegion.Region.BottomRight.x;
        for (int y = sheetRegion.Region.TopLeft.y; y < sheetRegion.Region.BottomRight.y; ++y)
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
            if (!v.StartsWith('^'))
            {
                continue;
            }

            if (!Enum.TryParse(v.TrimStart('^'), ignoreCase: true, out T e))
            {
                sbErr.AppendLine(Err.MessageWithInfo($"Invalid value = {v} / {pos.ToA1()}"));
                continue;
            }

            if (dic.TryGetValue(e, out PosCell? meta))
            {
                sbErr.AppendLine(Err.MessageWithInfo($"duplicate {e}"));
                continue;
            }

            dic[e] = new PosCell(pos, v);

            if (e.IsTable())
            {
                break;
            }
        }

        foreach (T e in es.Where(x => x.IsRequired()))
        {
            if (!dic.TryGetValue(e, out PosCell? mcOrNull))
            {
                sbErr.AppendLine(Err.MessageWithInfo($"{e} is required. but not found in {sheetRegion}"));
                continue;
            }

            if (e.GetPosition() != -1)
            {
                int pos = mcOrNull.Pos.y - sheetRegion.Region.TopLeft.y;
                int epos = e.GetPosition() + 1;
                if (pos != epos)
                {
                    sbErr.AppendLine(Err.MessageWithInfo($"    {e}'s pos = {mcOrNull.Pos.ToA1()}  | epos = {epos}"));
                    continue;
                }
            }
        }

        T eContainer = es.First(x => x.IsTable());
        KeyValuePair<T, PosCell> container = dic.OrderByDescending(x => x.Value.Pos.y).First();
        if (!container.Key.Equals(eContainer))
        {
            sbErr.AppendLine(Err.MessageWithInfo($"    {eContainer}'s pos = {container.Value.Pos.ToA1()}"));
        }

        string err = sbErr.ToString();
        if (!string.IsNullOrEmpty(err))
        {
            return ([], err);
        }
        return (dic, null);
    }
}
