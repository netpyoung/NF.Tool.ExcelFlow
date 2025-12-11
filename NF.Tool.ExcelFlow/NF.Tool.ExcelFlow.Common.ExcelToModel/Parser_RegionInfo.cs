using NF.Tool.ExcelFlow.Common.Model;
using NF.Tool.ExcelFlow.Common.Model.Infos;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;

namespace NF.Tool.ExcelFlow.Common.ExcelToModel;

internal static class Parser_RegionInfo
{
    public static List<List<RegionInfo>> CollectRegionsFromExcel(XSSFWorkbook excel)
    {
        List<List<RegionInfo>> ret = new List<List<RegionInfo>>(capacity: excel.NumberOfSheets);
        for (int sheetIndex = 0; sheetIndex < excel.NumberOfSheets; ++sheetIndex)
        {
            List<RegionInfo> lst = [];
            if (TryCollectMergedFlowSheetRegionInSheet(excel, sheetIndex, out List<RegionInfo> xs1))
            {
                lst.AddRange(xs1);
            }

            if (TryCollectDefaultClassFlowSheetRegionInSheet(excel, sheetIndex, out RegionInfo? xs2))
            {
                lst.Add(xs2!);
            }

            if (lst.Count != 0)
            {
                ret.Add(lst);
            }
        }
        return ret;
    }

    private static bool TryCollectMergedFlowSheetRegionInSheet(XSSFWorkbook excel, int sheetIndex, out List<RegionInfo> outRegionInfos)
    {
        ISheet sheet = excel.GetSheetAt(sheetIndex);
        string sheetName = sheet.SheetName;
        if (sheetName.StartsWith('_'))
        {
            outRegionInfos = [];
            return false;
        }

        IRow rowLast = sheet.GetRow(sheet.LastRowNum);
        int mergedRegionsCount = sheet.NumMergedRegions;
        List<RegionInfo> ret = new List<RegionInfo>(capacity: mergedRegionsCount);
        for (int regionIndex = 0; regionIndex < mergedRegionsCount; regionIndex++)
        {
            CellRangeAddress mergedRegion = sheet.GetMergedRegion(regionIndex);
            int titleWidth = mergedRegion.MaxColumn - mergedRegion.FirstColumn;

            int startX = mergedRegion.FirstColumn;
            int startY = mergedRegion.FirstRow;
            ICell firstCell = sheet.CellOrNullXY(startX, startY)!;
            if (firstCell.CellType != CellType.String)
            {
                continue;
            }

            string rawTitle = firstCell.StringCellValue;
            RegionInfo.E_REGION_TYPE regionType = GetRegionType(rawTitle);
            if (regionType == RegionInfo.E_REGION_TYPE.NONE)
            {
                continue;
            }

            int endX = mergedRegion.MaxColumn;
            if (!sheet.TryScanDown("^END", startX, mergedRegion.FirstRow, rowLast.RowNum, out int endY))
            {
                continue;
            }

            Rect rect = new Rect(startX, startY, endX, endY);
            RegionInfo fsr = RegionInfo.New(sheetIndex, sheetName, rawTitle, rect, regionType);
            ret.Add(fsr);
        }

        if (ret.Count == 0)
        {
            outRegionInfos = [];
            return false;
        }

        outRegionInfos = ret;
        return true;
    }

    private static bool TryCollectDefaultClassFlowSheetRegionInSheet(XSSFWorkbook excel, int sheetIndex, out RegionInfo? outRegion)
    {
        ISheet sheet = excel.GetSheetAt(sheetIndex);
        IRow row0 = sheet.GetRow(0);
        int2 max = new int2(row0.LastCellNum, sheet.LastRowNum);

        if (!sheet.TryScanDown("^END", 0, 0, max.y, out int endY))
        {
            outRegion = null;
            return false;
        }

        int endX = -1;
        foreach (E_HEADER_CLASS e in ExEnum.GetValues<E_HEADER_CLASS>())
        {
            if (sheet.TryScanRight($"^{e}", 0, 0, max.x, out endX))
            {
                break;
            }
        }

        if (endX == -1)
        {
            outRegion = null;
            return false;
        }

        Rect r = new Rect(0, 0, endX, endY);
        RegionInfo region = RegionInfo.New(sheetIndex, sheet.SheetName, string.Empty, r, RegionInfo.E_REGION_TYPE.CLASS);
        outRegion = region;
        return true;
    }

    private static RegionInfo.E_REGION_TYPE GetRegionType(string rawTitle)
    {
        if (rawTitle.StartsWith("|const|", StringComparison.OrdinalIgnoreCase))
        {
            return RegionInfo.E_REGION_TYPE.CONST;
        }

        if (rawTitle.StartsWith("|enum|", StringComparison.OrdinalIgnoreCase))
        {
            return RegionInfo.E_REGION_TYPE.ENUM;
        }

        if (rawTitle.StartsWith("|class|", StringComparison.OrdinalIgnoreCase))
        {
            return RegionInfo.E_REGION_TYPE.CLASS;
        }

        return RegionInfo.E_REGION_TYPE.NONE;
    }
}
