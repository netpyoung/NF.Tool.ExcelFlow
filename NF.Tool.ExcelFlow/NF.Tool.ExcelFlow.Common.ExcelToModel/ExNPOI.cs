using NF.Tool.ExcelFlow.Common.Model;
using NPOI.SS.UserModel;
using System;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;

namespace NF.Tool.ExcelFlow.Common.ExcelToModel;

public static class ExNPOI
{
    public static bool TryScanDown(this ISheet sheet, string scanString, int startX, int startY, int endY, out int outScanY)
    {
        for (int y = startY; y <= endY; ++y)
        {
            ICell? cellOrNull = sheet.CellOrNullXY(startX, y);
            if (cellOrNull == null)
            {
                continue;
            }

            ICell cell = cellOrNull;
            if (cell.CellType != CellType.String)
            {
                continue;
            }

            string value = cell.StringCellValue;
            if (!value.Equals(scanString, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            outScanY = y;
            return true;
        }

        outScanY = -1;
        return false;
    }

    public static bool TryScanRight(this ISheet sheet, string scanString, int startX, int startY, int endX, out int outScanX)
    {
        for (int x = startX; x <= endX; ++x)
        {
            ICell? cellOrNull = sheet.CellOrNullXY(x, startY);
            if (cellOrNull == null)
            {
                continue;
            }

            ICell cell = cellOrNull;
            if (cell.CellType != CellType.String)
            {
                continue;
            }

            string value = cell.StringCellValue;
            if (!value.Equals(scanString, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            outScanX = x;
            return true;
        }

        outScanX = -1;
        return false;
    }
    public static ICell? CellOrNullXY([NotNull] this ISheet sheet, int x, int y)
    {
        IRow rowY = sheet.GetRow(y);
        if (rowY == null)
        {
            return null;
        }

        ICell cell = rowY.GetCell(x);
        return cell;
    }

    public static string ToA1([NotNull] this ICell cell)
    {
        return ToA1(cell.Address.Column, cell.Address.Row);
    }

    public static string ToA1(in this int2 xy)
    {
        return ToA1(xy.x, xy.y);
    }

    public static string ToA1(int x, int y)
    {
        string col = "";
        for (x++; x > 0; x = (x - 1) / 26)
        {
            col = (char)('A' + (x - 1) % 26) + col;
        }
        return col + (y + 1);
    }

    public static object? GetValue(this ICell cell, in Type type, in IFormulaEvaluator evaluator)
    {
        if (cell == null)
        {
            return null;
        }

        if (cell.CellType == CellType.Blank)
        {
            return null;
        }

        if (type == typeof(string))
        {
            if (cell.CellType == CellType.Numeric)
            {
                double cellVal = cell.NumericCellValue;
                int convertedVal = Convert.ToInt32(cell.NumericCellValue);

                if (convertedVal != cellVal)
                {
                    return cellVal.ToString();
                }

                return convertedVal.ToString();
            }

            cell.SetCellType(CellType.String);
            if (string.IsNullOrEmpty(cell.StringCellValue))
            {
                return "";
            }

            return cell.StringCellValue;
        }

        if (type == typeof(float))
        {
            if (cell.CellType == CellType.Numeric)
            {
                return Convert.ToSingle(cell.NumericCellValue);
            }

            try
            {
                return Convert.ToSingle(GetStringVal(cell, evaluator));
            }
            catch (Exception e)
            {
                Console.Error.WriteLine(e);
                Console.Error.WriteLine($"{cell.Sheet.SheetName}: {cell.RowIndex + 1}/{cell.ColumnIndex + 1} | {cell}({type})");
                return 0;
            }
        }

        if (type == typeof(int))
        {
            try
            {
                if (cell.CellType == CellType.Numeric)
                {
                    return Convert.ToInt32(cell.NumericCellValue);
                }

                return Convert.ToInt32(GetStringVal(cell, evaluator));
            }
            catch (Exception e)
            {
                Console.Error.WriteLine(e);
                Console.Error.WriteLine($"{cell.Sheet.SheetName}: {cell.RowIndex + 1}/{cell.ColumnIndex + 1} | {cell}({type})");
                return 0;
            }
        }

        if (type == typeof(double))
        {
            return Convert.ToDouble(GetStringVal(cell, evaluator));
        }

        if (type == typeof(long))
        {
            return Convert.ToDouble(GetStringVal(cell, evaluator));
        }

        if (type == typeof(bool))
        {
            return Convert.ToBoolean(GetStringVal(cell, evaluator));
        }

        if (type == typeof(DateTime))
        {
            return DateTime.Parse(GetStringVal(cell, evaluator));
        }

        if (type == typeof(DateTimeOffset))
        {
            return DateTimeOffset.Parse(GetStringVal(cell, evaluator));
        }

        if (type == typeof(TimeSpan))
        {
            return TimeSpan.Parse(GetStringVal(cell, evaluator));
        }

        if (type.IsEnum)
        {
            //return 5;
            //return System.Convert.ToInt32(System.Convert.ToDouble(GetStringVal(cell, evaluator)));
            cell.SetCellType(CellType.String);
            return Convert.ToInt32(Enum.Parse(type, cell.StringCellValue));
        }

        return null;
    }

    private static string GetStringVal(in ICell cell, in IFormulaEvaluator evaluator)
    {
        switch (cell.CellType)
        {
            case CellType.Formula:
                switch (cell.CachedFormulaResultType)
                {
                    case CellType.Numeric:
                        return cell.NumericCellValue.ToString(CultureInfo.InvariantCulture);

                    case CellType.String:
                        return cell.StringCellValue;

                    default:
                        return evaluator.Evaluate(cell).FormatAsString();
                }

            case CellType.String:
                return cell.StringCellValue;

            default:
                cell.SetCellType(CellType.String);
                return cell.StringCellValue;
        }
    }

}