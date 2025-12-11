using NF.Tool.ExcelFlow.Common.Model;
using NF.Tool.ExcelFlow.Common.Model.Infos;
using NF.Tool.ExcelFlow.Common.Model.Intermediary;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace NF.Tool.ExcelFlow.Common.ExcelToModel;

public static class ExcelToModelBaker
{
    public static async Task<(ModelRoot[] vals, string? errOrNull)> Bake(string[] excelPaths)
    {
        ConcurrentBag<ModelRoot> bagModelRoot = [];
        ConcurrentBag<string> bagErr = [];
        await Parallel.ForEachAsync(excelPaths, async (path, ct) =>
        {
            (List<ModelRoot> list, string? errOrNull) = Import(path);
            if (errOrNull != null)
            {
                bagErr.Add(errOrNull);
            }
            else
            {
                foreach (ModelRoot item in list)
                {
                    bagModelRoot.Add(item);
                }
            }
        });

        string[] errs = bagErr.ToArray();
        if (errs.Length != 0)
        {
            return ([], Err.MessageWithInfo(string.Join("    \n", errs)));
        }

        ModelRoot[] models = bagModelRoot.ToArray();
        if (models.Length == 0)
        {
            return ([], Err.MessageWithInfo($"Fail: 0 Models in\n\n{string.Join("    \n", excelPaths)}"));
        }
        return (models, null);
    }

    private static (List<ModelRoot> val, string? errOrNull) Import(string excelPath)
    {
        using (FileStream fileStream = File.Open(excelPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
        using (XSSFWorkbook excel = new XSSFWorkbook(fileStream, readOnly: true))
        {
            List<List<RegionInfo>> sheetRegionss = Parser_RegionInfo.CollectRegionsFromExel(excel);
            List<ModelRoot> ret = new List<ModelRoot>(capacity: sheetRegionss.Count);

            StringBuilder sb = new StringBuilder();
            foreach (List<RegionInfo> sheetRegions in sheetRegionss)
            {
                (ModelRoot? modelRootOrNull, string? errOrNull) = GetModelRootOrNull(excel, excelPath, sheetRegions);
                if (errOrNull != null)
                {
                    sb.AppendLine(errOrNull);
                    continue;
                }

                ret.Add(modelRootOrNull!);
            }

            string err = sb.ToString();
            if (!string.IsNullOrEmpty(err))
            {
                return ([], err);
            }
            return (ret, null);
        }
    }

    private static (ModelRoot? modelRootOrNull, string? errOrNull) GetModelRootOrNull(XSSFWorkbook excel, string excelPath, List<RegionInfo> regionInfos)
    {
        bool isChanged = false;

        List<ClassInfo> classInfos = new List<ClassInfo>(capacity: regionInfos.Count);
        List<ConstInfo> constInfos = new List<ConstInfo>(capacity: regionInfos.Count);
        List<EnumInfo> enumInfos = new List<EnumInfo>(capacity: regionInfos.Count);
        StringBuilder sbErr = new StringBuilder();
        foreach (RegionInfo sheetRegion in regionInfos)
        {
            switch (sheetRegion.RegionType)
            {
                case RegionInfo.E_REGION_TYPE.CONST:
                    {
                        (ConstInfo? infoOrNull, string? errOrNull) = Parser_Const.Parse(excel, sheetRegion);
                        if (errOrNull != null)
                        {
                            sbErr.AppendLine(errOrNull);
                            break;
                        }
                        constInfos.Add(infoOrNull!);
                        isChanged = true;
                    }
                    break;
                case RegionInfo.E_REGION_TYPE.ENUM:
                    {
                        (EnumInfo? infoOrNull, string? errOrNull) = Parser_Enum.TryParse(excel, sheetRegion);
                        if (errOrNull != null)
                        {
                            sbErr.AppendLine(errOrNull);
                            break;
                        }
                        enumInfos.Add(infoOrNull!);
                        isChanged = true;
                    }
                    break;
                case RegionInfo.E_REGION_TYPE.CLASS:
                    {
                        (ClassInfo? infoOrNull, string? errOrNull) = Parser_Class.TryParse(excel, sheetRegion);
                        if (errOrNull != null)
                        {
                            sbErr.AppendLine(errOrNull);
                            break;
                        }

                        classInfos.Add(infoOrNull!);
                        isChanged = true;
                    }
                    break;
                case RegionInfo.E_REGION_TYPE.NONE:
                default:
                    break;
            }
        }

        string err = sbErr.ToString();
        if (!string.IsNullOrEmpty(err))
        {
            return (null, err);
        }

        if (!isChanged)
        {
            return (null, Err.MessageWithInfo("not found"));
        }

        int sheetIndex = regionInfos.First().SheetIndex;
        string sheetName = regionInfos.First().SheetName;
        string fileName = Path.GetFileName(excelPath);
        ModelRoot modelRoot = new ModelRoot(excelPath, fileName, sheetIndex, sheetName, regionInfos, classInfos, constInfos, enumInfos);
        return (modelRoot, null);
    }

    public static async Task<(ClassAndRows[] val, string? errOrNull)> CollectRows(ModelRoot[] models, Assembly assembly)
    {
        List<TypeAndClassInfo> collectList = MappingClassInfo(models, assembly);

        Dictionary<string, (FileStream fs, IWorkbook workbook)> dicWorkBook = new Dictionary<string, (FileStream, IWorkbook)>(capacity: models.Length);
        foreach (ModelRoot m in models)
        {
            if (dicWorkBook.ContainsKey(m.ExcelFilePath))
            {
                continue;
            }

            FileStream fileStream = File.Open(m.ExcelFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            XSSFWorkbook workbook = new XSSFWorkbook(fileStream, readOnly: true);
            dicWorkBook[m.ExcelFilePath] = (fileStream, workbook);
        }

        ConcurrentBag<ClassAndRows> bag = [];
        await Parallel.ForEachAsync(collectList.GroupBy(x => x.ExcelFilePath), async (xs, _) =>
        {
            foreach (TypeAndClassInfo t in xs)
            {
                (FileStream fs, IWorkbook workbook) di = dicWorkBook[t.ExcelFilePath];
                List<object> rows = ExtractRows(di.workbook, t);
                bag.Add(new ClassAndRows(t.Class, rows));
            }
        });

        foreach ((FileStream fs, IWorkbook workbook) in dicWorkBook.Values)
        {
            workbook.Close();
            fs.Close();
        }

        ClassAndRows[] ret = bag.ToArray();
        return (ret, null);
    }

    private static List<TypeAndClassInfo> MappingClassInfo(ModelRoot[] ms, Assembly assembly)
    {
        Type[] classArray = assembly.GetTypes().Where(x => x.IsClass).ToArray();
        List<TypeAndClassInfo> xxs = new List<TypeAndClassInfo>(capacity: classArray.Length);

        foreach (ModelRoot m in ms)
        {
            foreach (ClassInfo ci in m.ClassInfos)
            {
                foreach (Type cls in classArray)
                {
                    if (ci.ClassName.Value == cls.Name)
                    {
                        TypeAndClassInfo h = new TypeAndClassInfo(m.ExcelFilePath, m.SheetIndex, ci, cls);
                        xxs.Add(h);
                    }
                }
            }
        }
        return xxs;
    }
    private static List<object> ExtractRows(IWorkbook workbook, TypeAndClassInfo t)
    {
        ISheet sheet = workbook.GetSheetAt(t.SheetIndex);
        Type type = t.Class;

        int yMin = t.ClassInfo.HeaderColumns.First().Name.Pos.y + 1;
        int yMax = t.ClassInfo.RegionInfo.Region.BottomRight.y;
        IFormulaEvaluator evaluator = workbook.GetCreationHelper().CreateFormulaEvaluator();

        List<object> rows = new List<object>(capacity: yMax - yMin);
        for (int y = yMin; y < yMax; ++y)
        {
            IRow row = sheet.GetRow(y);
            object item = Activator.CreateInstance(type)!;
            foreach (HeaderColumnClass hc in t.ClassInfo.HeaderColumns)
            {
                int x = hc.Name.Pos.x;
                ICell cell = row.GetCell(x);
                if (cell == null)
                {
                    continue;
                }

                string name = hc.Name.Name;
                MemberInfo member = t.GetMemberInfo(name);
                switch (member.MemberType)
                {
                    case MemberTypes.Field:
                        {
                            object? value = cell.GetValue(((FieldInfo)member).FieldType, evaluator);
                            if (value == null)
                            {
                                continue;
                            }
                            ((FieldInfo)member).SetValue(item, value);
                        }
                        break;
                    case MemberTypes.Property:
                        {
                            object? value = cell.GetValue(((PropertyInfo)member).PropertyType, evaluator);
                            if (value == null)
                            {
                                continue;
                            }
                            ((PropertyInfo)member).SetValue(item, value, null);
                        }
                        break;
                }
            }
            rows.Add(item);
        }

        return rows;
    }
}
