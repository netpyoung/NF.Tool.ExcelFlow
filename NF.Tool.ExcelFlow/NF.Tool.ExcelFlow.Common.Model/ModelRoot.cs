using NF.Tool.ExcelFlow.Common.Model.Infos;
using System.Collections.Generic;

namespace NF.Tool.ExcelFlow.Common.Model;

public sealed class ModelRoot
{
    public string ExcelFilePath { get; }
    public string ExcelFileName { get; }
    public int SheetIndex { get; }
    public string SheetName { get; }
    public List<RegionInfo> RegionInfos { get; }

    public List<ClassInfo> ClassInfos { get; }
    public List<ConstInfo> ConstInfos { get; }
    public List<EnumInfo> EnumInfos { get; }

    public ModelRoot(string excelFilePath, string excelFileName, int sheetIndex, string sheetName, List<RegionInfo> regionInfos, List<ClassInfo> classInfos, List<ConstInfo> constInfos, List<EnumInfo> enumInfos)
    {
        ExcelFilePath = excelFilePath;
        ExcelFileName = excelFileName;
        SheetIndex = sheetIndex;
        SheetName = sheetName;
        RegionInfos = regionInfos;
        ClassInfos = classInfos;
        ConstInfos = constInfos;
        EnumInfos = enumInfos;
    }
}
