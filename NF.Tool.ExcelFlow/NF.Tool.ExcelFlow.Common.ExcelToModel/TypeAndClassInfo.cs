using NF.Tool.ExcelFlow.Common.Model.Infos;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace NF.Tool.ExcelFlow.Common.ExcelToModel;

public sealed class TypeAndClassInfo
{
    public string ExcelFilePath { get; }
    public int SheetIndex { get; }
    public ClassInfo ClassInfo { get; }
    public Type Class { get; }
    public Dictionary<string, MemberInfo> MemberDic { get; }

    public TypeAndClassInfo(string excelFilePath, int sheetIndex, ClassInfo classInfo, Type cls)
    {
        ExcelFilePath = excelFilePath;
        SheetIndex = sheetIndex;
        ClassInfo = classInfo;
        Class = cls;

        const BindingFlags bindingFlags = BindingFlags.Public | BindingFlags.Instance;

        MemberInfo[] members = cls.GetFields(bindingFlags)
            .Cast<MemberInfo>()
            .Concat(cls.GetProperties(bindingFlags))
            .ToArray();
        MemberDic = members.ToDictionary(x => x.Name, x => x);

    }

    public MemberInfo GetMemberInfo(string name)
    {
        return MemberDic[name];
    }
}
