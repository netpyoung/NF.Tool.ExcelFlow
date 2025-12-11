using NF.Tool.ExcelFlow.Common.Model.Attributes;
using NF.Tool.ExcelFlow.Common.Model.Infos;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Diagnostics.CodeAnalysis;
using System.Reflection;

namespace NF.Tool.ExcelFlow.Common.Model;

public static class ExEnum
{
    private readonly static Dictionary<Type, bool> _cache_StepDown = [];
    private readonly static Dictionary<Type, bool> _cache_StepRight = [];
    private readonly static Dictionary<Enum, bool> _cache_Required = [];
    private readonly static Dictionary<Enum, int> _cache_Position = [];
    private readonly static Dictionary<Enum, bool> _cache_Table = [];
    private readonly static Dictionary<Type, Array> _cache_EnumValues = [];

    static ExEnum()
    {
        foreach (Type type in new Type[] { typeof(E_HEADER_CONST), typeof(E_HEADER_ENUM), typeof(E_HEADER_CLASS), typeof(E_FIELD_ENUM), typeof(E_FIELD_CONST) })
        {
            _cache_EnumValues[type] = Enum.GetValues(type);
        }

        foreach (Type type in new Type[] { typeof(E_HEADER_CONST), typeof(E_HEADER_ENUM), typeof(E_HEADER_CLASS), typeof(E_FIELD_ENUM), typeof(E_FIELD_CONST) })
        {
            IsStepDown(type);
            IsStepRight(type);
        }

        foreach (Type type in new Type[] { typeof(E_HEADER_CONST), typeof(E_HEADER_ENUM), typeof(E_HEADER_CLASS), typeof(E_FIELD_ENUM), typeof(E_FIELD_CONST) })
        {
            foreach (Enum val in GetValues(type))
            {
                IsTable(val);
                GetPosition(val);
                IsRequired(val);
            }
        }
    }

    public static T[] GetValues<T>()
    {
        return (T[])GetValues(typeof(T));
    }

    private static Array GetValues(Type e)
    {
        return _cache_EnumValues[e];
    }

    public static bool IsStepDown([NotNull] this Type t)
    {
        if (_cache_StepDown.TryGetValue(t, out bool cachedIsStepDown))
        {
            return cachedIsStepDown;
        }

        bool isStepDown = t.GetCustomAttribute<StepDownAttribute>() != null;
        _cache_StepDown[t] = isStepDown;
        return isStepDown;
    }

    public static bool IsStepRight([NotNull] this Type t)
    {
        if (_cache_StepRight.TryGetValue(t, out bool cachedIsStepRight))
        {
            return cachedIsStepRight;
        }

        bool isStepRight = t.GetCustomAttribute<StepRightAttribute>() != null;
        _cache_StepRight[t] = isStepRight;
        return isStepRight;
    }

    public static bool IsTable([NotNull] this Enum e)
    {
        if (_cache_Table.TryGetValue(e, out bool cachedIsTable))
        {
            return cachedIsTable;
        }

        Type type = e.GetType();
        string name = Enum.GetName(type, e)!;
        FieldInfo field = type.GetField(name)!;

        bool isTable = field.GetCustomAttribute<TableAttribute>() != null;
        _cache_Table[e] = isTable;
        return isTable;
    }

    public static int GetPosition([NotNull] this Enum e)
    {
        if (_cache_Position.TryGetValue(e, out int cachedPosition))
        {
            return cachedPosition;
        }

        Type type = e.GetType();
        string name = Enum.GetName(type, e)!;
        FieldInfo field = type.GetField(name)!;

        PositionAttribute? attrOrNull = field.GetCustomAttribute<PositionAttribute>();
        if (attrOrNull == null)
        {
            _cache_Position[e] = -1;
            return -1;
        }

        PositionAttribute attr = attrOrNull;
        int position = attr.Index;
        _cache_Position[e] = position;
        return position;
    }

    public static bool IsRequired([NotNull] this Enum e)
    {
        if (_cache_Required.TryGetValue(e, out bool cachedIsRequired))
        {
            return cachedIsRequired;
        }

        if (IsTable(e))
        {
            _cache_Required[e] = true;
            return true;
        }

        Type type = e.GetType();
        string name = Enum.GetName(type, e)!;
        FieldInfo field = type.GetField(name)!;

        bool isRequired = field.GetCustomAttribute<RequiredAttribute>() != null;
        _cache_Required[e] = isRequired;
        return isRequired;
    }
}
