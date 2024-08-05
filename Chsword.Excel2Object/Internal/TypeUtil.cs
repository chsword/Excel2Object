﻿using System.ComponentModel;

namespace Chsword.Excel2Object.Internal;

public static class TypeUtil
{
    /// <summary>
    ///     Get a real type of Nullable type
    /// </summary>
    /// <param name="conversionType"></param>
    /// <returns></returns>
    public static Type GetUnNullableType(Type conversionType)
    {
        if (conversionType.IsGenericType && conversionType.GetGenericTypeDefinition() == typeof(Nullable<>))
        {
            var nullableConverter = new NullableConverter(conversionType);
            conversionType = nullableConverter.UnderlyingType;
        }

        return conversionType;
    }
}