using System;

namespace Chsword.Excel2Object.Internal
{
    public class TypeUtil
    {
        /// <summary>
        /// Get a real type of  a Nullable type
        /// </summary>
        /// <param name="conversionType"></param>
        /// <returns></returns>
        public static Type GetUnNullableType(Type conversionType)
        {
            if (conversionType.IsGenericType && conversionType.GetGenericTypeDefinition() == typeof(Nullable<>))
            {
                var nullableConverter = new System.ComponentModel.NullableConverter(conversionType);
                conversionType = nullableConverter.UnderlyingType;
            }

            return conversionType;
        }
    }
}