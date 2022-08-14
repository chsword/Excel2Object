using System.ComponentModel.DataAnnotations;
using System.Reflection;

namespace Chsword.Excel2Object.Internal;

internal static class ExcelUtil
{
    /// <summary>
    ///     Get the ExcelTitleAttribute on class
    /// </summary>
    /// <typeparam name="T"></typeparam>
    /// <returns>if there's not a ExcelTitleAttribute, will return null.</returns>
    public static ExcelTitleAttribute? GetClassExportAttribute<T>()
    { 
        var attrs = typeof(T).GetCustomAttributes(true);
        var attr = GetExcelTitleAttributeFromAttributes(attrs, 0);
        return attr;
    }

    /// <summary>
    ///     Get the ExcelTitleAttribute on properties
    /// </summary>
    /// <typeparam name="T"></typeparam>
    /// <returns></returns>
    public static Dictionary<PropertyInfo, ExcelTitleAttribute> GetPropertiesAttributesDict<T>()
    {
        var dict = new Dictionary<PropertyInfo, ExcelTitleAttribute>();
        var defaultOrder = 10000;
        var props = typeof(T).GetTypeInfo().GetRuntimeProperties();
        foreach (var propertyInfo in props)
        {
            var attrs = propertyInfo.GetCustomAttributes(true);
            var attr = GetExcelTitleAttributeFromAttributes(attrs, defaultOrder++);
            if (attr == null) continue;
            dict.Add(propertyInfo, attr);
        }

        return dict;
    }

    /// <summary>
    ///     get ExcelTitleAttribute in Attributes
    /// </summary>
    /// <param name="attrs"></param>
    /// <param name="defaultOrder">如果未设置Order则使用此Order</param>
    /// <returns></returns>
    private static ExcelTitleAttribute? GetExcelTitleAttributeFromAttributes(object[] attrs, int defaultOrder)
    {
        var attrTitle = attrs.FirstOrDefault(c => c is ExcelTitleAttribute or DisplayAttribute);
        if (attrTitle == null) return null;
        if (attrTitle is DisplayAttribute display)
            return new ExcelTitleAttribute(display.Name)
            {
                Order = display.GetOrder() ?? defaultOrder
            };

        var attrResult = attrTitle as ExcelTitleAttribute;
        if (attrResult?.Order == 0) attrResult.Order = defaultOrder;

        return attrResult;
    }
}