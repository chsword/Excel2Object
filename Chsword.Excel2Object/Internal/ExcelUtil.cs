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
        var defaultOrder = 10000;
        var props = typeof(T).GetTypeInfo().GetRuntimeProperties();
        
        return props
            .Select(propertyInfo => new
            {
                Property = propertyInfo,
                Attribute = GetExcelTitleAttributeFromAttributes(propertyInfo.GetCustomAttributes(true).ToArray(), defaultOrder++)
            })
            .Where(x => x.Attribute != null)
            .ToDictionary(x => x.Property, x => x.Attribute!);
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
            return new ExcelTitleAttribute(display.Name!)
            {
                Order = display.GetOrder() ?? defaultOrder
            };

        var attrResult = attrTitle as ExcelTitleAttribute;
        if (attrResult?.Order == 0) attrResult.Order = defaultOrder;

        return attrResult;
    }
}