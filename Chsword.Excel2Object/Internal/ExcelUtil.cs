using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Reflection;
using CTRC;

namespace Chsword.Excel2Object.Internal
{
    internal class ExcelUtil
    {

		/// <summary>
		/// get ExcelTitleAttribute in Attributes
		/// </summary>
		/// <param name="attrs"></param>
		/// <returns></returns>
		private static ExcelTitleAttribute GetExcelTitleAttributeFromAttributes(object[] attrs)
	    {
		    var attr = attrs.FirstOrDefault(c => c is ExcelTitleAttribute || c is DisplayAttribute);
		    if (attr == null) return null;

		    if (!(attr is DisplayAttribute display)) return attr as ExcelTitleAttribute;
		    return new ExcelTitleAttribute(display.Name)
		    {
			    Order = display.Order
		    };

	    }

		/// <summary>
		/// Get the ExcelTitleAttribute on class
		/// </summary>
		/// <typeparam name="T"></typeparam>
		/// <returns>if thers's not a ExcelTitleAttribute, will return null.</returns>
		public static ExcelTitleAttribute GetClassExportAttribute<T>()
		{
			var attrs = typeof(T).GetCustomAttributes(true);
			var attr = GetExcelTitleAttributeFromAttributes(attrs);
			return attr;
		}

	    /// <summary>
	    /// Get the ExcelTitleAttribute on proerties
	    /// </summary>
	    /// <typeparam name="T"></typeparam>
	    /// <returns></returns>
	    public static Dictionary<PropertyInfo, ExcelTitleAttribute> GetPropertiesAttributesDict<T>()
	    {
		    var dict = new Dictionary<PropertyInfo, ExcelTitleAttribute>();
		    foreach (var propertyInfo in CTRCHelper.GetPropertiesCache<T>())
		    {
			    var attrs = propertyInfo.GetCustomAttributes(true);
			    var attr = GetExcelTitleAttributeFromAttributes(attrs);

			    if (attr == null) continue;
			    dict.Add(propertyInfo, attr);
		    }

		    return dict;
	    }
    }
}