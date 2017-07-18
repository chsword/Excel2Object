using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Reflection;
using CTRC;

namespace Chsword.Excel2Object.Internal
{
    internal class ExcelUtil
    {
        public static Dictionary<PropertyInfo, ExcelTitleAttribute> GetExportAttrDict<T>()
        {
            var dict = new Dictionary<PropertyInfo, ExcelTitleAttribute>();
            foreach (var propertyInfo in CTRCHelper.GetPropertiesCache<T>())
            {
                var attr = propertyInfo.GetCustomAttributes(true)
                    .FirstOrDefault(c => c is ExcelTitleAttribute || c is DisplayAttribute);
                if (attr != null)
                {
                    var attr1 = attr;
                    if (attr is DisplayAttribute)
                    {
                        var display = attr as DisplayAttribute;
                        attr1 = new ExcelTitleAttribute(display.Name) {Order = display.Order};
                    }
                    dict.Add(propertyInfo, attr1 as ExcelTitleAttribute);
                }
            }
            return dict;
        }
    }
}