using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Reflection;

namespace Chsword.Excel2Object.Internal
{
    internal class ExcelUtil
    {
        public static Dictionary<PropertyInfo, ExcelAttribute> GetExportAttrDict<T>()
        {
            var dict = new Dictionary<PropertyInfo, ExcelAttribute>();
            foreach (var propertyInfo in typeof(T).GetProperties())
            {
                var attr = propertyInfo.GetCustomAttributes(true).FirstOrDefault(c => c is ExcelAttribute || c is DisplayAttribute);
                if (attr != null)
                {
                    var attr1 = attr;
                    if (attr is DisplayAttribute)
                    {
                        var display = attr as DisplayAttribute;
                        attr1 = new ExcelAttribute(display.Name) { Order = display.Order };
                    }
                    dict.Add(propertyInfo, attr1 as ExcelAttribute);

                }
            }
            return dict;
        }
    }
}