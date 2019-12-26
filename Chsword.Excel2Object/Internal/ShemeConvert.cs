using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using NPOI.SS.UserModel;

namespace Chsword.Excel2Object.Internal
{
    internal class Sheme
    {
        public static SheetModel GetSheetSheme<TModel>()
        {
            var classExportAttribute = ExcelUtil.GetClassExportAttribute<TModel>();
            var sheet = SheetModel.Create(classExportAttribute.Title);
            var dict = ExcelUtil.GetPropertiesAttributesDict<TModel>();
            foreach (var prop in dict)
            {
                sheet.Columns.Add(new ExcelColumn
                {
                    Order = prop.Value.Order,
                    Title = prop.Value.Title,
                    Type = prop.Key.PropertyType
                });
            }
            return sheet;
        }
    }
}