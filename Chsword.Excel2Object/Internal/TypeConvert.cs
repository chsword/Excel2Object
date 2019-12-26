using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace Chsword.Excel2Object.Internal
{
    internal class TypeConvert
    {
        internal static ExcelModel ConvertDataSetToExcelModel(DataTable dt, string sheetTitle)
        {
            var excel = new ExcelModel {Sheets = new List<SheetModel>()};
            var sheet = SheetModel.Create(sheetTitle);
            excel.Sheets.Add(sheet);
            var dataSetColumnArray = dt.Columns.Cast<DataColumn>().ToArray();

            var columns = dataSetColumnArray.Select((item, i) =>
                new ExcelColumn
                {
                    Order = i,
                    Title = item.ColumnName,
                    Type = item.DataType
                }).ToList();
            sheet.Columns = columns;

            var data = dt.Rows.Cast<DataRow>().ToArray();
            foreach (var item in data.Where(c => c != null))
            {
                var row = new Dictionary<string, object>();
                foreach (var column in dataSetColumnArray)
                {
                    row[column.ColumnName] = item[column.ColumnName];
                }

                sheet.Rows.Add(row);
            }

            return excel;
        }

        public static ExcelModel ConvertDictionaryToExcelModel(IEnumerable<Dictionary<string, object>> data,
            string sheetTitle=null)
        {
            var excel = new ExcelModel { Sheets = new List<SheetModel>() };
            var sheet = SheetModel.Create(sheetTitle);
            excel.Sheets.Add(sheet);
            var list = data.ToList();
            var title = list.FirstOrDefault();
            if (title==null) return excel;

            sheet.Columns = title.Keys.Select((c, i) => new ExcelColumn
            {
                Order = i,
                Title = c,
                Type = typeof(string)
            }).ToList();
            sheet.Rows = list;

            return excel;
        }
        public static ExcelModel ConvertObjectToExcelModel<TModel>(IEnumerable<TModel> data, string sheetTitle)
        {
            var excel = new ExcelModel {Sheets = new List<SheetModel>()};
            if (string.IsNullOrWhiteSpace(sheetTitle))
            {
                var classAttr = ExcelUtil.GetClassExportAttribute<TModel>();
                sheetTitle = classAttr == null ? sheetTitle : classAttr.Title;
            }

            var sheet = SheetModel.Create(sheetTitle);

            excel.Sheets.Add(sheet);
            var attrDict = ExcelUtil.GetPropertiesAttributesDict<TModel>();
            var objKeysArray = attrDict.OrderBy(c => c.Value.Order).ToArray();


            for (var i = 0; i < objKeysArray.Length; i++)
            {
                var column = new ExcelColumn
                {
                    Title = objKeysArray[i].Value.Title, Type = objKeysArray[i].Key.PropertyType, Order = i
                };
                sheet.Columns.Add(column);
            }

            foreach (var item in data.Where(c => c != null))
            {
                var row = new Dictionary<string, object>();
                foreach (var column in objKeysArray)
                {
                    var prop = column.Key;
                    row[column.Value.Title] = prop.GetValue(item, null);
                }

                sheet.Rows.Add(row);
            }

            return excel;
        }
    }
}