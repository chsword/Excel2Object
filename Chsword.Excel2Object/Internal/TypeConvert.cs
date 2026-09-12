using System.Data;
using System.Linq.Expressions;
using Chsword.Excel2Object.Options;

namespace Chsword.Excel2Object.Internal;

internal static class TypeConvert
{
    public static ExcelModel ConvertDictionaryToExcelModel(IEnumerable<Dictionary<string, object>> data,
        ExcelExporterOptions options)
    {
        var sheetTitle = options.SheetTitle;
        var excel = new ExcelModel {Sheets = new List<SheetModel>()};
        var sheet = SheetModel.Create(sheetTitle);
        excel.Sheets.Add(sheet);
        var list = data.ToList();
        var title = list.FirstOrDefault();
        if (title == null) return excel;
        var columns = title.Keys.Select((c, i) => new ExcelColumn
        {
            Order = i,
            Title = c,
            Type = typeof(string)
        }).ToList();

        sheet.Columns = AttachColumns(columns, options);
        sheet.Rows = list;

        return excel;
    }

    public static ExcelModel ConvertObjectToExcelModel<TModel>(IEnumerable<TModel> data,
        ExcelExporterOptions options)
    {
        var sheetTitle = options.SheetTitle;
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

        var columns = new List<ExcelColumn>();
        for (var i = 0; i < objKeysArray.Length; i++)
        {
            var titleAttr = objKeysArray[i].Value;
            var column = new ExcelColumn
            {
                Title = titleAttr.Title,
                Type = objKeysArray[i].Key.PropertyType,
                Order = i
            };
            if (titleAttr is ExcelColumnAttribute excelColumnAttr)
            {
                column.CellStyle = excelColumnAttr;
                column.HeaderStyle = excelColumnAttr;
                column.Dropdown = excelColumnAttr.Dropdown;
            }

            columns.Add(column);
        }

        sheet.Columns = AttachColumns(columns, options);
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

    internal static ExcelModel ConvertDataSetToExcelModel(DataTable dt, ExcelExporterOptions options)
    {
        var sheetTitle = options.SheetTitle;
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
        sheet.Columns = AttachColumns(columns, options);

        var data = dt.Rows.Cast<DataRow>().ToArray();
        foreach (var item in data.Where(c => c != null))
        {
            var row = new Dictionary<string, object>();
            foreach (var column in dataSetColumnArray) row[column.ColumnName] = item[column.ColumnName];

            sheet.Rows.Add(row);
        }

        return excel;
    }

    private static List<ExcelColumn> AttachColumns(List<ExcelColumn> columns, ExcelExporterOptions options)
    {
        columns = columns.OrderBy(c => c.Order).ToList();
        foreach (var formulaColumn in options.FormulaColumns)
        {
            var excelColumn = columns.FirstOrDefault(c => c.Title == formulaColumn.Title);
            if (excelColumn == null)
            {
                excelColumn = new ExcelColumn
                {
                    Title = formulaColumn.Title,
                    Order = 0,
                    Type = typeof(Expression),
                    Formula = formulaColumn.ModelFormula ?? formulaColumn.Formula,
                    ResultType = ResultTypeOf(formulaColumn, null)
                };
                if (string.IsNullOrWhiteSpace(formulaColumn.AfterColumnTitle))
                {
                    columns.Add(excelColumn);
                }
                else
                {
                    var i = columns.FindIndex(c => c.Title == formulaColumn.AfterColumnTitle);
                    if (i < 0)
                        throw new Excel2ObjectException(
                            $"can not find {formulaColumn.AfterColumnTitle} column.");

                    columns.Insert(i + 1, excelColumn);
                }
            }
            else
            {
                // a formula taking over a model column yields that column's type unless told otherwise,
                // so a DateTime property keeps its date format
                excelColumn.ResultType = ResultTypeOf(formulaColumn, excelColumn.Type);
                excelColumn.Type = typeof(Expression);
                excelColumn.Formula = formulaColumn.ModelFormula ?? formulaColumn.Formula;
            }
        }

        // after the formula columns are in, so a column a formula added can carry a dropdown too
        foreach (var column in columns)
            if (column.Title != null && options.Dropdowns.TryGetValue(column.Title, out var values))
                column.Dropdown = values;

        for (var i = 0; i < columns.Count; i++) columns[i].Order = i * 10;

        return columns;
    }

    /// <summary>
    ///     What a formula column yields: what it was told, else the type of the column it takes over - so a
    ///     DateTime property keeps its date format. Nullable either way, and the exporter compares the type
    ///     itself, so the underlying one is what gets stored.
    /// </summary>
    private static Type? ResultTypeOf(FormulaColumn formulaColumn, Type? columnType)
    {
        var type = formulaColumn.FormulaResultType ?? columnType;
        return type == null ? null : TypeUtil.GetUnNullableType(type);
    }
}