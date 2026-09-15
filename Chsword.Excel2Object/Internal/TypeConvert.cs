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
        var columns = title.Keys.Select((c, i) => new ExcelColumn(c, typeof(string)) {Order = i}).ToList();

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
            var column = new ExcelColumn(titleAttr.Title, objKeysArray[i].Key.PropertyType) {Order = i};
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
                // 值为 null 的属性不入字典：导出时取不到值与取到 null 同样写成空白单元格
                var value = column.Key.GetValue(item, null);
                if (value != null) row[column.Value.Title] = value;
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

        var columns = dataSetColumnArray
            .Select((item, i) => new ExcelColumn(item.ColumnName, item.DataType) {Order = i}).ToList();
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
            // FormulaColumns 加入时已校验标题，但该集合与 FormulaColumn.Title 都是公开可写的，
            // 加入之后仍可改回 null，故在此处消费前再确认一次
            var title = formulaColumn.Title;
            if (title == null || title.Trim().Length == 0)
                throw new Excel2ObjectException("公式列必须有标题：它既是表头上的名字，也是与模型列对应的依据。");

            var excelColumn = columns.FirstOrDefault(c => c.Title == title);
            if (excelColumn == null)
            {
                excelColumn = new ExcelColumn(title, typeof(Expression))
                {
                    Order = 0,
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
            if (options.Dropdowns.TryGetValue(column.Title, out var values))
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