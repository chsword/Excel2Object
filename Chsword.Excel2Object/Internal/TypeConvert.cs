using System.Data;
using System.Linq.Expressions;
using System.Reflection;
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
        // 列由首行的键决定，故首行须先取出；其余各行留待写入时逐行取用，以免整份数据都进内存
        var rows = data.GetEnumerator();
        bool any;
        try
        {
            any = rows.MoveNext();
        }
        catch
        {
            rows.Dispose();
            throw;
        }

        if (!any)
        {
            rows.Dispose();
            return excel;
        }

        var first = rows.Current;
        var columns = first.Keys.Select((c, i) => new ExcelColumn(c, typeof(string)) {Order = i}).ToList();

        sheet.Columns = AttachColumns(columns, options);
        sheet.Rows = FirstThenRest(first, rows);
        excel.RowSource = rows;

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
        sheet.Rows = ToRows(data, objKeysArray);

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

        sheet.Rows = ToRows(dt, dataSetColumnArray);

        return excel;
    }

    /// <summary>逐行取值，写一行取一行，而非先把整份数据摊成字典列表。</summary>
    private static IEnumerable<Dictionary<string, object>> ToRows<TModel>(IEnumerable<TModel> data,
        KeyValuePair<PropertyInfo, ExcelTitleAttribute>[] columns)
    {
        foreach (var item in data)
        {
            if (item == null) continue;

            var row = new Dictionary<string, object>();
            foreach (var column in columns)
            {
                // 值为 null 的属性不入字典：导出时取不到值与取到 null 同样写成空白单元格
                var value = column.Key.GetValue(item, null);
                if (value != null) row[column.Value.Title] = value;
            }

            yield return row;
        }
    }

    private static IEnumerable<Dictionary<string, object>> ToRows(DataTable dt, DataColumn[] columns)
    {
        foreach (DataRow item in dt.Rows)
        {
            if (item == null) continue;

            var row = new Dictionary<string, object>();
            foreach (var column in columns) row[column.ColumnName] = item[column.ColumnName];

            yield return row;
        }
    }

    /// <summary>把已取出的首行与余下各行接回一条序列，并在遍历结束时释放枚举器。</summary>
    private static IEnumerable<Dictionary<string, object>> FirstThenRest(Dictionary<string, object> first,
        IEnumerator<Dictionary<string, object>> rest)
    {
        using (rest)
        {
            yield return first;
            while (rest.MoveNext()) yield return rest.Current;
        }
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