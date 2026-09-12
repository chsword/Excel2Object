using NPOI.SS.UserModel;
using NPOI.SS.Util;

namespace Chsword.Excel2Object.Internal;

/// <summary>
///     Writes the dropdown a column declares - Excel's list validation - over the rows of that column.
/// </summary>
/// <remarks>
///     A short list is written into the validation itself. Excel only accepts 255 characters there
///     (and splits the list on commas), so a longer list, or one whose values carry a comma or a quote,
///     is written into a hidden sheet that a defined name points at, and the validation reads the name.
/// </remarks>
internal static class DropdownValidation
{
    /// <summary>The hidden sheet the long lists live on.</summary>
    private const string ListSheetName = "_excel2object_lists";

    /// <summary>The defined names pointing at those lists, numbered from one.</summary>
    private const string ListNamePrefix = "_excel2object_list";

    /// <summary>
    ///     What Excel accepts inside a list validation. The list is stored as one quoted literal, so the
    ///     two quotes around it count against the limit along with the commas between the values.
    /// </summary>
    private const int InlineLimit = 255;

    public static void Apply(ISheet sheet, ExcelColumn[] columns, int lastDataRowIndex)
    {
        // an export with no rows still gets its dropdown, so a template can be filled in
        var lastRow = Math.Max(lastDataRowIndex, ExcelConstants.DefaultDataStartRowIndex);
        var helper = sheet.GetDataValidationHelper();

        for (var i = 0; i < columns.Length; i++)
        {
            var values = columns[i].Dropdown;
            if (values == null || values.Length == 0) continue;
            Validate(values, columns[i].Title);

            var constraint = FitsInline(values)
                ? helper.CreateExplicitListConstraint(values)
                : helper.CreateFormulaListConstraint(WriteToListSheet(sheet.Workbook, values));

            var validation = helper.CreateValidation(constraint,
                new CellRangeAddressList(ExcelConstants.DefaultDataStartRowIndex, lastRow, i, i));
            // xlsx leaves the error box off unless it is asked for, and a dropdown that accepts
            // anything typed over it is not much of a dropdown
            validation.ShowErrorBox = true;
            sheet.AddValidationData(validation);
        }
    }

    private static void Validate(string[] values, string? title)
    {
        for (var i = 0; i < values.Length; i++)
            if (values[i] == null)
                throw new Excel2ObjectException(
                    $"Dropdown column [{title}] has no value at index {i}. Excel has no empty entry in a " +
                    "list; leave the cell blank instead of offering one.");
    }

    private static bool FitsInline(string[] values)
    {
        var length = values.Length - 1 + 2; // the commas between the values, and the quotes around them all
        foreach (var value in values)
        {
            if (value.IndexOf(',') >= 0 || value.IndexOf('"') >= 0) return false;
            length += value.Length;
        }

        return length <= InlineLimit;
    }

    /// <summary>
    ///     Puts the values in their own column on the hidden sheet and returns the defined name covering
    ///     them. A validation reads the name rather than the range itself, because the .xls format has no
    ///     way to point a validation at another sheet directly.
    /// </summary>
    private static string WriteToListSheet(IWorkbook workbook, string[] values)
    {
        var sheet = ListSheet(workbook);

        // the sheet grows to the right, one column per list, so lists written earlier keep their range
        var column = 0;
        var header = sheet.GetRow(0);
        if (header != null) column = header.LastCellNum < 0 ? 0 : header.LastCellNum;

        for (var i = 0; i < values.Length; i++)
        {
            var row = sheet.GetRow(i) ?? sheet.CreateRow(i);
            row.CreateCell(column).SetCellValue(values[i]);
        }

        var letter = CellReference.ConvertNumToColString(column);
        // the sheet name needs no quoting, which matters: xlsx keeps the formula as written while xls
        // reparses it
        var range = $"{sheet.SheetName}!${letter}$1:${letter}${values.Length}";

        var name = workbook.CreateName();
        name.NameName = FreeName(workbook);
        name.RefersToFormula = range;
        return name.NameName;
    }

    private static string FreeName(IWorkbook workbook)
    {
        for (var i = 1;; i++)
        {
            var name = ListNamePrefix + i;
            if (workbook.GetName(name) == null) return name;
        }
    }

    /// <summary>
    ///     The sheet the long lists are written to: the one an earlier export of this workbook left behind,
    ///     or a new hidden one. A visible sheet of that name belongs to whoever made the workbook, so a name
    ///     that is still free is taken instead of writing into it.
    /// </summary>
    private static ISheet ListSheet(IWorkbook workbook)
    {
        var name = ListSheetName;
        for (var suffix = 2;; suffix++)
        {
            var existing = workbook.GetSheet(name);
            if (existing == null)
            {
                var sheet = workbook.CreateSheet(name);
                workbook.SetSheetHidden(workbook.GetSheetIndex(sheet), SheetVisibility.Hidden);
                return sheet;
            }

            if (workbook.GetSheetVisibility(workbook.GetSheetIndex(existing)) != SheetVisibility.Visible)
                return existing;

            name = ListSheetName + suffix;
        }
    }
}
