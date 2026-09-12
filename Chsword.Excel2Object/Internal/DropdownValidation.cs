using NPOI.SS.UserModel;
using NPOI.SS.Util;

namespace Chsword.Excel2Object.Internal;

/// <summary>
///     Writes the dropdown a column declares - Excel's list validation - over the rows of that column.
/// </summary>
/// <remarks>
///     A short list is written into the validation itself. Excel only accepts 255 characters there
///     (and splits the list on commas), so a longer list, or one whose values carry a comma or a quote,
///     is written into a hidden sheet the validation then points at.
/// </remarks>
internal static class DropdownValidation
{
    /// <summary>The hidden sheet the long lists live on.</summary>
    private const string ListSheetName = "_excel2object_lists";

    /// <summary>What Excel accepts inside a list validation, commas and quotes included.</summary>
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

    private static bool FitsInline(string[] values)
    {
        var length = values.Length - 1; // the commas between them
        foreach (var value in values)
        {
            if (value.IndexOf(',') >= 0 || value.IndexOf('"') >= 0) return false;
            length += value.Length;
        }

        return length <= InlineLimit;
    }

    /// <summary>Puts the values in their own column on the hidden sheet and returns the range.</summary>
    private static string WriteToListSheet(IWorkbook workbook, string[] values)
    {
        var sheet = workbook.GetSheet(ListSheetName);
        if (sheet == null)
        {
            sheet = workbook.CreateSheet(ListSheetName);
            workbook.SetSheetHidden(workbook.GetSheetIndex(sheet), SheetVisibility.Hidden);
        }

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
        return $"{ListSheetName}!${letter}$1:${letter}${values.Length}";
    }
}
