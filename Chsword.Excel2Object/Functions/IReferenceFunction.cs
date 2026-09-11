namespace Chsword.Excel2Object.Functions;

/// <summary>
///     Excel lookup and reference functions. See <see cref="IExcelFunction" /> for how these are translated.
/// </summary>
public interface IReferenceFunction : IExcelFunction
{
    // ---- lookups ----

    /// <summary>VLOOKUP - looks a value up in the first column of a range.</summary>
    [ExcelFunctionName("VLOOKUP")]
    ColumnValue VLookup(ColumnValue val, ColumnMatrix tableArray, ColumnValue colIndexNum,
        bool rangeLookup = false);

    /// <summary>HLOOKUP - looks a value up in the first row of a range.</summary>
    [ExcelFunctionName("HLOOKUP")]
    ColumnValue HLookup(ColumnValue val, ColumnMatrix tableArray, ColumnValue rowIndexNum,
        bool rangeLookup = false);

    /// <summary>XLOOKUP - looks a value up in one range and returns the match from another.</summary>
    [ExcelFunctionName("XLOOKUP", Future = true)]
    ColumnValue XLookup(ColumnValue val, ColumnMatrix lookupArray, ColumnMatrix returnArray);

    /// <summary>XLOOKUP - as above, with a value to return when nothing matches.</summary>
    [ExcelFunctionName("XLOOKUP", Future = true)]
    ColumnValue XLookup(ColumnValue val, ColumnMatrix lookupArray, ColumnMatrix returnArray,
        ColumnValue ifNotFound);

    /// <summary>XLOOKUP - as above, with a match mode and a search mode.</summary>
    [ExcelFunctionName("XLOOKUP", Future = true)]
    ColumnValue XLookup(ColumnValue val, ColumnMatrix lookupArray, ColumnMatrix returnArray,
        ColumnValue ifNotFound, ColumnValue matchMode, ColumnValue searchMode);

    /// <summary>LOOKUP - looks a value up in a vector and returns the matching item of another.</summary>
    ColumnValue Lookup(ColumnValue val, ColumnMatrix lookupVector, ColumnMatrix resultVector);

    /// <summary>MATCH - position of a value within a range.</summary>
    ColumnValue Match(ColumnValue val, ColumnMatrix tableArray, int matchType);

    /// <summary>
    ///     MATCH - position of a value within a range, using Excel's default match type: the largest
    ///     value that is less than or equal to it, in a range sorted ascending. Pass a match type of
    ///     0 for an exact match.
    /// </summary>
    ColumnValue Match(ColumnValue val, ColumnMatrix tableArray);

    /// <summary>XMATCH - position of a value within a range, with a match mode.</summary>
    [ExcelFunctionName("XMATCH", Future = true)]
    ColumnValue XMatch(ColumnValue val, ColumnMatrix lookupArray);

    /// <summary>XMATCH - position of a value within a range, with match and search modes.</summary>
    [ExcelFunctionName("XMATCH", Future = true)]
    ColumnValue XMatch(ColumnValue val, ColumnMatrix lookupArray, ColumnValue matchMode, ColumnValue searchMode);

    /// <summary>INDEX - the cell of a range at the given row and column.</summary>
    ColumnValue Index(ColumnMatrix array, ColumnValue rowNum, ColumnValue columnNum);

    /// <summary>INDEX - the cell of a one dimensional range at the given position.</summary>
    ColumnValue Index(ColumnMatrix array, ColumnValue rowNum);

    /// <summary>CHOOSE - picks a value from a list by its position.</summary>
    ColumnValue Choose(ColumnValue indexNum, params ColumnValue[] values);

    // ---- addressing ----

    /// <summary>OFFSET - a range shifted from a starting reference.</summary>
    ColumnValue Offset(ColumnValue reference, ColumnValue rows, ColumnValue cols);

    /// <summary>OFFSET - a range of the given size, shifted from a starting reference.</summary>
    ColumnValue Offset(ColumnValue reference, ColumnValue rows, ColumnValue cols, ColumnValue height,
        ColumnValue width);

    /// <summary>INDIRECT - the reference named by a text value.</summary>
    ColumnValue Indirect(ColumnValue refText);

    /// <summary>ADDRESS - the address of a cell, as text.</summary>
    ColumnValue Address(ColumnValue rowNum, ColumnValue columnNum);

    /// <summary>ADDRESS - the address of a cell, as text, with an absolute/relative mode.</summary>
    ColumnValue Address(ColumnValue rowNum, ColumnValue columnNum, ColumnValue absNum);

    /// <summary>ROW - the row number of the formula cell.</summary>
    ColumnValue Row();

    /// <summary>ROW - the row number of a reference.</summary>
    ColumnValue Row(ColumnValue reference);

    /// <summary>ROWS - the number of rows of a range.</summary>
    ColumnValue Rows(ColumnMatrix array);

    /// <summary>COLUMN - the column number of the formula cell.</summary>
    ColumnValue Column();

    /// <summary>COLUMN - the column number of a reference.</summary>
    ColumnValue Column(ColumnValue reference);

    /// <summary>COLUMNS - the number of columns of a range.</summary>
    ColumnValue Columns(ColumnMatrix array);

    /// <summary>AREAS - the number of areas of a reference.</summary>
    ColumnValue Areas(ColumnValue reference);

    /// <summary>FORMULATEXT - the formula of a cell, as text.</summary>
    [ExcelFunctionName("FORMULATEXT", Future = true)]
    ColumnValue FormulaText(ColumnValue reference);

    /// <summary>HYPERLINK - a link to a document or a web address.</summary>
    ColumnValue Hyperlink(ColumnValue linkLocation);

    /// <summary>HYPERLINK - a link with the text to display.</summary>
    ColumnValue Hyperlink(ColumnValue linkLocation, ColumnValue friendlyName);

    // ---- dynamic arrays ----

    /// <summary>TRANSPOSE - flips a range from rows to columns.</summary>
    ColumnValue Transpose(ColumnMatrix array);

    /// <summary>UNIQUE - the distinct values of a range.</summary>
    [ExcelFunctionName("UNIQUE", Future = true)]
    ColumnValue Unique(ColumnMatrix array);

    /// <summary>SORT - a range sorted by one of its columns.</summary>
    [ExcelFunctionName("SORT", Future = true, WorksheetOnly = true)]
    ColumnValue Sort(ColumnMatrix array);

    /// <summary>SORT - a range sorted by the given index and order.</summary>
    [ExcelFunctionName("SORT", Future = true, WorksheetOnly = true)]
    ColumnValue Sort(ColumnMatrix array, ColumnValue sortIndex, ColumnValue sortOrder);

    /// <summary>SORTBY - a range sorted by the values of another range.</summary>
    [ExcelFunctionName("SORTBY", Future = true)]
    ColumnValue SortBy(ColumnMatrix array, ColumnMatrix byArray);

    /// <summary>FILTER - the rows of a range that meet a condition.</summary>
    [ExcelFunctionName("FILTER", Future = true, WorksheetOnly = true)]
    ColumnValue Filter(ColumnMatrix array, ColumnValue include);

    /// <summary>FILTER - the rows of a range that meet a condition, or a fallback when none do.</summary>
    [ExcelFunctionName("FILTER", Future = true, WorksheetOnly = true)]
    ColumnValue Filter(ColumnMatrix array, ColumnValue include, ColumnValue ifEmpty);

    /// <summary>SEQUENCE - a sequence of numbers.</summary>
    [ExcelFunctionName("SEQUENCE", Future = true)]
    ColumnValue Sequence(ColumnValue rows);

    /// <summary>SEQUENCE - a sequence of numbers laid out in rows and columns.</summary>
    [ExcelFunctionName("SEQUENCE", Future = true)]
    ColumnValue Sequence(ColumnValue rows, ColumnValue columns, ColumnValue start, ColumnValue step);

    /// <summary>TAKE - the first or last rows of a range.</summary>
    [ExcelFunctionName("TAKE", Future = true)]
    ColumnValue Take(ColumnMatrix array, ColumnValue rows);

    /// <summary>DROP - a range without its first or last rows.</summary>
    [ExcelFunctionName("DROP", Future = true)]
    ColumnValue Drop(ColumnMatrix array, ColumnValue rows);

    /// <summary>CHOOSECOLS - the named columns of a range.</summary>
    [ExcelFunctionName("CHOOSECOLS", Future = true)]
    ColumnValue ChooseCols(ColumnMatrix array, params ColumnValue[] colNums);

    /// <summary>CHOOSEROWS - the named rows of a range.</summary>
    [ExcelFunctionName("CHOOSEROWS", Future = true)]
    ColumnValue ChooseRows(ColumnMatrix array, params ColumnValue[] rowNums);

    /// <summary>VSTACK - stacks ranges vertically.</summary>
    [ExcelFunctionName("VSTACK", Future = true)]
    ColumnValue VStack(params ColumnValue[] arrays);

    /// <summary>HSTACK - stacks ranges horizontally.</summary>
    [ExcelFunctionName("HSTACK", Future = true)]
    ColumnValue HStack(params ColumnValue[] arrays);

    /// <summary>TOCOL - a range flattened into one column.</summary>
    [ExcelFunctionName("TOCOL", Future = true)]
    ColumnValue ToCol(ColumnMatrix array);

    /// <summary>TOROW - a range flattened into one row.</summary>
    [ExcelFunctionName("TOROW", Future = true)]
    ColumnValue ToRow(ColumnMatrix array);

    /// <summary>EXPAND - a range padded to the given size.</summary>
    [ExcelFunctionName("EXPAND", Future = true)]
    ColumnValue Expand(ColumnMatrix array, ColumnValue rows, ColumnValue columns, ColumnValue padWith);
}
