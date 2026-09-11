namespace Chsword.Excel2Object.Functions;

/// <summary>
///     Excel text functions. See <see cref="IExcelFunction" /> for how these are translated.
/// </summary>
public interface ITextFunction : IExcelFunction
{
    // ---- pieces of a string ----

    /// <summary>LEFT - the first character of a string.</summary>
    ColumnValue Left(ColumnValue text);

    /// <summary>LEFT - the first characters of a string.</summary>
    ColumnValue Left(ColumnValue text, ColumnValue numChars);

    /// <summary>RIGHT - the last character of a string.</summary>
    ColumnValue Right(ColumnValue text);

    /// <summary>RIGHT - the last characters of a string.</summary>
    ColumnValue Right(ColumnValue text, ColumnValue numChars);

    /// <summary>MID - the characters of a string from a position on.</summary>
    ColumnValue Mid(ColumnValue text, ColumnValue startNum, ColumnValue numChars);

    /// <summary>LEN - the number of characters of a string.</summary>
    [ExcelFunctionName("LEN")]
    ColumnValue Len(ColumnValue text);

    /// <summary>LEFTB - the first bytes of a string, double byte characters counting as two.</summary>
    [ExcelFunctionName("LEFTB")]
    ColumnValue LeftB(ColumnValue text, ColumnValue numBytes);

    /// <summary>RIGHTB - the last bytes of a string, double byte characters counting as two.</summary>
    [ExcelFunctionName("RIGHTB")]
    ColumnValue RightB(ColumnValue text, ColumnValue numBytes);

    /// <summary>MIDB - the bytes of a string from a position on.</summary>
    [ExcelFunctionName("MIDB")]
    ColumnValue MidB(ColumnValue text, ColumnValue startNum, ColumnValue numBytes);

    /// <summary>LENB - the number of bytes of a string, double byte characters counting as two.</summary>
    [ExcelFunctionName("LENB")]
    ColumnValue LenB(ColumnValue text);

    /// <summary>TEXTBEFORE - the text before a delimiter.</summary>
    [ExcelFunctionName("TEXTBEFORE", Future = true)]
    ColumnValue TextBefore(ColumnValue text, ColumnValue delimiter);

    /// <summary>TEXTAFTER - the text after a delimiter.</summary>
    [ExcelFunctionName("TEXTAFTER", Future = true)]
    ColumnValue TextAfter(ColumnValue text, ColumnValue delimiter);

    /// <summary>TEXTSPLIT - splits a string by a column delimiter.</summary>
    [ExcelFunctionName("TEXTSPLIT", Future = true)]
    ColumnValue TextSplit(ColumnValue text, ColumnValue colDelimiter);

    // ---- searching and editing ----

    /// <summary>FIND - the position of one string inside another, case sensitive.</summary>
    ColumnValue Find(ColumnValue findText, ColumnValue withinText);

    /// <summary>FIND - the position of one string inside another, from a start position.</summary>
    ColumnValue Find(ColumnValue findText, ColumnValue withinText, int startNum);

    /// <summary>FINDB - as FIND, counting double byte characters as two.</summary>
    [ExcelFunctionName("FINDB")]
    ColumnValue FindB(ColumnValue findText, ColumnValue withinText);

    /// <summary>SEARCH - the position of one string inside another, ignoring case and allowing wildcards.</summary>
    ColumnValue Search(ColumnValue findText, ColumnValue withinText);

    /// <summary>SEARCH - as above, from a start position.</summary>
    ColumnValue Search(ColumnValue findText, ColumnValue withinText, ColumnValue startNum);

    /// <summary>SEARCHB - as SEARCH, counting double byte characters as two.</summary>
    [ExcelFunctionName("SEARCHB")]
    ColumnValue SearchB(ColumnValue findText, ColumnValue withinText);

    /// <summary>SUBSTITUTE - replaces every occurrence of a string with another.</summary>
    ColumnValue Substitute(ColumnValue text, ColumnValue oldText, ColumnValue newText);

    /// <summary>SUBSTITUTE - replaces the n-th occurrence of a string with another.</summary>
    ColumnValue Substitute(ColumnValue text, ColumnValue oldText, ColumnValue newText, ColumnValue instanceNum);

    /// <summary>REPLACE - replaces part of a string, by position.</summary>
    ColumnValue Replace(ColumnValue oldText, ColumnValue startNum, ColumnValue numChars, ColumnValue newText);

    /// <summary>REPT - repeats a string the given number of times.</summary>
    [ExcelFunctionName("REPT")]
    ColumnValue Rept(ColumnValue text, ColumnValue numberTimes);

    /// <summary>TRIM - removes the extra spaces of a string.</summary>
    ColumnValue Trim(ColumnValue text);

    /// <summary>CLEAN - removes the non printable characters of a string.</summary>
    ColumnValue Clean(ColumnValue text);

    // ---- case and comparison ----

    /// <summary>UPPER - converts a string to upper case.</summary>
    ColumnValue Upper(ColumnValue text);

    /// <summary>LOWER - converts a string to lower case.</summary>
    ColumnValue Lower(ColumnValue text);

    /// <summary>PROPER - capitalizes the first letter of every word.</summary>
    ColumnValue Proper(ColumnValue text);

    /// <summary>EXACT - TRUE when two strings are exactly the same.</summary>
    ColumnValue Exact(ColumnValue text1, ColumnValue text2);

    // ---- joining ----

    /// <summary>CONCAT - joins its arguments into one string.</summary>
    [ExcelFunctionName("CONCAT", Future = true)]
    ColumnValue Concat(params ColumnValue[] values);

    /// <summary>CONCATENATE - joins its arguments into one string.</summary>
    [ExcelFunctionName("CONCATENATE")]
    ColumnValue Concatenate(params ColumnValue[] values);

    /// <summary>TEXTJOIN - joins its arguments with a delimiter.</summary>
    [ExcelFunctionName("TEXTJOIN", Future = true)]
    ColumnValue TextJoin(ColumnValue delimiter, ColumnValue ignoreEmpty, params ColumnValue[] values);

    // ---- conversion ----

    /// <summary>TEXT - formats a number with a format code, e.g. <c>"0.00"</c>.</summary>
    [ExcelFunctionName("TEXT")]
    ColumnValue Text(ColumnValue val, ColumnValue formatText);

    /// <summary>VALUE - converts a string to a number.</summary>
    ColumnValue Value(ColumnValue text);

    /// <summary>NUMBERVALUE - converts a string to a number, with explicit separators.</summary>
    [ExcelFunctionName("NUMBERVALUE", Future = true)]
    ColumnValue NumberValue(ColumnValue text, ColumnValue decimalSeparator, ColumnValue groupSeparator);

    /// <summary>FIXED - formats a number as text with a fixed number of decimals.</summary>
    ColumnValue Fixed(ColumnValue number, ColumnValue decimals);

    /// <summary>DOLLAR - formats a number as text in the currency format.</summary>
    ColumnValue Dollar(ColumnValue number, ColumnValue decimals);

    /// <summary>CHAR - the character of a character code.</summary>
    [ExcelFunctionName("CHAR")]
    ColumnValue Char(ColumnValue number);

    /// <summary>CODE - the character code of the first character of a string.</summary>
    ColumnValue Code(ColumnValue text);

    /// <summary>UNICHAR - the character of a Unicode code point.</summary>
    [ExcelFunctionName("UNICHAR", Future = true)]
    ColumnValue UniChar(ColumnValue number);

    /// <summary>UNICODE - the Unicode code point of the first character of a string.</summary>
    [ExcelFunctionName("UNICODE", Future = true)]
    ColumnValue Unicode(ColumnValue text);

    /// <summary>ASC - converts full width characters to half width ones.</summary>
    ColumnValue Asc(ColumnValue str);

    /// <summary>WIDECHAR - converts half width characters to full width ones.</summary>
    [ExcelFunctionName("WIDECHAR")]
    ColumnValue WideChar(ColumnValue text);

    /// <summary>T - the argument when it is text, an empty string otherwise.</summary>
    [ExcelFunctionName("T")]
    ColumnValue T(ColumnValue val);
}
