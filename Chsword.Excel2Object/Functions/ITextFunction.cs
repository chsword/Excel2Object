namespace Chsword.Excel2Object.Functions;

/// <summary>
/// Interface for text functions used in Excel to object conversion.
/// </summary>
public interface ITextFunction
{
    /// <summary>
    /// Converts a string to ASCII values.
    /// </summary>
    /// <param name="str">The string to convert.</param>
    /// <returns>A <see cref="ColumnValue"/> representing the ASCII values of the string.</returns>
    ColumnValue Asc(ColumnValue str);

    /// <summary>
    /// Finds one text string within another, starting at a specified position.
    /// </summary>
    /// <param name="findText">The text to find.</param>
    /// <param name="withinText">The text to search within.</param>
    /// <param name="startNum">The position to start the search from.</param>
    /// <returns>A <see cref="ColumnValue"/> representing the position of the found text.</returns>
    ColumnValue Find(ColumnValue findText, ColumnValue withinText, ColumnValue startNum);

    /// <summary>
    /// Finds one text string within another.
    /// </summary>
    /// <param name="findText">The text to find.</param>
    /// <param name="withinText">The text to search within.</param>
    /// <returns>A <see cref="ColumnValue"/> representing the position of the found text.</returns>
    ColumnValue Find(ColumnValue findText, ColumnValue withinText);

    /// <summary>
    /// Searches for one text string within another (case-insensitive).
    /// </summary>
    ColumnValue Search(ColumnValue findText, ColumnValue withinText, ColumnValue startNum);

    /// <summary>
    /// Searches for one text string within another (case-insensitive).
    /// </summary>
    ColumnValue Search(ColumnValue findText, ColumnValue withinText);

    /// <summary>
    /// Returns the leftmost characters from a text value.
    /// </summary>
    ColumnValue Left(ColumnValue text, ColumnValue numChars);

    /// <summary>
    /// Returns the rightmost characters from a text value.
    /// </summary>
    ColumnValue Right(ColumnValue text, ColumnValue numChars);

    /// <summary>
    /// Returns a specific number of characters from a text string starting at the position you specify.
    /// </summary>
    ColumnValue Mid(ColumnValue text, ColumnValue startNum, ColumnValue numChars);

    /// <summary>
    /// Returns the number of characters in a text string.
    /// </summary>
    ColumnValue Len(ColumnValue text);

    /// <summary>
    /// Converts text to uppercase.
    /// </summary>
    ColumnValue Upper(ColumnValue text);

    /// <summary>
    /// Converts text to lowercase.
    /// </summary>
    ColumnValue Lower(ColumnValue text);

    /// <summary>
    /// Capitalizes the first letter of each word in a text string.
    /// </summary>
    ColumnValue Proper(ColumnValue text);

    /// <summary>
    /// Removes spaces from text.
    /// </summary>
    ColumnValue Trim(ColumnValue text);

    /// <summary>
    /// Substitutes new text for old text in a text string.
    /// </summary>
    ColumnValue Substitute(ColumnValue text, ColumnValue oldText, ColumnValue newText, ColumnValue instanceNum);

    /// <summary>
    /// Substitutes new text for old text in a text string.
    /// </summary>
    ColumnValue Substitute(ColumnValue text, ColumnValue oldText, ColumnValue newText);

    /// <summary>
    /// Replaces part of a text string with a different text string.
    /// </summary>
    ColumnValue Replace(ColumnValue oldText, ColumnValue startNum, ColumnValue numChars, ColumnValue newText);

    /// <summary>
    /// Repeats text a given number of times.
    /// </summary>
    ColumnValue Rept(ColumnValue text, ColumnValue numberTimes);

    /// <summary>
    /// Joins several text strings into one text string.
    /// </summary>
    ColumnValue Concatenate(params ColumnValue[] texts);

    /// <summary>
    /// Converts a value to text in a specific number format.
    /// </summary>
    ColumnValue Text(ColumnValue value, ColumnValue formatText);

    /// <summary>
    /// Converts a text string that represents a number to a number.
    /// </summary>
    ColumnValue Value(ColumnValue text);

    /// <summary>
    /// Returns the numeric code for the first character in a text string.
    /// </summary>
    ColumnValue Code(ColumnValue text);

    /// <summary>
    /// Returns the character specified by the code number.
    /// </summary>
    ColumnValue Char(ColumnValue number);

    /// <summary>
    /// Checks to see if two text values are identical.
    /// </summary>
    ColumnValue Exact(ColumnValue text1, ColumnValue text2);

    /// <summary>
    /// Formats a number as text with a fixed number of decimals.
    /// </summary>
    ColumnValue Fixed(ColumnValue number, ColumnValue decimals, ColumnValue noCommas);

    /// <summary>
    /// Formats a number as text with a fixed number of decimals.
    /// </summary>
    ColumnValue Fixed(ColumnValue number, ColumnValue decimals);
}