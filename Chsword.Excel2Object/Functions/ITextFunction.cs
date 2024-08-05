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
    ColumnValue Find(ColumnValue findText, ColumnValue withinText, int startNum);

    /// <summary>
    /// Finds one text string within another.
    /// </summary>
    /// <param name="findText">The text to find.</param>
    /// <param name="withinText">The text to search within.</param>
    /// <returns>A <see cref="ColumnValue"/> representing the position of the found text.</returns>
    ColumnValue Find(ColumnValue findText, ColumnValue withinText);
}