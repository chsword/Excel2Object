namespace Chsword.Excel2Object;

/// <summary>
/// Represents errors that occur during Excel to object conversion.
/// </summary>
public class Excel2ObjectException : Exception
{
    /// <summary>
    /// Initializes a new instance of the <see cref="Excel2ObjectException"/> class with a specified error message.
    /// </summary>
    /// <param name="message">The message that describes the error.</param>
    public Excel2ObjectException(string message) : base(message)
    {
    }
    
    /// <summary>
    /// Initializes a new instance of the <see cref="Excel2ObjectException"/> class with a specified error message and inner exception.
    /// </summary>
    /// <param name="message">The message that describes the error.</param>
    /// <param name="innerException">The exception that is the cause of the current exception.</param>
    public Excel2ObjectException(string message, Exception innerException) : base(message, innerException)
    {
    }
}