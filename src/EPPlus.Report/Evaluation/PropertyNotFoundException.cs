using System;

namespace EPPlus.Report.Evaluation;

/// <summary>
///     Exception thrown when a property referenced in a template expression cannot be found on the target type.
/// </summary>
public class PropertyNotFoundException : ArgumentException
{
    /// <summary>
    ///     Initializes a new instance of the <see cref="PropertyNotFoundException" /> class with the specified error message.
    /// </summary>
    /// <param name="message">The message that describes the error.</param>
    public PropertyNotFoundException(string message) : base(message)
    {
    }
}