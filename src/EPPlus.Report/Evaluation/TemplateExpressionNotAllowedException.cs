using System;

namespace EPPlus.Report.Evaluation;

/// <summary>
///     Exception thrown when a template expression references a property that is not in the allowed properties list.
/// </summary>
/// <remarks>
///     <para>This exception is thrown by the <see cref="ExpressionEvaluator" /> when an allowlist is configured
///     via <see cref="ExpressionEvaluator.AllowedProperties" /> or <see cref="TemplateVisibleAttribute" />
///     and the expression being evaluated is not present in that allowlist.</para>
///     <para>Unlike <see cref="UnauthorizedAccessException" />, this is a domain-specific exception that
///     does not expose system-level security information and can be safely caught by rendering code.</para>
/// </remarks>
public class TemplateExpressionNotAllowedException : InvalidOperationException
{
    /// <summary>
    ///     Initializes a new instance of the <see cref="TemplateExpressionNotAllowedException" /> class
    ///     with the specified error message.
    /// </summary>
    /// <param name="message">The message that describes the error.</param>
    public TemplateExpressionNotAllowedException(string message) : base(message)
    {
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="TemplateExpressionNotAllowedException" /> class
    ///     with the specified error message and inner exception.
    /// </summary>
    /// <param name="message">The message that describes the error.</param>
    /// <param name="innerException">The exception that is the cause of the current exception.</param>
    public TemplateExpressionNotAllowedException(string message, Exception innerException) : base(message, innerException)
    {
    }
}