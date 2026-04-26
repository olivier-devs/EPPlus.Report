namespace EPPlus.Report.Model;

/// <summary>
///     Represents a template node containing an expression such as a property path or function call.
/// </summary>
public class ExpressionNode : TemplateNode
{
    /// <summary>
    ///     Gets or sets the property path or variable name to evaluate.
    /// </summary>
    public string ExpressionPath { get; set; } = string.Empty;

    /// <summary>
    ///     Gets or sets the optional function name to apply to the evaluated result.
    /// </summary>
    public string FunctionName { get; set; } = string.Empty;
}