namespace EPPlus.Report.Evaluation;

/// <summary>
///     Defines methods for evaluating template expressions against a data context.
/// </summary>
public interface IExpressionEvaluator
{
    /// <summary>
    ///     Evaluates the specified expression against the provided context object.
    /// </summary>
    /// <param name="expression">The expression to evaluate, such as a property path.</param>
    /// <param name="context">The object against which the expression is evaluated.</param>
    /// <returns>The result of the evaluation.</returns>
    object Evaluate(string expression, object context);
}