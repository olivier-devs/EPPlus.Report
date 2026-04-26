namespace EPPlus.Report.Evaluation
{
    public interface IExpressionEvaluator
    {
        object Evaluate(string expression, object context);
    }
}
