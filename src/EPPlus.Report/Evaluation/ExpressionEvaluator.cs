using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Reflection;

namespace EPPlus.Report.Evaluation;

/// <summary>
///     Evaluates template expressions using reflection and caches compiled property paths for performance.
/// </summary>
public class ExpressionEvaluator : IExpressionEvaluator
{
    private readonly ConcurrentDictionary<string, PropertyInfo[]> _cache = new();
    private readonly ConcurrentDictionary<string, Func<object, object>> _functions = new();

    /// <summary>
    ///     Initializes a new instance of the <see cref="ExpressionEvaluator" /> class
    ///     with built-in functions (Upper, Lower, Trim).
    /// </summary>
    public ExpressionEvaluator()
    {
        _functions["Upper"] = x => x?.ToString()?.ToUpperInvariant();
        _functions["Lower"] = x => x?.ToString()?.ToLowerInvariant();
        _functions["Trim"] = x => x?.ToString()?.Trim();
    }

    /// <summary>
    ///     Evaluates the specified expression against the provided context object.
    /// </summary>
    /// <param name="expression">The expression to evaluate, such as a property path.</param>
    /// <param name="context">The object against which the expression is evaluated.</param>
    /// <returns>The result of the evaluation.</returns>
    public object Evaluate(string expression, object context)
    {
        return Evaluate(expression, context, string.Empty);
    }

    /// <summary>
    ///     Registers a custom function that can be applied to expression results.
    /// </summary>
    /// <param name="name">The name of the function.</param>
    /// <param name="func">The function implementation.</param>
    /// <exception cref="ArgumentException">Thrown when <paramref name="name" /> is null or whitespace.</exception>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="func" /> is null.</exception>
    public void RegisterFunction(string name, Func<object, object> func)
    {
        if (string.IsNullOrWhiteSpace(name))
        {
            throw new ArgumentException("Function name cannot be empty", nameof(name));
        }

        _functions[name] = func ?? throw new ArgumentNullException(nameof(func));
    }

    /// <summary>
    ///     Applies the named function to the specified value.
    /// </summary>
    /// <param name="functionName">The name of the registered function.</param>
    /// <param name="value">The value to transform.</param>
    /// <returns>The transformed value.</returns>
    /// <exception cref="ArgumentException">Thrown when the function is not registered.</exception>
    public object ApplyFunction(string functionName, object value)
    {
        if (string.IsNullOrWhiteSpace(functionName))
        {
            throw new ArgumentException("Function name cannot be empty", nameof(functionName));
        }

        return !_functions.TryGetValue(functionName, out var func) 
            ? throw new ArgumentException($"Function '{functionName}' is not registered", nameof(functionName)) 
            : func(value);
    }

    /// <summary>
    ///     Evaluates the specified expression against the provided context object and optionally applies a function.
    /// </summary>
    /// <param name="expression">The expression to evaluate, such as a property path.</param>
    /// <param name="context">The object against which the expression is evaluated.</param>
    /// <param name="functionName">The optional function name to apply to the result.</param>
    /// <returns>The result of the evaluation.</returns>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="context" /> is null.</exception>
    /// <exception cref="ArgumentException">Thrown when <paramref name="expression" /> is null or whitespace.</exception>
    public object Evaluate(string expression, object context, string functionName)
    {
        if (context == null)
        {
            throw new ArgumentNullException(nameof(context));
        }

        if (string.IsNullOrWhiteSpace(expression))
        {
            throw new ArgumentException("Expression cannot be empty", nameof(expression));
        }

        var cacheKey = $"{context.GetType().FullName}:{expression}";
        if (!_cache.TryGetValue(cacheKey, out var properties))
        {
            properties = CompileExpression(expression, context.GetType());
            _cache[cacheKey] = properties;
        }

        var current = context;
        foreach (var property in properties)
        {
            if (current == null)
            {
                return null;
            }

            current = property.GetValue(current);
        }

        if (!string.IsNullOrEmpty(functionName))
        {
            if (!_functions.TryGetValue(functionName, out var func))
            {
                throw new ArgumentException($"Function '{functionName}' is not registered", nameof(functionName));
            }

            current = func(current);
        }

        return current;
    }

    private static PropertyInfo[] CompileExpression(string expression, Type contextType)
    {
        var parts = expression.Split('.');
        var properties = new List<PropertyInfo>();
        var currentType = contextType;

        foreach (var part in parts)
        {
            var trimmedPart = part.Trim();
            var property = currentType.GetProperty(trimmedPart, BindingFlags.Public | BindingFlags.Instance);
            if (property == null)
            {
                throw new PropertyNotFoundException($"Property '{trimmedPart}' not found on type '{currentType.Name}'");
            }

            properties.Add(property);
            currentType = property.PropertyType;
        }

        return properties.ToArray();
    }
}