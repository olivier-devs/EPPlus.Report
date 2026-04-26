using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Reflection;

namespace EPPlus.Report.Evaluation
{
    public class ExpressionEvaluator : IExpressionEvaluator
    {
        private readonly ConcurrentDictionary<string, PropertyInfo[]> _cache = new ConcurrentDictionary<string, PropertyInfo[]>();
        private readonly ConcurrentDictionary<string, Func<object, object>> _functions = new ConcurrentDictionary<string, Func<object, object>>();

        public ExpressionEvaluator()
        {
            _functions["Upper"] = x => x?.ToString()?.ToUpperInvariant();
            _functions["Lower"] = x => x?.ToString()?.ToLowerInvariant();
            _functions["Trim"] = x => x?.ToString()?.Trim();
        }

        public void RegisterFunction(string name, Func<object, object> func)
        {
            if (string.IsNullOrWhiteSpace(name))
                throw new ArgumentException("Function name cannot be empty", nameof(name));
            if (func == null)
                throw new ArgumentNullException(nameof(func));

            _functions[name] = func;
        }

        public object ApplyFunction(string functionName, object value)
        {
            if (string.IsNullOrWhiteSpace(functionName))
                throw new ArgumentException("Function name cannot be empty", nameof(functionName));

            if (!_functions.TryGetValue(functionName, out var func))
                throw new ArgumentException($"Function '{functionName}' is not registered", nameof(functionName));

            return func(value);
        }

        public object Evaluate(string expression, object context)
        {
            return Evaluate(expression, context, string.Empty);
        }

        public object Evaluate(string expression, object context, string functionName)
        {
            if (context == null)
                throw new ArgumentNullException(nameof(context));
            if (string.IsNullOrWhiteSpace(expression))
                throw new ArgumentException("Expression cannot be empty", nameof(expression));

            var cacheKey = $"{context.GetType().FullName}:{expression}";
            if (!_cache.TryGetValue(cacheKey, out var properties))
            {
                properties = CompileExpression(expression, context.GetType());
                _cache[cacheKey] = properties;
            }

            object current = context;
            foreach (var property in properties)
            {
                if (current == null)
                    return null;
                current = property.GetValue(current);
            }

            if (!string.IsNullOrEmpty(functionName))
            {
                if (!_functions.TryGetValue(functionName, out var func))
                    throw new ArgumentException($"Function '{functionName}' is not registered", nameof(functionName));
                current = func(current);
            }

            return current;
        }

        private PropertyInfo[] CompileExpression(string expression, Type contextType)
        {
            var parts = expression.Split('.');
            var properties = new List<PropertyInfo>();
            Type currentType = contextType;

            foreach (var part in parts)
            {
                var trimmedPart = part.Trim();
                var property = currentType.GetProperty(trimmedPart, BindingFlags.Public | BindingFlags.Instance);
                if (property == null)
                    throw new PropertyNotFoundException($"Property '{trimmedPart}' not found on type '{currentType.Name}'");
                
                properties.Add(property);
                currentType = property.PropertyType;
            }

            return properties.ToArray();
        }
    }
}
