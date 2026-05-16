using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using EPPlus.Report.Evaluation;
using EPPlus.Report.Model;
using EPPlus.Report.Parsing;
using EPPlus.Report.Rendering;
using OfficeOpenXml;

namespace EPPlus.Report;

/// <summary>
///     Provides the main API for parsing Excel templates, injecting data, and generating rendered Excel files.
/// </summary>
public class TemplateEngine : IDisposable
{
    private readonly ExcelPackage _package;
    private readonly ITemplateParser _parser;
    private readonly Dictionary<string, Func<object, object>> _registeredFunctions = new();
    private readonly string _templatePath;
    private readonly Dictionary<string, object> _variables = new();
    private readonly HashSet<string> _allowedProperties = new();
    private object _rootValue;

    /// <summary>
    ///     Initializes a new instance of the <see cref="TemplateEngine" /> class from a file path.
    /// </summary>
    /// <param name="templatePath">The path to the Excel template file.</param>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="templatePath" /> is null.</exception>
    public TemplateEngine(string templatePath)
    {
        _templatePath = templatePath ?? throw new ArgumentNullException(nameof(templatePath));
        _package = new ExcelPackage(new FileInfo(templatePath));
        _parser = new TemplateParser();
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="TemplateEngine" /> class from a stream.
    /// </summary>
    /// <param name="stream">The stream containing the Excel template.</param>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="stream" /> is null.</exception>
    public TemplateEngine(Stream stream)
    {
        if (stream == null)
        {
            throw new ArgumentNullException(nameof(stream));
        }

        _package = new ExcelPackage(stream);
        _parser = new TemplateParser();
    }

    /// <summary>
    ///     Adds a root variable that will be used as the default data context during rendering.
    /// </summary>
    /// <param name="value">The data object to inject into the template.</param>
    public void AddVariable(object value)
    {
        _rootValue = value;
    }

    /// <summary>
    ///     Adds a named variable that can be referenced in template expressions.
    /// </summary>
    /// <param name="name">The name of the variable.</param>
    /// <param name="value">The value of the variable.</param>
    /// <exception cref="ArgumentException">Thrown when <paramref name="name" /> is null or whitespace.</exception>
    public void AddVariable(string name, object value)
    {
        if (string.IsNullOrWhiteSpace(name))
        {
            throw new ArgumentException("Variable name cannot be empty", nameof(name));
        }

        _variables[name] = value;
    }

    /// <summary>
    ///     Registers a custom function that can be used in template expressions.
    /// </summary>
    /// <remarks>
    ///     <para><b>Security notice:</b> this method allows arbitrary code execution via the supplied delegate.
    ///     Only register functions from trusted sources. Never register user-supplied functions.</para>
    /// </remarks>
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

        _registeredFunctions[name] = func ?? throw new ArgumentNullException(nameof(func));
    }

    /// <summary>
    ///     Adds a property expression path to the allowed properties list.
    ///     When an allowlist is active, only expressions in the allowlist will be evaluated.
    /// </summary>
    /// <remarks>
    ///     <para>Properties added via this method are merged with properties discovered from
    ///     <see cref="TemplateVisibleAttribute" /> decorations. The resulting allowlist is the
    ///     union of both sources.</para>
    ///     <para>If no properties are added and no <see cref="TemplateVisibleAttribute" /> is used,
    ///     all public properties remain accessible (default behavior).</para>
    /// </remarks>
    /// <param name="propertyPath">
    ///     The property expression path to allow (e.g., "Name", "Address.City").
    ///     The path is trimmed of whitespace before being added.
    /// </param>
    /// <exception cref="ArgumentException">Thrown when <paramref name="propertyPath" /> is null or whitespace.</exception>
    public void AllowProperty(string propertyPath)
    {
        if (string.IsNullOrWhiteSpace(propertyPath))
        {
            throw new ArgumentException("Property path cannot be empty", nameof(propertyPath));
        }

        _allowedProperties.Add(propertyPath.Trim());
    }

    /// <summary>
    ///     Parses the template, evaluates expressions, and renders the result into the workbook.
    /// </summary>
    /// <param name="options">Optional generation options, such as formula evaluation.</param>
    /// <returns>A <see cref="TemplateGenerateResult" /> containing any parsing, rendering, or warning information.</returns>
    public TemplateGenerateResult Generate(GenerateOptions options = null)
    {
        var parsingErrors = new TemplateErrors();
        var renderingErrors = new TemplateErrors();
        var warnings = new TemplateErrors();
        var tracker = new RowOperationTracker();
        var evaluator = new ExpressionEvaluator();
        foreach (var kvp in _registeredFunctions)
        {
            evaluator.RegisterFunction(kvp.Key, kvp.Value);
        }

        // Collect TemplateVisible properties from root value and variables
        var visibleProperties = CollectTemplateVisibleProperties();

        // Merge: AllowProperty() entries + [TemplateVisible] discovered entries
        evaluator.AllowedProperties = new HashSet<string>(_allowedProperties);
        foreach (var property in visibleProperties)
        {
            evaluator.AllowedProperties.Add(property);
        }

        // If no allowlist entries exist (neither from AllowProperty nor from [TemplateVisible]),
        // keep AllowedProperties null to allow all properties (default behavior)
        if (evaluator.AllowedProperties.Count == 0)
        {
            evaluator.AllowedProperties = null;
        }

        var renderer = new TemplateRenderer(evaluator, renderingErrors, tracker, warnings);

        foreach (var worksheet in _package.Workbook.Worksheets)
        {
            var template = _parser.Parse(worksheet, parsingErrors);
            var context = new RenderContext
            {
                Current = _rootValue,
                Variables = new Dictionary<string, object>(_variables)
            };
            renderer.Render(template, context, worksheet);
        }

        var adjuster = new ExcelTableAdjuster(tracker);
        adjuster.AdjustAll(_package);

        if (options != null && options.EvaluateFormulas)
        {
            _package.Workbook.Calculate();
        }

        return new TemplateGenerateResult(parsingErrors, renderingErrors, warnings);
    }

    /// <summary>
    ///     Saves the rendered workbook back to the original template file path.
    /// </summary>
    /// <exception cref="InvalidOperationException">Thrown when the engine was created from a stream.</exception>
    public void Save()
    {
        if (string.IsNullOrEmpty(_templatePath))
        {
            throw new InvalidOperationException(
                "Cannot save when engine was created from a stream. Use SaveAs instead.");
        }

        _package.SaveAs(new FileInfo(_templatePath));
    }

    /// <summary>
    ///     Saves the rendered workbook to the specified file path.
    /// </summary>
    /// <param name="path">The destination file path.</param>
    public void SaveAs(string path)
    {
        _package.SaveAs(new FileInfo(path));
    }

    /// <summary>
    ///     Saves the rendered workbook to the specified file.
    /// </summary>
    /// <param name="fileInfo">The destination file information.</param>
    public void SaveAs(FileInfo fileInfo)
    {
        _package.SaveAs(fileInfo);
    }

    /// <summary>
    ///     Saves the rendered workbook to the specified stream.
    /// </summary>
    /// <param name="stream">The destination stream.</param>
    public void SaveAs(Stream stream)
    {
        _package.SaveAs(stream);
    }

    /// <summary>
    ///     Saves the rendered workbook to the specified file path with save options.
    /// </summary>
    /// <param name="path">The destination file path.</param>
    /// <param name="saveOptions">Options controlling the save behavior.</param>
    public void SaveAs(string path, SaveOptions saveOptions)
    {
        ApplySaveOptions(saveOptions);
        _package.SaveAs(new FileInfo(path));
    }

    /// <summary>
    ///     Saves the rendered workbook to the specified file with save options.
    /// </summary>
    /// <param name="fileInfo">The destination file information.</param>
    /// <param name="saveOptions">Options controlling the save behavior.</param>
    public void SaveAs(FileInfo fileInfo, SaveOptions saveOptions)
    {
        ApplySaveOptions(saveOptions);
        _package.SaveAs(fileInfo);
    }

    /// <summary>
    ///     Saves the rendered workbook to the specified stream with save options.
    /// </summary>
    /// <param name="stream">The destination stream.</param>
    /// <param name="saveOptions">Options controlling the save behavior.</param>
    public void SaveAs(Stream stream, SaveOptions saveOptions)
    {
        ApplySaveOptions(saveOptions);
        _package.SaveAs(stream);
    }

    private void ApplySaveOptions(SaveOptions options)
    {
        if (options != null && options.EvaluateFormulasBeforeSave)
        {
            _package.Workbook.Calculate();
        }
    }

    /// <summary>
    ///     Collects all property paths marked with <see cref="TemplateVisibleAttribute" /> from the root value and variables.
    /// </summary>
    /// <returns>A set of property expression paths that are visible in templates.</returns>
    private HashSet<string> CollectTemplateVisibleProperties()
    {
        var result = new HashSet<string>();

        // Collect from root value (no prefix)
        if (_rootValue != null)
        {
            var visitedTypes = new HashSet<Type>();
            CollectVisiblePropertiesFromType(_rootValue.GetType(), "", result, visitedTypes);
        }

        // Collect from named variables (with variable name as prefix)
        foreach (var kvp in _variables)
        {
            var variableName = kvp.Key;
            var variableValue = kvp.Value;
            if (variableValue != null)
            {
                var visitedTypes = new HashSet<Type>();
                // Add properties with the variable name as prefix (e.g., "invoice.InvoiceNumber")
                CollectVisiblePropertiesFromType(variableValue.GetType(), variableName, result, visitedTypes);
                // Also add properties without prefix for direct access in loops
                visitedTypes.Clear();
                CollectVisiblePropertiesFromType(variableValue.GetType(), "", result, visitedTypes);
            }
        }

        return result;
    }

    /// <summary>
    ///     Recursively collects visible property paths from a type and its nested properties.
    /// </summary>
    private void CollectVisiblePropertiesFromType(Type type, string prefix, HashSet<string> result, HashSet<Type> visitedTypes)
    {
        // Prevent infinite recursion with circular references
        if (visitedTypes.Contains(type))
        {
            return;
        }

        // Skip primitive types and strings
        if (type.IsPrimitive || type == typeof(string) || type == typeof(decimal) || type == typeof(DateTime))
        {
            return;
        }

        visitedTypes.Add(type);

        // Get all properties including inherited ones
        var properties = type.GetProperties(BindingFlags.Public | BindingFlags.Instance);

        foreach (var property in properties)
        {
            // Check if property has TemplateVisibleAttribute
            var hasVisibleAttribute = property.GetCustomAttribute<TemplateVisibleAttribute>() != null;

            if (hasVisibleAttribute)
            {
                var propertyPath = string.IsNullOrEmpty(prefix) ? property.Name : $"{prefix}.{property.Name}";
                result.Add(propertyPath);

                // Recursively collect from nested types
                var propertyType = property.PropertyType;

                // Handle nullable types
                if (propertyType.IsGenericType && propertyType.GetGenericTypeDefinition() == typeof(Nullable<>))
                {
                    propertyType = propertyType.GetGenericArguments()[0];
                }

                // Handle collections - get the element type
                if (propertyType.IsGenericType && typeof(System.Collections.IEnumerable).IsAssignableFrom(propertyType))
                {
                    var elementType = propertyType.GetGenericArguments().FirstOrDefault();
                    if (elementType != null && !elementType.IsPrimitive && elementType != typeof(string))
                    {
                        CollectVisiblePropertiesFromType(elementType, "", result, visitedTypes);
                    }
                }
                else if (!propertyType.IsPrimitive && propertyType != typeof(string))
                {
                    // Recurse into complex types
                    CollectVisiblePropertiesFromType(propertyType, propertyPath, result, visitedTypes);
                }
            }
        }
    }

    /// <summary>
    ///     Releases the underlying Excel package and all associated resources.
    /// </summary>
    public void Dispose()
    {
        _package?.Dispose();
    }
}