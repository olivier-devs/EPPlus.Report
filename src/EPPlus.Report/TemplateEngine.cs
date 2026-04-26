using System;
using System.Collections.Generic;
using System.IO;
using EPPlus.Report.Evaluation;
using EPPlus.Report.Model;
using EPPlus.Report.Parsing;
using EPPlus.Report.Rendering;
using OfficeOpenXml;

namespace EPPlus.Report;

/// <summary>
///     Provides the main API for parsing Excel templates, injecting data, and generating rendered Excel files.
/// </summary>
public class TemplateEngine
{
    private readonly ExcelPackage _package;
    private readonly ITemplateParser _parser;
    private readonly Dictionary<string, Func<object, object>> _registeredFunctions = new();
    private readonly string _templatePath;
    private readonly Dictionary<string, object> _variables = new();
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
}