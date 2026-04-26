using EPPlus.Report.Evaluation;
using EPPlus.Report.Model;
using EPPlus.Report.Parsing;
using EPPlus.Report.Rendering;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;

namespace EPPlus.Report
{
    public class TemplateEngine
    {
        private readonly string _templatePath;
        private readonly ITemplateParser _parser;
        private ExcelPackage _package;
        private readonly Dictionary<string, object> _variables = new Dictionary<string, object>();
        private readonly Dictionary<string, Func<object, object>> _registeredFunctions = new Dictionary<string, Func<object, object>>();
        private object _rootValue;

        public TemplateEngine(string templatePath)
        {
            _templatePath = templatePath ?? throw new ArgumentNullException(nameof(templatePath));
            _package = new ExcelPackage(new FileInfo(templatePath));
            _parser = new TemplateParser();
        }

        public TemplateEngine(Stream stream)
        {
            if (stream == null)
                throw new ArgumentNullException(nameof(stream));
            _package = new ExcelPackage(stream);
            _parser = new TemplateParser();
        }

        public void AddVariable(object value)
        {
            _rootValue = value;
        }

        public void AddVariable(string name, object value)
        {
            if (string.IsNullOrWhiteSpace(name))
                throw new ArgumentException("Variable name cannot be empty", nameof(name));
            _variables[name] = value;
        }

        public void RegisterFunction(string name, Func<object, object> func)
        {
            if (string.IsNullOrWhiteSpace(name))
                throw new ArgumentException("Function name cannot be empty", nameof(name));
            _registeredFunctions[name] = func ?? throw new ArgumentNullException(nameof(func));
        }

        public TemplateGenerateResult Generate(GenerateOptions options = null)
        {
            var parsingErrors = new TemplateErrors();
            var renderingErrors = new TemplateErrors();
            var warnings = new TemplateErrors();
            var tracker = new Rendering.RowOperationTracker();
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

            var adjuster = new Rendering.ExcelTableAdjuster(tracker);
            adjuster.AdjustAll(_package);

            if (options != null && options.EvaluateFormulas)
            {
                _package.Workbook.Calculate();
            }

            return new TemplateGenerateResult(parsingErrors, renderingErrors, warnings);
        }

        public void Save()
        {
            if (string.IsNullOrEmpty(_templatePath))
                throw new InvalidOperationException("Cannot save when engine was created from a stream. Use SaveAs instead.");
            _package.SaveAs(new FileInfo(_templatePath));
        }

        public void SaveAs(string path)
        {
            _package.SaveAs(new FileInfo(path));
        }

        public void SaveAs(FileInfo fileInfo)
        {
            _package.SaveAs(fileInfo);
        }

        public void SaveAs(Stream stream)
        {
            _package.SaveAs(stream);
        }

        public void SaveAs(string path, SaveOptions saveOptions)
        {
            ApplySaveOptions(saveOptions);
            _package.SaveAs(new FileInfo(path));
        }

        public void SaveAs(FileInfo fileInfo, SaveOptions saveOptions)
        {
            ApplySaveOptions(saveOptions);
            _package.SaveAs(fileInfo);
        }

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
}
