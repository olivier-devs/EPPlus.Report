# EPPlus.Report

Template engine for Excel reports using EPPlus.

## Installation

```bash
dotnet add package EPPlus.Report
```

> **Important:** EPPlus requires a license context. Set it before using the library:
> ```csharp
> ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // or Commercial
> ```

## Quick Start

```csharp
using EPPlus.Report;
using OfficeOpenXml;

ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

var engine = new TemplateEngine("template.xlsx");

engine.AddVariable(new
{
    Title = "Sales Report",
    Items = new[]
    {
        new { Name = "Product A", Price = 10.99m },
        new { Name = "Product B", Price = 20.50m }
    }
});

var result = engine.Generate();

if (!result.HasErrors)
{
    engine.SaveAs("output.xlsx");
}
```

## Template Syntax

| Feature | Syntax |
|---------|--------|
| Expressions | `{{PropertyName}}`, `{{Object.Property}}` |
| Functions | `{{Upper(Name)}}`, `{{Trim(Address.City)}}` |
| Loops | `<<foreach Items>> ... <</foreach>>` |
| Conditions | `<<if Condition>> ... <</if>>` |
| Aggregations | `<<sum Price>>`, `<<count Items>>` |
| Grouping | `<<group Items by Category>> ... <</group>>` |
| Named Range | Define an Excel Named Range matching your variable name |

## Advanced Usage

### Named Variables

```csharp
var engine = new TemplateEngine("report.xlsx");
engine.AddVariable("clients", clientList);
engine.AddVariable("company", companyInfo);
var result = engine.Generate();
engine.SaveAs("output.xlsx");
```

### Custom Functions

```csharp
var engine = new TemplateEngine("template.xlsx");
engine.AddVariable(data);
engine.RegisterFunction("Double", x => x?.ToString() + x?.ToString());
var result = engine.Generate();
engine.SaveAs("output.xlsx");
```

### Formula Evaluation

```csharp
var result = engine.Generate(new GenerateOptions { EvaluateFormulas = true });
// or
engine.SaveAs("output.xlsx", new SaveOptions { EvaluateFormulasBeforeSave = true });
```

### Error Handling

```csharp
var result = engine.Generate();

if (result.HasErrors)
{
    foreach (var error in result.ParsingErrors)
        Console.WriteLine($"[PARSE] {error.Location}: {error.Message}");
    foreach (var error in result.RenderingErrors)
        Console.WriteLine($"[RENDER] {error.Location}: {error.Message}");
}

if (result.HasWarnings)
{
    foreach (var warning in result.Warnings)
        Console.WriteLine($"[WARN] {warning.Location}: {warning.Message}");
}
```

## Multi-targeting

This library targets `net47`, `net48`, and `net8.0`.

## Documentation

- [Usage Guide](docs/usage-guide.md) - Step-by-step guide with examples
- [API Reference](docs/api-reference.md) - Full API documentation
- [Architecture](ARCHITECTURE.md) - Design decisions and data flow
- [Français](README.fr.md) - Version française

## License

MIT
