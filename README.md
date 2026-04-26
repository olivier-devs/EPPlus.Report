# EPPlus.Report

Template engine for Excel reports using EPPlus.

## Installation

```bash
dotnet add package EPPlus.Report
```

## Usage

```csharp
var engine = new TemplateEngine("template.xlsx");
engine.Render(data);
engine.Save("output.xlsx");
```

## Template Syntax

- Expressions: `{{PropertyName}}`
- Loops: `<<foreach Items>> ... <</foreach>>`
- Conditions: `<<if Condition>> ... <</if>>`
- Aggregations: `<<sum Property>>`, `<<count Items>>`

## Example

```csharp
var engine = new TemplateEngine("template.xlsx");
var data = new
{
    Title = "Sales Report",
    Items = new[]
    {
        new { Name = "Product A", Price = 10.99m },
        new { Name = "Product B", Price = 20.50m }
    }
};
engine.Render(data);
engine.Save("output.xlsx");
```

## Multi-targeting

This library targets `net47`, `net48`, and `net8.0`.
