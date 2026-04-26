# EPPlus.Report - Usage Guide

Step-by-step guide with practical examples.

---

## Table of Contents

1. [Getting Started](#getting-started)
2. [Basic Usage](#basic-usage)
3. [Named Variables](#named-variables)
4. [Loops](#loops)
5. [Named Range Flat Tables](#named-range-flat-tables)
6. [Conditions](#conditions)
7. [Aggregations](#aggregations)
8. [Custom Functions](#custom-functions)
9. [Grouping](#grouping)
10. [Formula Evaluation](#formula-evaluation)
11. [Working with Streams](#working-with-streams)
12. [Error Handling](#error-handling)
13. [ASP.NET Core Example](#aspnet-core-example)

---

## Getting Started

### Installation

```bash
dotnet add package EPPlus.Report
```

> **Important:** EPPlus requires a license context. Set it before using the library:
>
> ```csharp
> ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // or Commercial
> ```

### Creating Your First Template

Create an Excel file (`invoice_template.xlsx`) with this content:

| | A | B | C |
|---|---|---|---|
| 1 | Invoice for: | {{CustomerName}} | |
| 2 | Date: | {{InvoiceDate}} | |
| 3 | | | |
| 4 | Item | Qty | Price |
| 5 | <<foreach Items>> | | |
| 6 | {{Name}} | {{Quantity}} | {{Price}} |
| 7 | <</foreach>> | | |
| 8 | | | |
| 9 | Total: | | <<sum Price>> |

---

## Basic Usage

The simplest way to generate a report:

```csharp
using EPPlus.Report;
using OfficeOpenXml;

ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

var engine = new TemplateEngine("invoice_template.xlsx");

engine.AddVariable(new
{
    CustomerName = "Acme Corp",
    InvoiceDate = DateTime.Now.ToString("yyyy-MM-dd"),
    Items = new[]
    {
        new { Name = "Widget", Quantity = 2, Price = 10.50m },
        new { Name = "Gadget", Quantity = 1, Price = 25.00m }
    }
});

var result = engine.Generate();

if (!result.HasErrors)
{
    engine.SaveAs("invoice_output.xlsx");
    Console.WriteLine("Invoice generated successfully!");
}
else
{
    foreach (var error in result.ParsingErrors)
    {
        Console.WriteLine($"Error in {error.CellAddress}: {error.Message}");
    }
}
```

---

## Named Variables

Use `AddVariable(string name, object value)` to pass multiple independent data sources:

```csharp
var engine = new TemplateEngine("report.xlsx");

// Pass different objects with clear names
engine.AddVariable("company", new { Name = "Acme Corp", Address = "123 Main St" });
engine.AddVariable("clients", new[]
{
    new { Name = "Alice", City = "Paris" },
    new { Name = "Bob", City = "London" }
});
engine.AddVariable("reportDate", DateTime.Now);

var result = engine.Generate();
engine.SaveAs("report_output.xlsx");
```

**Template (`report.xlsx`):**

| | A |
|---|---|
| 1 | {{company.Name}} |
| 2 | {{company.Address}} |
| 3 | Report date: {{reportDate}} |
| 4 | |
| 5 | <<foreach clients>> |
| 6 | {{Name}} - {{City}} |
| 7 | <</foreach>> |

---

## Loops

Repeat rows for each item in a collection:

```csharp
var engine = new TemplateEngine("products.xlsx");

var products = new[]
{
    new { Name = "Laptop", Category = "Electronics", Price = 999.99m },
    new { Name = "Mouse", Category = "Electronics", Price = 29.99m },
    new { Name = "Desk", Category = "Furniture", Price = 199.99m }
};

engine.AddVariable("products", products);
var result = engine.Generate();
engine.SaveAs("products_output.xlsx");
```

**Template (`products.xlsx`):**

| | A | B | C |
|---|---|---|---|
| 1 | Product | Category | Price |
| 2 | <<foreach products>> | | |
| 3 | {{Name}} | {{Category}} | {{Price}} |
| 4 | <</foreach>> | | |

---

## Named Range Flat Tables

An alternative to `<<foreach>>` using Excel Named Ranges. No loop tags needed—just create a named range with the same name as your variable.

```csharp
var engine = new TemplateEngine("orders.xlsx");

var orders = new[]
{
    new { OrderNo = 100, Amount = 50m },
    new { OrderNo = 101, Amount = 75m },
    new { OrderNo = 102, Amount = 30m }
};

engine.AddVariable("Orders", orders);
var result = engine.Generate();
engine.SaveAs("orders_output.xlsx");
```

**Template (`orders.xlsx`):**

Create a Named Range called `Orders` covering `A1:B3`:

| | A | B |
|---|---|---|
| 1 | Order No | Amount |
| 2 | {{item.OrderNo}} | {{item.Amount}} |
| 3 | <<sum>> | |

**How it works:**
- `{{item.Property}}` references the current element's property
- `{{index}}` gives the 0-based index
- The last row is the **service row** for aggregation tags
- `<<sum>>` inserts a `=SUBTOTAL(9, ...)` formula
- `<<count>>` inserts a `=SUBTOTAL(3, ...)` formula

---

## Conditions

Show or hide sections based on boolean values:

```csharp
var engine = new TemplateEngine("conditional.xlsx");

engine.AddVariable(new
{
    ShowHeader = true,
    ShowFooter = false,
    Title = "Quarterly Report"
});

var result = engine.Generate();
engine.SaveAs("conditional_output.xlsx");
```

**Template (`conditional.xlsx`):**

| | A |
|---|---|
| 1 | <<if ShowHeader>> |
| 2 | {{Title}} |
| 3 | <</if>> |
| 4 | Main content here |
| 5 | <<if ShowFooter>> |
| 6 | Footer text |
| 7 | <</if>> |

---

## Aggregations

### Sum

Calculate the sum of a numeric property across a collection:

```csharp
// Inside a loop:
// <<foreach Items>>
// {{Name}} | {{Price}}
// Total: <<sum Price>>
// <</foreach>>
```

### Count

Count items in a collection:

```csharp
// Total items: <<count Items>>
```

### Combined Example

```csharp
var engine = new TemplateEngine("sales.xlsx");

var sales = new[]
{
    new { Product = "A", Amount = 100m },
    new { Product = "B", Amount = 200m },
    new { Product = "C", Amount = 150m }
};

engine.AddVariable("sales", sales);
var result = engine.Generate();
engine.SaveAs("sales_output.xlsx");
```

**Template (`sales.xlsx`):**

| | A | B |
|---|---|---|
| 1 | Product | Amount |
| 2 | <<foreach sales>> | |
| 3 | {{Product}} | {{Amount}} |
| 4 | <</foreach>> | |
| 5 | Count: | <<count sales>> |
| 6 | Total: | <<sum Amount>> |

---

## Custom Functions

Transform expression values at render time.

### Built-in Functions

```csharp
var engine = new TemplateEngine("products.xlsx");
engine.AddVariable(new { Name = "widget" });
var result = engine.Generate();
engine.SaveAs("products_output.xlsx");
```

**Template (`products.xlsx`):**

| | A |
|---|---|
| 1 | {{Upper(Name)}} |
| 2 | {{Lower(Name)}} |
| 3 | {{Trim(Name)}} |

**Result:**
- A1 = `WIDGET`
- A2 = `widget`
- A3 = `widget` (trimmed)

### Registering Custom Functions

```csharp
var engine = new TemplateEngine("report.xlsx");
engine.AddVariable(new { Code = "ABC" });

// Register a custom function
engine.RegisterFunction("Prefix", x => $"REF-{x}");

var result = engine.Generate();
engine.SaveAs("report_output.xlsx");
```

**Template (`report.xlsx`):**

| | A |
|---|---|
| 1 | {{Prefix(Code)}} |

**Result:** A1 = `REF-ABC`

Functions receive the resolved property value (or `null` if the property is null/missing). They must return an `object`.

---

## Grouping

Group a collection by one or more properties, with optional subtotals and grand totals.

### Basic Grouping

```csharp
var engine = new TemplateEngine("sales.xlsx");

var sales = new[]
{
    new { Product = "A", Category = "Electronics", Amount = 100m },
    new { Product = "B", Category = "Electronics", Amount = 200m },
    new { Product = "C", Category = "Furniture", Amount = 150m }
};

engine.AddVariable("sales", sales);
var result = engine.Generate();
engine.SaveAs("sales_output.xlsx");
```

**Template (`sales.xlsx`):**

| | A | B | C |
|---|---|---|---|
| 1 | Category | Product | Amount |
| 2 | <<group sales by Category>> | | |
| 3 | {{Category}} | {{Product}} | {{Amount}} |
| 4 | | | <<sum Amount>> |
| 5 | <</group>> | | |

**Result:** Items are grouped by Category. Each group shows its items, followed by a subtotal row. The last row with only aggregation nodes (`<<sum>>`) becomes the subtotal template.

### Named Range Grouping

In Named Range flat tables, use the service row syntax:

| | A | B |
|---|---|---|
| 1 | Product | Amount |
| 2 | {{item.Product}} | {{item.Amount}} |
| 3 | <<group Category>> | |
| 4 | <<sum>> | |

This groups items by `Category` and inserts a subtotal row per group.

---

## Formula Evaluation

By default, formulas in your template are preserved and evaluated by Excel when the file is opened. You can force evaluation during generation:

```csharp
var engine = new TemplateEngine("formulas.xlsx");
engine.AddVariable(new { Value1 = 10, Value2 = 20 });

// Evaluate formulas immediately after rendering
var result = engine.Generate(new GenerateOptions { EvaluateFormulas = true });
engine.SaveAs("formulas_output.xlsx");
```

Or evaluate just before saving:

```csharp
var engine = new TemplateEngine("formulas.xlsx");
engine.AddVariable(new { Value1 = 10, Value2 = 20 });

var result = engine.Generate();
engine.SaveAs("formulas_output.xlsx", new SaveOptions { EvaluateFormulasBeforeSave = true });
```

**Template (`formulas.xlsx`):**

| | A | B |
|---|---|---|
| 1 | Value 1 | {{Value1}} |
| 2 | Value 2 | {{Value2}} |
| 3 | Sum | =B1+B2 |

---

## Working with Streams

Useful for web applications where you want to serve files directly without writing to disk:

```csharp
using EPPlus.Report;
using OfficeOpenXml;
using System.IO;

ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

// Load template from stream
using var templateStream = File.OpenRead("template.xlsx");
var engine = new TemplateEngine(templateStream);

engine.AddVariable(new { Name = "Alice", Age = 30 });
var result = engine.Generate();

// Write to output stream
using var outputStream = new MemoryStream();
engine.SaveAs(outputStream);

// Now outputStream contains the generated Excel file
// In ASP.NET Core:
// return File(outputStream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "report.xlsx");
```

---

## Error Handling

Always check `result.HasErrors` and `result.HasWarnings` after generation:

```csharp
var result = engine.Generate();

if (result.HasErrors)
{
    foreach (var error in result.ParsingErrors)
    {
        Console.WriteLine($"[PARSE] {error.Location}: {error.Message}");
    }
    foreach (var error in result.RenderingErrors)
    {
        Console.WriteLine($"[RENDER] {error.Location}: {error.Message}");
    }
}

if (result.HasWarnings)
{
    foreach (var warning in result.Warnings)
    {
        Console.WriteLine($"[WARN] {warning.Location}: {warning.Message}");
    }
}
```

**Common error types:**

| ErrorType | Cause | Severity |
|-----------|-------|----------|
| `Parsing` | Unclosed `<<foreach>>` or `<<if>>` block | Usually non-fatal |
| `Parsing` | Malformed directive syntax | Non-fatal (treated as text) |
| `Evaluation` | Property not found on context object | Fatal for that expression |
| `Rendering` | Excel operation failure | Fatal for that operation |
| `Warning` | Missing property (rendering continues) | Non-fatal |

---

## ASP.NET Core Example

Complete controller action for generating a downloadable Excel report:

```csharp
using EPPlus.Report;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;

public class ReportsController : Controller
{
    [HttpGet("api/reports/sales")]
    public IActionResult GetSalesReport()
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        // Load template (could be cached)
        var engine = new TemplateEngine("Templates/sales_report.xlsx");

        // Add data
        engine.AddVariable("reportDate", DateTime.Now);
        engine.AddVariable("sales", GetSalesData()); // Your data source

        // Generate
        var result = engine.Generate();

        if (result.HasErrors)
        {
            return BadRequest(new
            {
                Errors = result.ParsingErrors.Select(e => e.Message)
            });
        }

        // Return as downloadable file
        using var stream = new MemoryStream();
        engine.SaveAs(stream);
        stream.Position = 0;

        return File(
            stream.ToArray(),
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            $"sales_report_{DateTime.Now:yyyyMMdd}.xlsx"
        );
    }
}
```

---

## Best Practices

1. **Always set `LicenseContext`** before using EPPlus
2. **Check `result.HasErrors`** before saving
3. **Use named variables** (`AddVariable(name, value)`) for clarity when passing multiple objects
4. **Use `AddVariable(object)`** as the root context for simple single-object scenarios
5. **Prefer `SaveAs` over `Save`** to avoid accidentally overwriting your template
6. **Use streams** in web applications to avoid disk I/O
7. **Set `EvaluateFormulas = true`** only if downstream consumers cannot evaluate formulas themselves
