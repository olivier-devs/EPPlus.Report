# EPPlus.Report

Moteur de templates pour rapports Excel basé sur EPPlus.

## Installation

```bash
dotnet add package EPPlus.Report
```

> **Important :** EPPlus nécessite un contexte de licence. Définissez-le avant d'utiliser la bibliothèque :
> ```csharp
> ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // ou Commercial
> ```

## Démarrage rapide

```csharp
using EPPlus.Report;
using OfficeOpenXml;

ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

var engine = new TemplateEngine("template.xlsx");

engine.AddVariable(new
{
    Title = "Rapport des ventes",
    Items = new[]
    {
        new { Name = "Produit A", Price = 10.99m },
        new { Name = "Produit B", Price = 20.50m }
    }
});

var result = engine.Generate();

if (!result.HasErrors)
{
    engine.SaveAs("output.xlsx");
}
```

## Syntaxe des templates

| Fonctionnalité | Syntaxe |
|---------|--------|
| Expressions | `{{PropertyName}}`, `{{Object.Property}}` |
| Fonctions | `{{Upper(Name)}}`, `{{Trim(Address.City)}}` |
| Boucles | `<<foreach Items>> ... <</foreach>>` |
| Conditions | `<<if Condition>> ... <</if>>` |
| Agrégations | `<<sum Price>>`, `<<count Items>>` |
| Groupement | `<<group Items by Category>> ... <</group>>` |
| Plage nommée | Définissez une plage nommée Excel correspondant au nom de votre variable |

## Utilisation avancée

### Variables nommées

```csharp
var engine = new TemplateEngine("report.xlsx");
engine.AddVariable("clients", clientList);
engine.AddVariable("company", companyInfo);
var result = engine.Generate();
engine.SaveAs("output.xlsx");
```

### Fonctions personnalisées

```csharp
var engine = new TemplateEngine("template.xlsx");
engine.AddVariable(data);
engine.RegisterFunction("Double", x => x?.ToString() + x?.ToString());
var result = engine.Generate();
engine.SaveAs("output.xlsx");
```

### Évaluation des formules

```csharp
var result = engine.Generate(new GenerateOptions { EvaluateFormulas = true });
// ou
engine.SaveAs("output.xlsx", new SaveOptions { EvaluateFormulasBeforeSave = true });
```

### Gestion des erreurs

```csharp
var result = engine.Generate();

if (result.HasErrors)
{
    foreach (var error in result.ParsingErrors)
        Console.WriteLine($"[PARSE] {error.Location} : {error.Message}");
    foreach (var error in result.RenderingErrors)
        Console.WriteLine($"[RENDER] {error.Location} : {error.Message}");
}

if (result.HasWarnings)
{
    foreach (var warning in result.Warnings)
        Console.WriteLine($"[WARN] {warning.Location} : {warning.Message}");
}
```

## Multi-ciblage

Cette bibliothèque cible `net47`, `net48` et `net8.0`.

## Documentation

- [Guide d'utilisation](docs/usage-guide.fr.md) - Guide pas à pas avec exemples
- [Référence API](docs/api-reference.fr.md) - Documentation API complète
- [Architecture](ARCHITECTURE.md) - Décisions de conception et flux de données
- [English](README.md) - English version

## Licence

MIT
