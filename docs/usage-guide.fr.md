# EPPlus.Report - Guide d'utilisation

Guide pas à pas avec des exemples pratiques.

---

## Table des matières

1. [Démarrage](#démarrage)
2. [Utilisation basique](#utilisation-basique)
3. [Variables nommées](#variables-nommées)
4. [Boucles](#boucles)
5. [Tables planches via plages nommées](#tables-planches-via-plages-nommées)
6. [Conditions](#conditions)
7. [Agrégations](#agrégations)
8. [Fonctions personnalisées](#fonctions-personnalisées)
9. [Groupement](#groupement)
10. [Évaluation des formules](#évaluation-des-formules)
11. [Travail avec les streams](#travail-avec-les-streams)
12. [Gestion des erreurs](#gestion-des-erreurs)
13. [Exemple ASP.NET Core](#exemple-aspnet-core)

---

## Démarrage

### Installation

```bash
dotnet add package EPPlus.Report
```

> **Important :** EPPlus nécessite un contexte de licence. Définissez-le avant d'utiliser la bibliothèque :
>
> ```csharp
> ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // ou Commercial
> ```

### Créer votre premier template

Créez un fichier Excel (`facture_template.xlsx`) avec ce contenu :

| | A | B | C |
|---|---|---|---|
| 1 | Facture pour : | {{CustomerName}} | |
| 2 | Date : | {{InvoiceDate}} | |
| 3 | | | |
| 4 | Article | Qté | Prix |
| 5 | <<foreach Items>> | | |
| 6 | {{Name}} | {{Quantity}} | {{Price}} |
| 7 | <</foreach>> | | |
| 8 | | | |
| 9 | Total : | | <<sum Price>> |

---

## Utilisation basique

La façon la plus simple de générer un rapport :

```csharp
using EPPlus.Report;
using OfficeOpenXml;

ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

var engine = new TemplateEngine("facture_template.xlsx");

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
    engine.SaveAs("facture_output.xlsx");
    Console.WriteLine("Facture générée avec succès !");
}
else
{
    foreach (var error in result.ParsingErrors)
    {
        Console.WriteLine($"Erreur dans {error.CellAddress} : {error.Message}");
    }
}
```

---

## Variables nommées

Utilisez `AddVariable(string name, object value)` pour passer plusieurs sources de données indépendantes :

```csharp
var engine = new TemplateEngine("rapport.xlsx");

// Passez différents objets avec des noms clairs
engine.AddVariable("company", new { Name = "Acme Corp", Address = "123 Rue Principale" });
engine.AddVariable("clients", new[]
{
    new { Name = "Alice", City = "Paris" },
    new { Name = "Bob", City = "Londres" }
});
engine.AddVariable("reportDate", DateTime.Now);

var result = engine.Generate();
engine.SaveAs("rapport_output.xlsx");
```

**Template (`rapport.xlsx`) :**

| | A |
|---|---|
| 1 | {{company.Name}} |
| 2 | {{company.Address}} |
| 3 | Date du rapport : {{reportDate}} |
| 4 | |
| 5 | <<foreach clients>> |
| 6 | {{Name}} - {{City}} |
| 7 | <</foreach>> |

---

## Boucles

Répétez des lignes pour chaque élément d'une collection :

```csharp
var engine = new TemplateEngine("produits.xlsx");

var products = new[]
{
    new { Name = "Ordinateur portable", Category = "Électronique", Price = 999.99m },
    new { Name = "Souris", Category = "Électronique", Price = 29.99m },
    new { Name = "Bureau", Category = "Mobilier", Price = 199.99m }
};

engine.AddVariable("products", products);
var result = engine.Generate();
engine.SaveAs("produits_output.xlsx");
```

**Template (`produits.xlsx`) :**

| | A | B | C |
|---|---|---|---|
| 1 | Produit | Catégorie | Prix |
| 2 | <<foreach products>> | | |
| 3 | {{Name}} | {{Category}} | {{Price}} |
| 4 | <</foreach>> | | |

---

## Tables planches via plages nommées

Alternative aux `<<foreach>>` utilisant les plages nommées Excel. Pas besoin de tags de boucle — créez simplement une plage nommée avec le même nom que votre variable.

```csharp
var engine = new TemplateEngine("commandes.xlsx");

var orders = new[]
{
    new { OrderNo = 100, Amount = 50m },
    new { OrderNo = 101, Amount = 75m },
    new { OrderNo = 102, Amount = 30m }
};

engine.AddVariable("Orders", orders);
var result = engine.Generate();
engine.SaveAs("commandes_output.xlsx");
```

**Template (`commandes.xlsx`) :**

Créez une plage nommée appelée `Orders` couvrant `A1:B3` :

| | A | B |
|---|---|---|
| 1 | Order No | Amount |
| 2 | {{item.OrderNo}} | {{item.Amount}} |
| 3 | <<sum>> | |

**Fonctionnement :**
- `{{item.Property}}` référence la propriété de l'élément courant
- `{{index}}` donne l'index 0-based
- La dernière ligne est la **rangée de service** pour les tags d'agrégation
- `<<sum>>` insère une formule `=SUBTOTAL(9, ...)`
- `<<count>>` insère une formule `=SUBTOTAL(3, ...)`

---

## Conditions

Affichez ou masquez des sections en fonction de valeurs booléennes :

```csharp
var engine = new TemplateEngine("conditionnel.xlsx");

engine.AddVariable(new
{
    ShowHeader = true,
    ShowFooter = false,
    Title = "Rapport trimestriel"
});

var result = engine.Generate();
engine.SaveAs("conditionnel_output.xlsx");
```

**Template (`conditionnel.xlsx`) :**

| | A |
|---|---|
| 1 | <<if ShowHeader>> |
| 2 | {{Title}} |
| 3 | <</if>> |
| 4 | Contenu principal ici |
| 5 | <<if ShowFooter>> |
| 6 | Texte du pied de page |
| 7 | <</if>> |

---

## Agrégations

### Somme

Calculez la somme d'une propriété numérique sur une collection :

```csharp
// À l'intérieur d'une boucle :
// <<foreach Items>>
// {{Name}} | {{Price}}
// Total : <<sum Price>>
// <</foreach>>
```

### Comptage

Comptez les éléments d'une collection :

```csharp
// Nombre total d'articles : <<count Items>>
```

### Exemple combiné

```csharp
var engine = new TemplateEngine("ventes.xlsx");

var sales = new[]
{
    new { Product = "A", Amount = 100m },
    new { Product = "B", Amount = 200m },
    new { Product = "C", Amount = 150m }
};

engine.AddVariable("sales", sales);
var result = engine.Generate();
engine.SaveAs("ventes_output.xlsx");
```

**Template (`ventes.xlsx`) :**

| | A | B |
|---|---|---|
| 1 | Produit | Montant |
| 2 | <<foreach sales>> | |
| 3 | {{Product}} | {{Amount}} |
| 4 | <</foreach>> | |
| 5 | Nombre : | <<count sales>> |
| 6 | Total : | <<sum Amount>> |

---

## Fonctions personnalisées

Transforment les valeurs d'expression au moment du rendu.

### Fonctions intégrées

```csharp
var engine = new TemplateEngine("produits.xlsx");
engine.AddVariable(new { Name = "widget" });
var result = engine.Generate();
engine.SaveAs("produits_output.xlsx");
```

**Template (`produits.xlsx`) :**

| | A |
|---|---|
| 1 | {{Upper(Name)}} |
| 2 | {{Lower(Name)}} |
| 3 | {{Trim(Name)}} |

**Résultat :**
- A1 = `WIDGET`
- A2 = `widget`
- A3 = `widget` (trimmed)

### Enregistrer des fonctions personnalisées

```csharp
var engine = new TemplateEngine("rapport.xlsx");
engine.AddVariable(new { Code = "ABC" });

// Enregistre une fonction personnalisée
engine.RegisterFunction("Prefix", x => $"REF-{x}");

var result = engine.Generate();
engine.SaveAs("rapport_output.xlsx");
```

**Template (`rapport.xlsx`) :**

| | A |
|---|---|
| 1 | {{Prefix(Code)}} |

**Résultat :** A1 = `REF-ABC`

Les fonctions reçoivent la valeur de la propriété résolue (ou `null` si la propriété est null/manquante). Elles doivent retourner un `object`.

---

## Groupement

Groupe une collection par une ou plusieurs propriétés, avec sous-totaux et total général optionnels.

### Groupement basique

```csharp
var engine = new TemplateEngine("ventes.xlsx");

var sales = new[]
{
    new { Product = "A", Category = "Électronique", Amount = 100m },
    new { Product = "B", Category = "Électronique", Amount = 200m },
    new { Product = "C", Category = "Mobilier", Amount = 150m }
};

engine.AddVariable("sales", sales);
var result = engine.Generate();
engine.SaveAs("ventes_output.xlsx");
```

**Template (`ventes.xlsx`) :**

| | A | B | C |
|---|---|---|---|
| 1 | Catégorie | Produit | Montant |
| 2 | <<group sales by Category>> | | |
| 3 | {{Category}} | {{Product}} | {{Amount}} |
| 4 | | | <<sum Amount>> |
| 5 | <</group>> | | |

**Résultat :** Les éléments sont groupés par Catégorie. Chaque groupe affiche ses éléments, suivi d'une rangée de sous-total. La dernière ligne contenant uniquement des nœuds d'agrégation (`<<sum>>`) devient le template de sous-total.

### Groupement via Named Range

Dans les tables planches via plages nommées, utilisez la syntaxe de rangée de service :

| | A | B |
|---|---|---|
| 1 | Produit | Montant |
| 2 | {{item.Product}} | {{item.Amount}} |
| 3 | <<group Category>> | |
| 4 | <<sum>> | |

Cela groupe les éléments par `Category` et insère une rangée de sous-total par groupe.

---

## Évaluation des formules

Par défaut, les formules dans votre template sont préservées et évaluées par Excel à l'ouverture du fichier. Vous pouvez forcer l'évaluation pendant la génération :

```csharp
var engine = new TemplateEngine("formules.xlsx");
engine.AddVariable(new { Value1 = 10, Value2 = 20 });

// Évalue les formules immédiatement après le rendu
var result = engine.Generate(new GenerateOptions { EvaluateFormulas = true });
engine.SaveAs("formules_output.xlsx");
```

Ou évaluez juste avant la sauvegarde :

```csharp
var engine = new TemplateEngine("formules.xlsx");
engine.AddVariable(new { Value1 = 10, Value2 = 20 });

var result = engine.Generate();
engine.SaveAs("formules_output.xlsx", new SaveOptions { EvaluateFormulasBeforeSave = true });
```

**Template (`formules.xlsx`) :**

| | A | B |
|---|---|---|
| 1 | Valeur 1 | {{Value1}} |
| 2 | Valeur 2 | {{Value2}} |
| 3 | Somme | =B1+B2 |

---

## Travail avec les streams

Utile pour les applications web où vous voulez servir des fichiers directement sans écrire sur le disque :

```csharp
using EPPlus.Report;
using OfficeOpenXml;
using System.IO;

ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

// Charge le template depuis un stream
using var templateStream = File.OpenRead("template.xlsx");
var engine = new TemplateEngine(templateStream);

engine.AddVariable(new { Name = "Alice", Age = 30 });
var result = engine.Generate();

// Écrit dans un stream de sortie
using var outputStream = new MemoryStream();
engine.SaveAs(outputStream);

// Maintenant outputStream contient le fichier Excel généré
// Dans ASP.NET Core :
// return File(outputStream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "rapport.xlsx");
```

---

## Gestion des erreurs

Vérifiez toujours `result.HasErrors` et `result.HasWarnings` après la génération :

```csharp
var result = engine.Generate();

if (result.HasErrors)
{
    foreach (var error in result.ParsingErrors)
    {
        Console.WriteLine($"[PARSE] {error.Location} : {error.Message}");
    }
    foreach (var error in result.RenderingErrors)
    {
        Console.WriteLine($"[RENDER] {error.Location} : {error.Message}");
    }
}

if (result.HasWarnings)
{
    foreach (var warning in result.Warnings)
    {
        Console.WriteLine($"[WARN] {warning.Location} : {warning.Message}");
    }
}
```

**Types d'erreurs courants :**

| ErrorType | Cause | Gravité |
|-----------|-------|---------|
| `Parsing` | Bloc `<<foreach>>` ou `<<if>>` non fermé | Généralement non fatale |
| `Parsing` | Syntaxe de directive malformée | Non fatale (traitée comme texte) |
| `Evaluation` | Propriété non trouvée sur l'objet contexte | Fatale pour cette expression |
| `Rendering` | Échec d'opération Excel | Fatale pour cette opération |
| `Warning` | Propriété manquante (le rendu continue) | Non fatale |

---

## Exemple ASP.NET Core

Action de contrôleur complète pour générer un rapport Excel téléchargeable :

```csharp
using EPPlus.Report;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;

public class RapportsController : Controller
{
    [HttpGet("api/rapports/ventes")]
    public IActionResult GetSalesReport()
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        // Charge le template (pourrait être mis en cache)
        var engine = new TemplateEngine("Templates/rapport_ventes.xlsx");

        // Ajoute les données
        engine.AddVariable("reportDate", DateTime.Now);
        engine.AddVariable("sales", GetSalesData()); // Votre source de données

        // Génère
        var result = engine.Generate();

        if (result.HasErrors)
        {
            return BadRequest(new
            {
                Errors = result.ParsingErrors.Select(e => e.Message)
            });
        }

        // Retourne en tant que fichier téléchargeable
        using var stream = new MemoryStream();
        engine.SaveAs(stream);
        stream.Position = 0;

        return File(
            stream.ToArray(),
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            $"rapport_ventes_{DateTime.Now:yyyyMMdd}.xlsx"
        );
    }
}
```

---

## Bonnes pratiques

1. **Définissez toujours `LicenseContext`** avant d'utiliser EPPlus
2. **Vérifiez `result.HasErrors`** avant de sauvegarder
3. **Utilisez les variables nommées** (`AddVariable(name, value)`) pour la clarté quand vous passez plusieurs objets
4. **Utilisez `AddVariable(object)`** comme contexte racine pour les scénarios simples à un seul objet
5. **Préférez `SaveAs` à `Save`** pour éviter d'écraser accidentellement votre template
6. **Utilisez les streams** dans les applications web pour éviter les E/S disque
7. **Définissez `EvaluateFormulas = true`** uniquement si les consommateurs en aval ne peuvent pas évaluer les formules eux-mêmes
