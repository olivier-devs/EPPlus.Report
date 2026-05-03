# EPPlus.Report - Référence API

Référence complète de l'API publique.

---

## `TemplateEngine`

Point d'entrée principal pour la génération Excel basée sur des templates. Inspiré par ClosedXML.Report.

### Constructeurs

#### `TemplateEngine(string templatePath)`

Crée un nouveau moteur à partir d'un fichier template Excel sur disque.

**Paramètres :**
- `templatePath` (`string`) : Chemin vers le fichier template `.xlsx`.

**Exceptions :**
- `ArgumentNullException` : Si `templatePath` est null.

#### `TemplateEngine(Stream stream)`

Crée un nouveau moteur à partir d'un stream contenant un classeur Excel.

**Paramètres :**
- `stream` (`Stream`) : Stream lisible contenant un fichier `.xlsx` valide.

**Exceptions :**
- `ArgumentNullException` : Si `stream` est null.

---

### Méthodes

#### `AddVariable(object value)`

Définit l'objet de contexte racine. Utilisé quand une expression ne correspond à aucune variable nommée.

**Paramètres :**
- `value` (`object`) : L'objet de données racine.

#### `AddVariable(string name, object value)`

Ajoute une variable nommée accessible dans les templates via `{{Name}}`.

**Paramètres :**
- `name` (`string`) : Nom de la variable utilisé dans les templates.
- `value` (`object`) : Valeur de la variable.

**Exceptions :**
- `ArgumentException` : Si `name` est null, vide, ou constitué uniquement d'espaces.

**Priorité de résolution :** Les variables nommées sont résolues en premier. Si aucune variable nommée ne correspond à une expression, le contexte racine (`AddVariable(object)`) est utilisé comme fallback.

#### `Generate()`

Parse toutes les feuilles, effectue le rendu du template, et retourne le résultat.

**Retour :** `TemplateGenerateResult` contenant les éventuelles erreurs de parsing ou de rendu.

**Comportement :**
- Modifie le `ExcelPackage` interne en mémoire.
- N'écrit pas sur le disque.
- Les formules ne sont pas évaluées (laissées à Excel à l'ouverture).

#### `Generate(GenerateOptions options)`

Effectue le rendu du template avec des options supplémentaires.

**Paramètres :**
- `options` (`GenerateOptions`) : Options de génération.

**Retour :** `TemplateGenerateResult`

#### `RegisterFunction(string name, Func<object, object> func)`

Enregistre une fonction personnalisée utilisable dans les expressions template.

**Paramètres :**
- `name` (`string`) : Nom de la fonction utilisée dans les templates.
- `func` (`Func<object, object>`) : Implémentation de la fonction.

**Exceptions :**
- `ArgumentException` : Si `name` est null, vide, ou constitué uniquement d'espaces.
- `ArgumentNullException` : Si `func` est null.

**Exemple :**
```csharp
engine.RegisterFunction("Double", x => x?.ToString() + x?.ToString());
// Template : {{Double(Name)}}
```

**Fonctions intégrées :** `Upper`, `Lower`, `Trim`.

---

### Méthodes de sauvegarde

#### `Save()`

écrase le fichier template original.

**Exceptions :**
- `InvalidOperationException` : Si le moteur a été créé à partir d'un stream (aucun chemin d'origine n'existe).

#### `SaveAs(string path)`

Sauvegarde vers un nouveau chemin de fichier.

**Paramètres :**
- `path` (`string`) : Chemin du fichier de sortie.

#### `SaveAs(FileInfo fileInfo)`

Sauvegarde en utilisant un `FileInfo`.

**Paramètres :**
- `fileInfo` (`FileInfo`) : Informations sur le fichier cible.

#### `SaveAs(Stream stream)`

écrit le classeur dans un stream.

**Paramètres :**
- `stream` (`Stream`) : Stream writable.

#### `SaveAs(string path, SaveOptions saveOptions)`

Sauvegarde vers un chemin avec des options.

**Paramètres :**
- `path` (`string`) : Chemin du fichier de sortie.
- `saveOptions` (`SaveOptions`) : Options de sauvegarde (ex: évaluation des formules).

#### `SaveAs(FileInfo fileInfo, SaveOptions saveOptions)`

Sauvegarde en utilisant `FileInfo` avec des options.

#### `SaveAs(Stream stream, SaveOptions saveOptions)`

écrit dans un stream avec des options.

---

## `TemplateGenerateResult`

Objet de résultat retourné par `TemplateEngine.Generate()`.

### Propriétés

#### `HasErrors` (get)

`bool` - `true` si au moins une erreur a été collectée pendant le parsing ou le rendu.

#### `HasWarnings` (get)

`bool` - `true` si au moins un avertissement non bloquant a été collecté pendant le rendu.

#### `ParsingErrors` (get)

`TemplateErrors` - Collection des erreurs de parsing (ex: blocs non fermés). Vide si aucune erreur.

#### `RenderingErrors` (get)

`TemplateErrors` - Collection des erreurs de rendu (ex: échec d'évaluation de formule). Vide si aucune erreur.

#### `Warnings` (get)

`TemplateErrors` - Collection des avertissements non bloquants (ex: propriétés manquantes). Le rendu continue. Vide si aucun avertissement.

---

## `TemplateErrors`

Hérite de `List<TemplateError>`. Collection d'erreurs de template.

---

## `TemplateError`

Représente une erreur unique rencontrée pendant le traitement du template.

### Propriétés

| Propriété | Type | Description |
|-----------|------|-------------|
| `Message` | `string` | Description lisible de l'erreur. |
| `CellAddress` | `string` | Adresse de cellule Excel (ex: `"A5"`). |
| `WorksheetName` | `string` | Nom de la feuille où l'erreur s'est produite. |
| `Row` | `int` | Numéro de ligne (1-based). |
| `Column` | `int` | Numéro de colonne (1-based). |
| `Expression` | `string` | L'expression template qui a causé l'erreur. |
| `Location` | `string` | Emplacement combiné : `WorksheetName!CellAddress`. |
| `Type` | `ErrorType` | Catégorie de l'erreur : `Parsing`, `Evaluation`, `Rendering` ou `Warning`. |

---

## `ErrorType`

```csharp
public enum ErrorType
{
    Parsing,     // Erreurs de syntaxe template (blocs non fermés, etc.)
    Evaluation,  // Erreurs d'évaluation d'expression (fatales)
    Rendering,   // Erreurs Excel pendant le rendu
    Warning      // Problèmes non bloquants (ex: propriétés manquantes)
}
```

---

## `GenerateOptions`

Options pour `TemplateEngine.Generate(GenerateOptions)`.

### Propriétés

| Propriété | Type | Défaut | Description |
|-----------|------|--------|-------------|
| `EvaluateFormulas` | `bool` | `false` | Si `true`, évalue toutes les formules Excel après le rendu via `package.Workbook.Calculate()`. |

---

## `SaveOptions`

Options pour les opérations de sauvegarde.

### Propriétés

| Propriété | Type | Défaut | Description |
|-----------|------|--------|-------------|
| `EvaluateFormulasBeforeSave` | `bool` | `false` | Si `true`, évalue les formules juste avant l'écriture sur disque/stream. |

> **Note :** Si `GenerateOptions.EvaluateFormulas` était déjà `true`, les formules ne sont pas réévaluées pendant la sauvegarde.

---

## `RenderContext` (Avancé)

Passé en interne au renderer. Disponible pour un usage avancé.

### Propriétés

| Propriété | Type | Description |
|-----------|------|-------------|
| `Current` | `object` | Contexte de données courant (racine ou élément de boucle). |
| `Variables` | `Dictionary<string, object>` | Dictionnaire de variables nommées. |
| `CurrentCollection` | `IEnumerable` | Contexte de collection courant pour les agrégations. |

---

## Interfaces (Points d'extension)

### `ITemplateParser`

```csharp
public interface ITemplateParser
{
    Template Parse(ExcelWorksheet worksheet, TemplateErrors errors);
}
```

### `ITemplateRenderer`

```csharp
public interface ITemplateRenderer
{
    void Render(Template template, RenderContext context, ExcelWorksheet worksheet);
}
```

### `IExpressionEvaluator`

```csharp
public interface IExpressionEvaluator
{
    object Evaluate(string expression, object context);
}
```

---

## Directives de Template

| Directive | Syntaxe | Description |
|-----------|---------|-------------|
| Expression | `{{Property}}` | Remplacée par la valeur de la propriété |
| Expression imbriquée | `{{Object.Property}}` | Remplacée par la valeur de la propriété imbriquée |
| Boucle | `<<foreach Items>>` | Répète le bloc pour chaque élément |
| Fin de boucle | `<</foreach>>` | Ferme le bloc de boucle |
| Condition | `<<if Condition>>` | Affiche le bloc si la condition est vraie |
| Fin de condition | `<</if>>` | Ferme le bloc de condition |
| Somme | `<<sum Property>>` | Calcule la somme d'une propriété sur la collection |
| Comptage | `<<count Items>>` | Compte les éléments d'une collection |
| Groupement | `<<group Items by Category>>` | Groupe une collection par propriété avec sous-totaux |
| Fin de groupement | `<</group>>` | Ferme le bloc de groupement |

---

## Fonctions personnalisées

Les fonctions transforment les valeurs d'expression au moment du rendu.

### Fonctions intégrées

| Fonction | Exemple | Résultat |
|----------|---------|--------|
| `Upper` | `{{Upper(Name)}}` | Convertit en majuscules |
| `Lower` | `{{Lower(Name)}}` | Convertit en minuscules |
| `Trim` | `{{Trim(Name)}}` | Supprime les espaces de début/fin |

### Fonctions enregistrées par l'utilisateur

Enregistrez via `TemplateEngine.RegisterFunction(name, func)` :

```csharp
engine.RegisterFunction("Double", x => x?.ToString() + x?.ToString());
// Template : {{Double(Name)}}
```

Les fonctions sont appliquées après la résolution de la propriété. Si la valeur est `null`, la fonction reçoit `null`.

---

## Groupement

Groupe une collection par une ou plusieurs propriétés, avec sous-totaux et total général optionnels.

### Syntaxe explicite par bloc

```
<<group Items by Category>>
{{Category}}
{{Name}} | {{Price}}
<</group>>
```

**Options :**
- Suffixe `desc` pour le tri décroissant : `<<group Items by Category desc>>`
- `MergeLabels` fusionne les cellules de clé de groupe verticalement
- La rangée de template de sous-total (dernière ligne avec seulement des nœuds d'agrégation) est auto-détectée
- Le total général est rendu après tous les groupes (sauf désactivé)

### Groupement via Named Range

Dans les tables planches via plages nommées, utilisez la syntaxe de rangée de service :

```
<<group Category>>
```

Cela groupe la boucle de plage nommée par la propriété spécifiée et rend des rangées de sous-total par groupe.

---

## Named Range Flat Tables (Tables planches via plages nommées)

Alternative aux boucles `<<foreach>>` utilisant les plages nommées Excel.

### Fonctionnement

1. Créez une **plage nommée** dans Excel avec le même nom que votre variable collection.
2. À l'intérieur de la plage nommée, utilisez `{{item.Property}}` pour référencer les propriétés des éléments.
3. La dernière ligne de la plage nommée est la **rangée de service** pour les tags comme `<<sum>>`.

### Exigences

- La plage nommée doit être rectangulaire et continue (pas de trous).
- Tables verticales : au moins 2 lignes et 2 colonnes.
- La dernière ligne est réservée pour les tags de service.

### Variables spéciales

| Variable | Résolution |
|----------|------------|
| `item` | Élément courant de la collection |
| `item.Property` | Propriété de l'élément courant |
| `index` | Index 0-based de l'élément courant |
| `items` | La collection entière |

### Tags de la rangée de service

| Tag | Description |
|-----|-------------|
| `<<sum>>` | Insère une formule `=SUBTOTAL(9, ...)` pour la colonne |
| `<<count>>` | Insère une formule `=SUBTOTAL(3, ...)` pour la colonne |

### Exemple

**Template Excel :** Plage nommée `Orders` couvrant `A1:B3`

| | A | B |
|---|---|---|
| 1 | Order No | Amount |
| 2 | {{item.OrderNo}} | {{item.Amount}} |
| 3 | <<sum>> | |

**Code :**
```csharp
var engine = new TemplateEngine("template.xlsx");
engine.AddVariable("Orders", new[]
{
    new { OrderNo = 100, Amount = 50m },
    new { OrderNo = 101, Amount = 75m }
});
var result = engine.Generate();
engine.SaveAs("output.xlsx");
```

**Résultat :** Les lignes A2:B2 sont dupliquées pour chaque commande, avec `=SUBTOTAL(9, ...)` en A3.

---

## Thread Safety

`TemplateEngine` n'est **pas thread-safe**. Créez une instance par opération de génération.