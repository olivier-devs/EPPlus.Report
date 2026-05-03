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
| `IsNamedRangeLoop` | `bool` | Indique si le contexte courant est à l'intérieur d'une boucle de plage nommée. |
| `CurrentIndex` | `int` | Index 0-based de l'élément courant dans la collection. |

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

## `ExpressionEvaluator` (Avancé)

Implémentation par défaut de `IExpressionEvaluator`. Résout les chemins de propriétés via reflection avec mise en cache.

### Constructeurs

#### `ExpressionEvaluator()`

Initialise avec les fonctions intégrées : `Upper`, `Lower`, `Trim`.

### Méthodes

#### `Evaluate(string expression, object context)`

Évalue un chemin de propriété par rapport à l'objet contexte.

#### `Evaluate(string expression, object context, string functionName)`

Évalue un chemin de propriété et applique la fonction nommée au résultat.

#### `RegisterFunction(string name, Func<object, object> func)`

Enregistre une fonction personnalisée par nom.

#### `ApplyFunction(string functionName, object value)`

Applique une fonction enregistrée à une valeur.

---

## Classes du Modèle (Avancé)

### `TemplateNode` (abstract)

Classe de base pour tous les nœuds AST du template.

| Propriété | Type | Description |
|-----------|------|-------------|
| `Row` | `int` | Numéro de ligne dans la feuille. |
| `Column` | `int` | Numéro de colonne dans la feuille. |
| `RawContent` | `string` | Contenu texte brut de la cellule. |

### `ExpressionNode`

Représente une expression template `{{Property}}` ou `{{Function(Property)}}`.

| Propriété | Type | Description |
|-----------|------|-------------|
| `ExpressionPath` | `string` | Chemin de propriété à évaluer (ex: `"Object.Property"`). |
| `FunctionName` | `string` | Nom de fonction optionnel à appliquer (ex: `"Upper"`). |

### `LoopNode`

Représente un bloc `<<foreach Items>>`.

| Propriété | Type | Description |
|-----------|------|-------------|
| `CollectionName` | `string` | Nom de la collection à itérer. |
| `Children` | `List<TemplateNode>` | Nœuds enfants dans le bloc de boucle. |
| `EndRow` | `int` | Ligne où se termine le bloc de boucle. |
| `ConditionalFormattingRules` | `List<ConditionalFormattingRule>` | Règles de mise en forme conditionnelle associées à ce bloc. |

### `IfNode`

Représente un bloc `<<if Condition>>`.

| Propriété | Type | Description |
|-----------|------|-------------|
| `ConditionExpression` | `string` | Expression booléenne à évaluer. |
| `Children` | `List<TemplateNode>` | Nœuds enfants dans le bloc conditionnel. |
| `EndRow` | `int` | Ligne où se termine le bloc conditionnel. |
| `ConditionalFormattingRules` | `List<ConditionalFormattingRule>` | Règles de mise en forme conditionnelle associées à ce bloc. |

### `GroupNode`

Représente un bloc `<<group Items by Category>>`. Hérite de `LoopNode`.

| Propriété | Type | Description |
|-----------|------|-------------|
| `GroupByPaths` | `List<string>` | Chemins de propriétés utilisés pour grouper les éléments. |
| `Options` | `GroupOptions` | Options contrôlant le rendu du groupement. |
| `SubtotalTemplate` | `List<TemplateNode>` | Nœuds template pour les rangées de sous-total. |

### `NamedRangeLoopNode`

Représente une boucle dérivée d'une plage nommée Excel. Hérite de `LoopNode`.

| Propriété | Type | Description |
|-----------|------|-------------|
| `RangeName` | `string` | Nom de la plage nommée Excel. |
| `IsHorizontal` | `bool` | Indique si la boucle itère horizontalement. |
| `ServiceRowCount` | `int` | Nombre de rangées de service à la fin. |
| `ServiceTags` | `List<ServiceTag>` | Tags de service (sum, count) dans la rangée de service. |
| `EndColumn` | `int` | Colonne où se termine la plage nommée. |
| `HeaderRowCount` | `int` | Nombre de rangées d'en-tête au début. |
| `GroupByDefinitions` | `List<GroupByDefinition>` | Définitions de groupement pour la plage nommée. |
| `RangeGroupOptions` | `GroupOptions` | Options de groupement spécifiques à la plage nommée. |

### `AggregationNode`

Représente `<<sum Property>>` ou `<<count Items>>`.

| Propriété | Type | Description |
|-----------|------|-------------|
| `AggregationType` | `string` | Type d'agrégation (`"sum"` ou `"count"`). |
| `PropertyName` | `string` | Nom de propriété ou de collection à agréger. |

### `GroupOptions`

Options pour contrôler le comportement du rendu des groupes.

| Propriété | Type | Défaut | Description |
|-----------|------|--------|-------------|
| `Collapse` | `bool` | `false` | Si les rangées groupées doivent être collapseées. |
| `MergeLabels` | `MergeMode` | `None` | Comment fusionner les cellules de label de groupe. |
| `PlaceToColumn` | `int` | `0` | Index de colonne pour les labels de groupe. |
| `WithHeader` | `bool` | `false` | Si le groupe inclut un en-tête. |
| `DisableSubtotals` | `bool` | `false` | Si les rangées de sous-total doivent être désactivées. |
| `DisableOutline` | `bool` | `false` | Si le contour du groupe doit être désactivé. |
| `PageBreaks` | `bool` | `false` | Si des sauts de page doivent être insérés entre les groupes. |
| `TotalLabel` | `string` | `"Total"` | Texte de label pour les rangées de total. |
| `GrandLabel` | `string` | `"Grand"` | Texte de label pour les rangées de total général. |
| `SummaryAbove` | `bool` | `false` | Si les rangées résumé apparaissent au-dessus des détails. |
| `DisableGrandTotal` | `bool` | `false` | Si la rangée de total général doit être désactivée. |
| `Descending` | `bool` | `false` | Si les groupes doivent être triés en ordre décroissant. |

### `MergeMode`

Définit comment les labels de groupe doivent être fusionnés.

| Valeur | Description |
|-------|-------------|
| `None` | Les labels de groupe ne sont pas fusionnés. |
| `Merge1` | Fusionne les labels de groupe et efface les cellules dupliquées. |
| `Merge2` | Fusionne les labels de groupe et efface les cellules dupliquées (variante 2). |
| `Merge3` | Fusionne les labels de groupe sans effacer les cellules dupliquées. |

### `GroupByDefinition`

Définit un critère de groupement pour les boucles de plages nommées.

| Propriété | Type | Description |
|-----------|------|-------------|
| `PropertyPath` | `string` | Chemin de propriété utilisé pour extraire les clés de groupe. |
| `Column` | `int` | Index de colonne où la clé de groupe est située. |
| `Descending` | `bool` | Si le groupe doit être trié en ordre décroissant. |
| `Options` | `GroupOptions` | Options contrôlant le rendu de ce groupe. |

### `ServiceTag`

Représente un tag de service dans une plage nommée (ex: `<<sum>>`).

| Propriété | Type | Description |
|-----------|------|-------------|
| `TagName` | `string` | Nom du tag (ex: `"sum"` ou `"count"`). |
| `Row` | `int` | Ligne où le tag est situé. |
| `Column` | `int` | Colonne où le tag est situé. |

### `ConditionalFormattingRule`

Représente une règle de mise en forme conditionnelle extraite d'un bloc template.

| Propriété | Type | Description |
|-----------|------|-------------|
| `Address` | `string` | Plage d'adresses de cellules. |
| `Formula` | `string` | Formule principale. |
| `Formula2` | `string` | Formule secondaire (pour les règles nécessitant deux formules). |
| `Type` | `eExcelConditionalFormattingRuleType` | Type de règle de mise en forme conditionnelle. |
| `Priority` | `int` | Priorité de la règle. |
| `StopIfTrue` | `bool` | Si l'évaluation doit s'arrêter si cette règle est vraie. |

### `PropertyNotFoundException`

Exception levée lorsqu'une propriété référencée dans une expression template est introuvable. Hérite de `ArgumentException`.

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
- Suffixe `asc` pour le tri explicite ascendant : `<<group Items by Category asc>>`
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
| `<<sum>>` | Calcule et insère la somme des valeurs dans la colonne |
| `<<count>>` | Calcule et insère le nombre de valeurs non vides dans la colonne |

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