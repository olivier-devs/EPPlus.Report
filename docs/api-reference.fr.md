# EPPlus.Report - Référence API

Référence complète de l'API publique.

---

## `TemplateEngine`

Point d'entr�e principal pour la g�n�ration Excel bas�e sur des templates. Inspir� par ClosedXML.Report.

### Constructeurs

#### `TemplateEngine(string templatePath)`

Crée un nouveau moteur � partir d'un fichier template Excel sur disque.

**Paramètres :**
- `templatePath` (`string`) : Chemin vers le fichier template `.xlsx`.

**Exceptions :**
- `ArgumentNullException` : Si `templatePath` est null.

#### `TemplateEngine(Stream stream)`

Crée un nouveau moteur � partir d'un stream contenant un classeur Excel.

**Paramètres :**
- `stream` (`Stream`) : Stream lisible contenant un fichier `.xlsx` valide.

**Exceptions :**
- `ArgumentNullException` : Si `stream` est null.

---

### M�thodes

#### `AddVariable(object value)`

Définit l'objet de contexte racine. Utilisé quand une expression ne correspond � aucune variable nommée.

**Paramètres :**
- `value` (`object`) : L'objet de données racine.

#### `AddVariable(string name, object value)`

Ajoute une variable nomm�e accessible dans les templates via `{{Name}}`.

**Paramètres :**
- `name` (`string`) : Nom de la variable utilisé dans les templates.
- `value` (`object`) : Valeur de la variable.

**Exceptions :**
- `ArgumentException` : Si `name` est null, vide, ou constitué uniquement d'espaces.

**Priorité de résolution :** Les variables nommées sont résolues en premier. Si aucune variable nomm�e ne correspond � une expression, le contexte racine (`AddVariable(object)`) est utilisé comme fallback.

#### `Generate()`

Parse toutes les feuilles, effectue le rendu du template, et retourne le résultat.

**Retour :** `TemplateGenerateResult` contenant les éventuelles erreurs de parsing ou de rendu.

**Comportement :**
- Modifie le `ExcelPackage` interne en mémoire.
- N'�crit pas sur le disque.
- Les formules ne sont pas évaluées (laissées � Excel � l'ouverture).

#### `Generate(GenerateOptions options)`

Effectue le rendu du template avec des options suppl�mentaires.

**Paramètres :**
- `options` (`GenerateOptions`) : Options de génération.

**Retour :** `TemplateGenerateResult`

#### `RegisterFunction(string name, Func<object, object> func)`

Enregistre une fonction personnalis�e utilisable dans les expressions template.

**Paramètres :**
- `name` (`string`) : Nom de la fonction utilis�e dans les templates.
- `func` (`Func<object, object>`) : Impl�mentation de la fonction.

**Exceptions :**
- `ArgumentException` : Si `name` est null, vide, ou constitu� uniquement d'espaces.
- `ArgumentNullException` : Si `func` est null.

**Exemple :**
```csharp
engine.RegisterFunction("Double", x => x?.ToString() + x?.ToString());
// Template : {{Double(Name)}}
```

**Fonctions int�gr�es :** `Upper`, `Lower`, `Trim`.

---

### M�thodes de sauvegarde

#### `Save()`

écrase le fichier template original.

**Exceptions :**
- `InvalidOperationException` : Si le moteur a �t� créé � partir d'un stream (aucun chemin d'origine n'existe).

#### `SaveAs(string path)`

Sauvegarde vers un nouveau chemin de fichier.

**Paramètres :**
- `path` (`string`) : Chemin du fichier de sortie.

#### `SaveAs(FileInfo fileInfo)`

Sauvegarde en utilisant un `FileInfo`.

**Paramètres :**
- `fileInfo` (`FileInfo`) : Informations sur le fichier cible.

#### `SaveAs(Stream stream)`

�crit le classeur dans un stream.

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

�crit dans un stream avec des options.

---

## `TemplateGenerateResult`

Objet de r�sultat retourn� par `TemplateEngine.Generate()`.

### Propri�t�s

#### `HasErrors` (get)

`bool` - `true` si au moins une erreur a �t� collect�e pendant le parsing ou le rendu.

#### `HasWarnings` (get)

`bool` - `true` si au moins un avertissement non bloquant a �t� collect� pendant le rendu.

#### `ParsingErrors` (get)

`TemplateErrors` - Collection des erreurs de parsing (ex: blocs non ferm�s). Vide si aucune erreur.

#### `RenderingErrors` (get)

`TemplateErrors` - Collection des erreurs de rendu (ex: �chec d'�valuation de formule). Vide si aucune erreur.

#### `Warnings` (get)

`TemplateErrors` - Collection des avertissements non bloquants (ex: propri�t�s manquantes). Le rendu continue. Vide si aucun avertissement.

---

## `TemplateErrors`

H�rite de `List<TemplateError>`. Collection d'erreurs de template.

---

## `TemplateError`

Repr�sente une erreur unique rencontr�e pendant le traitement du template.

### Propri�t�s

| Propri�t� | Type | Description |
|-----------|------|-------------|
| `Message` | `string` | Description lisible de l'erreur. |
| `CellAddress` | `string` | Adresse de cellule Excel (ex: `"A5"`). |
| `WorksheetName` | `string` | Nom de la feuille o� l'erreur s'est produite. |
| `Row` | `int` | Num�ro de ligne (1-based). |
| `Column` | `int` | Num�ro de colonne (1-based). |
| `Expression` | `string` | L'expression template qui a caus� l'erreur. |
| `Location` | `string` | Emplacement combin� : `WorksheetName!CellAddress`. |
| `Type` | `ErrorType` | Cat�gorie de l'erreur : `Parsing`, `Evaluation`, `Rendering` ou `Warning`. |

---

## `ErrorType`

```csharp
public enum ErrorType
{
    Parsing,     // Erreurs de syntaxe template (blocs non ferm�s, etc.)
    Evaluation,  // Erreurs d'�valuation d'expression (fatales)
    Rendering,   // Erreurs Excel pendant le rendu
    Warning      // Probl�mes non bloquants (ex: propri�t�s manquantes)
}
```

---

## `GenerateOptions`

Options pour `TemplateEngine.Generate(GenerateOptions)`.

### Propri�t�s

| Propri�t� | Type | Défaut | Description |
|-----------|------|--------|-------------|
| `EvaluateFormulas` | `bool` | `false` | Si `true`, évalue toutes les formules Excel après le rendu via `package.Workbook.Calculate()`. |

---

## `SaveOptions`

Options pour les opérations de sauvegarde.

### Propri�t�s

| Propri�t� | Type | Défaut | Description |
|-----------|------|--------|-------------|
| `EvaluateFormulasBeforeSave` | `bool` | `false` | Si `true`, évalue les formules juste avant l'écriture sur disque/stream. |

> **Note :** Si `GenerateOptions.EvaluateFormulas` �tait déjà `true`, les formules ne sont pas r�évaluées pendant la sauvegarde.

---

## `RenderContext` (Avancé)

Pass� en interne au renderer. Disponible pour un usage avanc�.

### Propri�t�s

| Propri�t� | Type | Description |
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
