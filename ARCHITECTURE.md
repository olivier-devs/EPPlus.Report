# EPPlus.Report - Architecture

## Vue d'ensemble

EPPlus.Report fonctionne comme un **pipeline de templating Excel en 3 phases** :

```
Template Excel (.xlsx)
        ↓
    [Parser]  ──→  AST (TemplateNode)
        ↓
   [Evaluator] ──→  Valeurs résolues
        ↓
   [Renderer]  ──→  Fichier Excel final
```

## Composants

### 1. AST (Abstract Syntax Tree)

**Responsabilité** : Représenter la structure du template sous forme d'arbre de nœuds.

**Hiérarchie** :
```
TemplateNode (abstract)
├── Row, Column, RawContent
│
├── TextNode
│   └── Contenu statique
│
├── ExpressionNode
│   ├── ExpressionPath : chemin de propriété (ex: "Object.Property")
│   └── FunctionName : nom de fonction optionnel (ex: "Upper")
│
├── LoopNode
│   ├── CollectionName : nom de la collection
│   ├── Children : nœuds dans la boucle
│   ├── EndRow : ligne de fin du bloc
│   └── ConditionalFormattingRules : règles CF associées
│
├── IfNode
│   ├── ConditionExpression : expression booléenne
│   ├── Children : nœuds conditionnels
│   ├── EndRow : ligne de fin du bloc
│   └── ConditionalFormattingRules : règles CF associées
│
├── GroupNode (hérite de LoopNode)
│   ├── CollectionName : nom de la collection (hérité)
│   ├── GroupByPaths : liste de chemins de propriété (ex: ["Category"])
│   ├── Options : GroupOptions (MergeLabels, DisableSubtotals, etc.)
│   ├── SubtotalTemplate : nœuds template du sous-total
│   └── Children : nœuds dans le groupe (hérité)
│
├── NamedRangeLoopNode (hérite de LoopNode)
│   ├── RangeName : nom de la plage nommée Excel
│   ├── IsHorizontal : itération horizontale
│   ├── ServiceRowCount : nombre de rangées de service
│   ├── ServiceTags : tags de service (sum, count)
│   ├── HeaderRowCount : nombre de rangées d'en-tête
│   ├── EndColumn : colonne de fin
│   ├── GroupByDefinitions : critères de groupement
│   └── RangeGroupOptions : options de groupe spécifiques
│
└── AggregationNode
    ├── AggregationType : type ("sum" ou "count")
    └── PropertyName : propriété ou collection à agréger
```

**Classes supplémentaires** :
- `Template` : Conteneur racine (`List<TemplateNode> Nodes`)
- `RenderContext` : Contexte d'exécution (`Current`, `Variables`)

### 2. Parser

**Responsabilité** : Lire un `ExcelWorksheet` et construire l'AST.

**Algorithme** :
1. Itérer sur toutes les cellules du worksheet (`Dimension`)
2. Détecter les directives par regex :
   - `\{\{(.+?)\}\}` → ExpressionNode (ou ExpressionNode avec FunctionName si fonction)
   - `\{\{\s*([A-Za-z_][A-Za-z0-9_]*)\s*\(\s*([A-Za-z_][A-Za-z0-9_\.]*)\s*\)\s*\}\}` → ExpressionNode avec FunctionName
   - `<<foreach\s+(\w+)>>` → LoopNode (mode bloc)
   - `<<if\s+(\w+)>>` → IfNode (mode bloc)
   - `<<group\s+(\w+)\s+by\s+(.+?)>>` → GroupNode (mode bloc)
   - `<<sum\s+(\w+)>>` → AggregationNode (type "sum")
   - `<<count\s+(\w+)>>` → AggregationNode (type "count")
3. Pour les blocs (foreach/if/group), parser récursivement les enfants jusqu'à la balise de fermeture

**Gestion du nesting** :
- Utiliser une **pile** (stack) pour suivre les blocs ouverts
- Lorsqu'une balise de fermeture est rencontrée, dépiler et rattacher les nœuds

### 3. Evaluator

**Responsabilité** : Résoudre les expressions dynamiques via reflection.

**Interface** :
```csharp
object Evaluate(string expression, object context);
```

**Implémentation** :
- `expression` : chemin de propriété (ex: "Address.City")
- `context` : objet racine (le modèle de données)
- Découper l'expression par `.`
- Résoudre chaque propriété via `Type.GetProperty()`
- **Cache** : `ConcurrentDictionary<string, PropertyInfo[]>` pour éviter de recompiler la même expression (thread-safe)

**Exemple** :
```csharp
// Expression: "Person.Address.City"
// Compilation: [Person_Property, Address_Property, City_Property]
// Évaluation: context.Person → .Address → .City → "Paris"
```

### 4. Renderer

**Responsabilité** : Transformer l'AST en contenu Excel.

**Interface** :
```csharp
void Render(Template template, RenderContext context, ExcelWorksheet worksheet);
```

**Stratégie de rendu** :

#### Expressions simples
- Remplacer la valeur de la cellule par le résultat de `Evaluate()`
- Conserver le style existant

#### Boucles (foreach)
1. Évaluer la collection
2. Si la collection est vide : supprimer le bloc entier (`DeleteRow`)
3. Si la collection a des éléments :
   - Rendre le premier élément dans les lignes existantes
   - Pour chaque élément supplémentaire :
     - `InsertRow` avec copie des styles (`copyStylesFromRow`)
     - Rendre les enfants dans la nouvelle ligne
   - Supprimer les balises de directive (`<<foreach>>`, `<</foreach>>`)

#### Conditions (if)
1. Évaluer l'expression booléenne
2. Si `true` : rendre les enfants, supprimer les balises
3. Si `false` : supprimer le bloc entier (`DeleteRow`)

**Gestion des offsets** :
- `InsertRow` et `DeleteRow` modifient les indices des lignes suivantes
- Le renderer maintient un **offset dynamique** (`rowOffset`) qui est ajusté après chaque opération

## Flux de données

```
Données C# (objet anonyme / POCO)
            ↓
     TemplateEngine.AddVariable(data)
     TemplateEngine.Generate()
            ↓
     ┌─────────────────┐
     │ 1. Parse()      │ → AST
     │ 2. Evaluate()   │ → Valeurs
     │ 3. Render()     │ → Excel modifié
     └─────────────────┘
            ↓
     Fichier .xlsx généré (via Save / SaveAs)
```

## Points d'extension

- **ITemplateParser** : Permet d'implémenter des syntaxes de template alternatives
- **IExpressionEvaluator** : Permet d'ajouter des fonctions personnalisées ou un langage d'expression différent
- **ITemplateRenderer** : Permet de changer la stratégie de rendu (ex: rendu en streaming)

## Décisions d'architecture

### Pourquoi un AST intermédiaire ?
- **Séparation des concerns** : Parser, Evaluator et Renderer sont indépendants
- **Testabilité** : On peut tester chaque phase isolément
- **Extensibilité** : Facile d'ajouter de nouveaux types de nœuds

### Pourquoi la reflection avec cache ?
- **Simplicité** : Pas de dépendance externe (pas de Roslyn, pas d'expression trees compilés)
- **Performance acceptable** : Le cache des `PropertyInfo[]` élimine le coût de la reflection après la première évaluation
- **Compatibilité** : Fonctionne sur .NET Framework 4.7+ sans complication

### Pourquoi xUnit pour les tests ?
- **Standard de l'industrie** pour .NET
- **Bon support dans .NET 8**
- **Tests paramétrés** faciles avec `[Theory]`

## Anti-patterns à éviter

- Ne pas stocker de `StyleID` en dur (ils changent lors des insertions)
- Ne pas itérer sur un worksheet en modifiant ses dimensions simultanément
- Éviter la reflection sans cache dans des boucles
- Ne pas mélanger la logique de parsing et de rendu
