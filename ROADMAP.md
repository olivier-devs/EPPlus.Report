# EPPlus.Report - Roadmap / Plan de développement

## Phase 1 : Expressions simples (MVP)
**Objectif** : Injecter des valeurs simples dans des cellules Excel

**Fonctionnalités** :
- [x] Syntaxe `{{Property}}`
- [x] Accès aux propriétés imbriquées `{{Object.Property}}`
- [x] Évaluation via reflection avec cache
- [x] API publique `TemplateEngine`

**Tests** :
- [x] `Evaluate_SimpleProperty_ReturnsValue`
- [x] `Evaluate_NestedProperty_ReturnsValue`
- [x] `Render_SingleExpression_ReplacesValue`
- [x] `Render_SimpleTemplate_GeneratesOutput`

**Livrable** : Bibliothèque fonctionnelle pour l'injection de données simples

---

## Phase 2 : Boucles (foreach)
**Objectif** : Répliquer des lignes pour chaque élément d'une collection

**Fonctionnalités** :
- [ ] Syntaxe `<<foreach Items>> ... <</foreach>>`
- [ ] Duplication dynamique des lignes
- [ ] Copie des styles EPPlus (`InsertRow` avec `copyStylesFromRow`)
- [ ] Suppression du bloc si collection vide

**Tests** :
- [ ] `Parse_SimpleLoop_CreatesLoopNodeWithChildren`
- [ ] `Render_Loop_DuplicatesRows`
- [ ] `Render_Loop_EmptyCollection_RemovesBlock`

**Livrable** : Templates avec listes et tableaux dynamiques

---

## Phase 3 : Conditions (if)
**Objectif** : Afficher/masquer des blocs selon une condition

**Fonctionnalités** :
- [ ] Syntaxe `<<if Condition>> ... <</if>>`
- [ ] Évaluation booléenne (y compris nullables)
- [ ] Inclusion/exclusion de blocs

**Tests** :
- [ ] `Parse_SimpleIf_CreatesIfNodeWithChildren`
- [ ] `Render_IfTrue_ShowsContent`
- [ ] `Render_IfFalse_HidesContent`

**Livrable** : Templates conditionnels

---

## Phase 4 : Boucles imbriquées
**Objectif** : Boucles à l'intérieur de boucles (collections dans collections)

**Fonctionnalités** :
- [ ] Parser avec stack pour la gestion du nesting
- [ ] Renderer avec contexte imbriqué
- [ ] Gestion correcte des offsets de lignes

**Tests** :
- [ ] `Parse_NestedLoop_CreatesNestedStructure`
- [ ] `Render_NestedLoop_RendersCorrectly`
- [ ] `Render_TripleNested_HandlesCorrectly`

**Livrable** : Templates complexes avec hiérarchies

---

## Phase 5 : Agrégations
**Objectif** : Calculer des valeurs agrégées sur des collections

**Fonctionnalités** :
- [ ] Syntaxe `<<sum Property>>`
- [ ] Syntaxe `<<count Items>>`
- [ ] Accès au contexte parent pour les calculs

**Tests** :
- [ ] `Render_Sum_CalculatesTotal`
- [ ] `Render_Count_ReturnsItemCount`

**Livrable** : Tableaux avec totaux et statistiques

---

## Phase 6 : Groupement avancé (optionnel)
**Objectif** : Regrouper des données par catégorie

**Fonctionnalités** :
- [ ] Syntaxe `<<group Items by Category>>`
- [ ] En-têtes de groupe automatiques
- [ ] Agrégations par groupe

**Tests** :
- [ ] `Parse_Group_CreatesGroupNode`
- [ ] `Render_Group_CreatesSections`

**Livrable** : Rapports avec sections groupées

---

## Phase 7 : Optimisation et polish
**Objectif** : Performance et robustesse

**Tâches** :
- [ ] Benchmarks sur fichiers volumineux (10 000+ lignes)
- [ ] Gestion des erreurs et messages d'erreur clairs
- [ ] Documentation XML sur l'API publique
- [ ] Exemples d'utilisation dans le README

---

## Légende
- [x] Terminé
- [ ] À faire

## Notes
- Chaque phase doit être **validée** avant de passer à la suivante
- Les phases 1-5 constituent le **MVP complet**
- La phase 6 est **optionnelle** et peut être repoussée
- La phase 7 est **continue** et inclut la dette technique
