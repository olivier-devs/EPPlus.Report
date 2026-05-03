using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using EPPlus.Report.Evaluation;
using EPPlus.Report.Model;
using OfficeOpenXml;
using OfficeOpenXml.ConditionalFormatting.Contracts;
using OfficeOpenXml.Style;

namespace EPPlus.Report.Rendering;

/// <summary>
///     Renders a parsed <see cref="Template" /> into an Excel worksheet using a data context and expression evaluator.
/// </summary>
public class TemplateRenderer : ITemplateRenderer
{
    private readonly IExpressionEvaluator _evaluator;
    private readonly TemplateErrors _renderingErrors;
    private readonly RowOperationTracker _tracker;
    private readonly TemplateErrors _warnings;

    /// <summary>
    ///     Initializes a new instance of the <see cref="TemplateRenderer" /> class.
    /// </summary>
    /// <param name="evaluator">The expression evaluator used to resolve template expressions.</param>
    /// <param name="renderingErrors">Optional collection to populate with rendering errors.</param>
    /// <param name="tracker">Optional tracker for row insertions and deletions.</param>
    /// <param name="warnings">Optional collection to populate with warnings.</param>
    public TemplateRenderer(IExpressionEvaluator evaluator, TemplateErrors renderingErrors = null,
        RowOperationTracker tracker = null, TemplateErrors warnings = null)
    {
        _evaluator = evaluator;
        _renderingErrors = renderingErrors;
        _tracker = tracker;
        _warnings = warnings;
    }

    /// <summary>
    ///     Renders the specified template into the worksheet using the provided context.
    /// </summary>
    /// <param name="template">The parsed template to render.</param>
    /// <param name="context">The data context for evaluating expressions.</param>
    /// <param name="worksheet">The target Excel worksheet.</param>
    public void Render(Template template, RenderContext context, ExcelWorksheet worksheet)
    {
        RenderNodes(template.Nodes, context, worksheet, 0);
    }

    // ... rest of the file remains unchanged, only Render method is public API
    private int RenderNodes(List<TemplateNode> nodes, RenderContext context, ExcelWorksheet worksheet, int rowOffset)
    {
        foreach (var node in nodes)
        {
            rowOffset = RenderNode(node, context, worksheet, rowOffset);
        }

        return rowOffset;
    }

    private int RenderNode(TemplateNode node, RenderContext context, ExcelWorksheet worksheet, int rowOffset)
    {
        switch (node)
        {
            case ExpressionNode exprNode:
            {
                object value = null;
                var targetRow = exprNode.Row + rowOffset;
                var targetCol = exprNode.Column;
                var cellAddress = worksheet.Cells[targetRow, targetCol].Address;

                if (context.Variables != null &&
                    context.Variables.TryGetValue(exprNode.ExpressionPath, out var varValue))
                {
                    value = varValue;
                }
                else if (context.IsNamedRangeLoop)
                {
                    value = ResolveNamedRangeExpression(exprNode.ExpressionPath, context, worksheet, targetRow,
                        targetCol, cellAddress);
                }

                if (value == null)
                {
                    value = TryEvaluate(exprNode.ExpressionPath, context.Current, worksheet, targetRow, targetCol,
                        cellAddress, exprNode.FunctionName);
                }
                else if (!string.IsNullOrEmpty(exprNode.FunctionName))
                {
                    value = ApplyFunction(exprNode.FunctionName, value, worksheet, targetRow, targetCol, cellAddress);
                }

                worksheet.Cells[targetRow, targetCol].Value = value;
                return rowOffset;
            }

            case GroupNode groupNode:
                return RenderGroup(groupNode, context, worksheet, rowOffset);

            case LoopNode loopNode:
                return RenderLoop(loopNode, context, worksheet, rowOffset);

            case IfNode ifNode:
                return RenderIf(ifNode, context, worksheet, rowOffset);

            case AggregationNode aggNode:
                return RenderAggregation(aggNode, context, worksheet, rowOffset);

            default:
                return rowOffset;
        }
    }

    private int RenderLoop(LoopNode loopNode, RenderContext context, ExcelWorksheet worksheet, int rowOffset)
    {
        if (loopNode is NamedRangeLoopNode namedRangeLoop && namedRangeLoop.GroupByDefinitions.Count > 0)
        {
            return RenderNamedRangeGroup(namedRangeLoop, context, worksheet, rowOffset);
        }

        if (loopNode is NamedRangeLoopNode namedRangeLoopNoGroup)
        {
            return RenderNamedRangeLoop(namedRangeLoopNoGroup, context, worksheet, rowOffset);
        }

        IEnumerable collection = null;
        var blockStartRow = loopNode.Row + rowOffset;
        var blockStartCol = loopNode.Column;
        var blockStartAddress = worksheet.Cells[blockStartRow, blockStartCol].Address;

        if (context.Variables != null && context.Variables.TryGetValue(loopNode.CollectionName, out var varValue))
        {
            collection = varValue as IEnumerable;
        }
        else if (context.Current != null)
        {
            collection = TryEvaluate(loopNode.CollectionName, context.Current, worksheet, blockStartRow, blockStartCol,
                blockStartAddress) as IEnumerable;
        }

        var items = collection?.Cast<object>().ToList() ?? [];

        var blockEndRow = loopNode.EndRow + rowOffset;
        var contentHeight = loopNode.EndRow - loopNode.Row - 1;

        if (items.Count == 0)
        {
            _tracker?.RecordDelete(worksheet, blockStartRow, loopNode.EndRow - loopNode.Row + 1);
            worksheet.DeleteRow(blockStartRow, loopNode.EndRow - loopNode.Row + 1);
            return rowOffset - (loopNode.EndRow - loopNode.Row + 1);
        }

        // Remove start tag
        worksheet.Cells[blockStartRow, loopNode.Column].Value = null;

        // Remove end tag
        worksheet.Cells[blockEndRow, loopNode.Column].Value = null;

        var currentOffset = rowOffset;

        for (var i = 0; i < items.Count; i++)
        {
            var itemContext = new RenderContext
            {
                Current = items[i],
                Variables = context.Variables != null ? new Dictionary<string, object>(context.Variables) : null,
                CurrentCollection = collection,
                IsNamedRangeLoop = false,
                CurrentIndex = i
            };

            if (i == 0)
            {
                // Render first item in existing rows
                currentOffset = RenderNodes(loopNode.Children, itemContext, worksheet, currentOffset);
            }
            else
            {
                // Insert block of rows for this item
                var insertAtRow = blockEndRow + currentOffset - rowOffset;
                if (contentHeight > 0)
                {
                    _tracker?.RecordInsert(worksheet, insertAtRow, contentHeight);
                    worksheet.InsertRow(insertAtRow, contentHeight, blockStartRow + 1 + currentOffset - rowOffset);
                }

                // Render children in the new block
                var childOffset = insertAtRow - loopNode.Row - 1;
                currentOffset = RenderNodes(loopNode.Children, itemContext, worksheet, childOffset);
            }
        }

        ReconcileConditionalFormatting(loopNode.ConditionalFormattingRules, loopNode.Row + rowOffset,
            loopNode.EndRow + rowOffset, loopNode.Row + rowOffset, loopNode.EndRow + currentOffset, worksheet);
        return currentOffset;
    }

    private int RenderNamedRangeLoop(NamedRangeLoopNode loopNode, RenderContext context, ExcelWorksheet worksheet,
        int rowOffset)
    {
        IEnumerable collection = null;
        var blockStartRow = loopNode.Row + rowOffset;
        var blockStartCol = loopNode.Column;
        var blockStartAddress = worksheet.Cells[blockStartRow, blockStartCol].Address;

        if (context.Variables != null && context.Variables.TryGetValue(loopNode.CollectionName, out var varValue))
        {
            collection = varValue as IEnumerable;
        }
        else if (context.Current != null)
        {
            collection = TryEvaluate(loopNode.CollectionName, context.Current, worksheet, blockStartRow, blockStartCol,
                blockStartAddress) as IEnumerable;
        }

        var items = collection?.Cast<object>().ToList() ?? [];

        var blockEndRow = loopNode.EndRow + rowOffset;
        var headerRowCount = loopNode.HeaderRowCount;
        var serviceRowCount = loopNode.ServiceRowCount;
        var dataStartRow = loopNode.Row + headerRowCount;
        var dataEndRow = loopNode.EndRow - serviceRowCount;
        var dataRowCount = dataEndRow - dataStartRow + 1;
        if (dataRowCount < 0)
        {
            dataRowCount = 0;
        }

        if (items.Count == 0)
        {
            _tracker?.RecordDelete(worksheet, blockStartRow, loopNode.EndRow - loopNode.Row + 1);
            worksheet.DeleteRow(blockStartRow, loopNode.EndRow - loopNode.Row + 1);
            return rowOffset - (loopNode.EndRow - loopNode.Row + 1);
        }

        // Separate children into header and data
        var headerChildren = new List<TemplateNode>();
        var dataChildren = new List<TemplateNode>();
        foreach (var child in loopNode.Children)
        {
            if (child.Row < dataStartRow)
            {
                headerChildren.Add(child);
            }
            else if (child.Row <= dataEndRow)
            {
                dataChildren.Add(child);
            }
        }

        // Render header once
        RenderNodes(headerChildren, context, worksheet, rowOffset);

        // Render data rows for each item
        var currentOffset = rowOffset;
        for (var i = 0; i < items.Count; i++)
        {
            var itemContext = new RenderContext
            {
                Current = items[i],
                Variables = context.Variables != null ? new Dictionary<string, object>(context.Variables) : null,
                CurrentCollection = collection,
                IsNamedRangeLoop = true,
                CurrentIndex = i
            };

            if (i == 0)
            {
                // Render first item in existing data rows
                currentOffset = RenderNodes(dataChildren, itemContext, worksheet, currentOffset);
            }
            else
            {
                // Insert data rows for this item before the service row
                var insertAtRow = dataEndRow + currentOffset + 1;
                if (dataRowCount > 0)
                {
                    var copyStylesFromRow = dataStartRow + currentOffset;
                    _tracker?.RecordInsert(worksheet, insertAtRow, dataRowCount);
                    worksheet.InsertRow(insertAtRow, dataRowCount, copyStylesFromRow);
                }

                var childOffset = insertAtRow - dataStartRow;
                currentOffset = RenderNodes(dataChildren, itemContext, worksheet, childOffset);
            }
        }

        // Process service tags
        if (loopNode.ServiceTags.Count > 0)
        {
            var serviceRow = loopNode.EndRow + currentOffset;
            var firstDataRow = loopNode.Row + loopNode.HeaderRowCount + rowOffset;
            var lastDataRow = serviceRow - 1;

            foreach (var tag in loopNode.ServiceTags)
            {
                ApplyServiceTag(tag, firstDataRow, lastDataRow, worksheet, serviceRow);
            }
        }

        ReconcileConditionalFormatting(loopNode.ConditionalFormattingRules, loopNode.Row + rowOffset,
            loopNode.EndRow + rowOffset, loopNode.Row + rowOffset, loopNode.EndRow + currentOffset, worksheet);
        return currentOffset;
    }

    private int RenderNamedRangeGroup(NamedRangeLoopNode loopNode, RenderContext context, ExcelWorksheet worksheet,
        int rowOffset)
    {
        IEnumerable collection = null;
        var blockStartRow = loopNode.Row + rowOffset;
        var blockStartCol = loopNode.Column;
        var blockStartAddress = worksheet.Cells[blockStartRow, blockStartCol].Address;

        if (context.Variables != null && context.Variables.TryGetValue(loopNode.CollectionName, out var varValue))
        {
            collection = varValue as IEnumerable;
        }
        else if (context.Current != null)
        {
            collection = TryEvaluate(loopNode.CollectionName, context.Current, worksheet, blockStartRow, blockStartCol,
                blockStartAddress) as IEnumerable;
        }

        var items = collection?.Cast<object>().ToList() ?? [];

        var blockEndRow = loopNode.EndRow + rowOffset;
        var headerRowCount = loopNode.HeaderRowCount;
        var serviceRowCount = loopNode.ServiceRowCount;
        var dataStartRow = loopNode.Row + headerRowCount;
        var dataEndRow = loopNode.EndRow - serviceRowCount;
        var dataRowCount = dataEndRow - dataStartRow + 1;
        if (dataRowCount < 0)
        {
            dataRowCount = 0;
        }

        if (items.Count == 0)
        {
            _tracker?.RecordDelete(worksheet, blockStartRow, loopNode.EndRow - loopNode.Row + 1);
            worksheet.DeleteRow(blockStartRow, loopNode.EndRow - loopNode.Row + 1);
            return rowOffset - (loopNode.EndRow - loopNode.Row + 1);
        }

        // Separate children into header and data
        var headerChildren = new List<TemplateNode>();
        var dataChildren = new List<TemplateNode>();
        foreach (var child in loopNode.Children)
        {
            if (child.Row < dataStartRow)
            {
                headerChildren.Add(child);
            }
            else if (child.Row <= dataEndRow)
            {
                dataChildren.Add(child);
            }
        }

        // Render header once
        RenderNodes(headerChildren, context, worksheet, rowOffset);

        // Sort and group items using the first GroupByDefinition
        var groupDef = loopNode.GroupByDefinitions[0];
        var groupPaths = new List<string> { groupDef.PropertyPath };
        var groups = GroupRenderer.SortAndGroup(items, groupPaths, _evaluator, groupDef.Descending);

        var currentOffset = rowOffset;
        var firstDataRendered = false;

        for (var g = 0; g < groups.Count; g++)
        {
            var group = groups[g];
            var groupFirstDataRow = 0;

            for (var i = 0; i < group.Items.Count; i++)
            {
                var item = group.Items[i];
                var itemContext = new RenderContext
                {
                    Current = item,
                    Variables = context.Variables != null ? new Dictionary<string, object>(context.Variables) : null,
                    CurrentCollection = collection,
                    IsNamedRangeLoop = true,
                    CurrentIndex = i
                };

                if (!firstDataRendered)
                {
                    groupFirstDataRow = dataStartRow + currentOffset;
                    currentOffset = RenderNodes(dataChildren, itemContext, worksheet, currentOffset);
                    firstDataRendered = true;
                }
                else
                {
                    var insertAtRow = dataEndRow + currentOffset + 1;
                    if (i == 0)
                    {
                        groupFirstDataRow = insertAtRow;
                    }

                    if (dataRowCount > 0)
                    {
                        var copyStylesFromRow = dataStartRow + currentOffset;
                        _tracker?.RecordInsert(worksheet, insertAtRow, dataRowCount);
                        worksheet.InsertRow(insertAtRow, dataRowCount, copyStylesFromRow);
                    }

                    var childOffset = insertAtRow - dataStartRow;
                    currentOffset = RenderNodes(dataChildren, itemContext, worksheet, childOffset);
                }
            }

            // Insert subtotal row
            if (!loopNode.RangeGroupOptions.DisableSubtotals && loopNode.ServiceTags.Count > 0)
            {
                var insertAtRow = dataEndRow + currentOffset + 1;
                var copyStylesFromRow = loopNode.EndRow + currentOffset;
                _tracker?.RecordInsert(worksheet, insertAtRow, 1);
                worksheet.InsertRow(insertAtRow, 1, copyStylesFromRow);

                var lastDataRow = insertAtRow - 1;

                foreach (var tag in loopNode.ServiceTags)
                {
                    ApplyServiceTag(tag, groupFirstDataRow, lastDataRow, worksheet, insertAtRow);
                }

                currentOffset++;
            }
        }

        // Remove original service row
        if (serviceRowCount > 0)
        {
            _tracker?.RecordDelete(worksheet, loopNode.EndRow + currentOffset, serviceRowCount);
            worksheet.DeleteRow(loopNode.EndRow + currentOffset, serviceRowCount);
            currentOffset -= serviceRowCount;
        }

        ReconcileConditionalFormatting(loopNode.ConditionalFormattingRules, loopNode.Row + rowOffset,
            loopNode.EndRow + rowOffset, loopNode.Row + rowOffset, loopNode.EndRow + currentOffset, worksheet);
        return currentOffset;
    }

    private int RenderAggregation(AggregationNode aggNode, RenderContext context, ExcelWorksheet worksheet,
        int rowOffset)
    {
        object result = null;

        if (aggNode.AggregationType == "sum")
        {
            var collection = context.CurrentCollection;
            if (collection != null)
            {
                decimal sum = 0;
                foreach (var item in collection)
                {
                    var val = TryEvaluate(aggNode.PropertyName, item, worksheet, aggNode.Row + rowOffset,
                        aggNode.Column, worksheet.Cells[aggNode.Row + rowOffset, aggNode.Column].Address);
                    if (val != null)
                    {
                        sum += Convert.ToDecimal(val);
                    }
                }

                result = sum;
            }
        }
        else if (aggNode.AggregationType == "count")
        {
            var collection = context.CurrentCollection;
            if (collection != null)
            {
                var count = 0;
                foreach (var _ in collection)
                {
                    count++;
                }

                result = count;
            }
            else
            {
                var val = TryEvaluate(aggNode.PropertyName, context.Current, worksheet, aggNode.Row + rowOffset,
                    aggNode.Column, worksheet.Cells[aggNode.Row + rowOffset, aggNode.Column].Address);
                if (val is IEnumerable enumerable && !(val is string))
                {
                    var count = 0;
                    foreach (var _ in enumerable)
                    {
                        count++;
                    }

                    result = count;
                }
                else
                {
                    result = val;
                }
            }
        }

        worksheet.Cells[aggNode.Row + rowOffset, aggNode.Column].Value = result ?? "AGG";
        return rowOffset;
    }

    private int RenderIf(IfNode ifNode, RenderContext context, ExcelWorksheet worksheet, int rowOffset)
    {
        var conditionRow = ifNode.Row + rowOffset;
        var conditionCol = ifNode.Column;
        var conditionAddress = worksheet.Cells[conditionRow, conditionCol].Address;
        var conditionValue = TryEvaluate(ifNode.ConditionExpression, context.Current, worksheet, conditionRow,
            conditionCol, conditionAddress);
        var isTrue = conditionValue is bool and true;

        var blockStartRow = ifNode.Row + rowOffset;
        var blockEndRow = ifNode.EndRow + rowOffset;

        if (isTrue)
        {
            // Remove start tag
            worksheet.Cells[blockStartRow, ifNode.Column].Value = null;

            // Render children
            var childOffset = RenderNodes(ifNode.Children, context, worksheet, rowOffset);

            // Remove end tag (adjust for any insertions/deletions by children)
            worksheet.Cells[ifNode.EndRow + childOffset, ifNode.Column].Value = null;

            ReconcileConditionalFormatting(ifNode.ConditionalFormattingRules, ifNode.Row + rowOffset,
                ifNode.EndRow + rowOffset, ifNode.Row + rowOffset, ifNode.EndRow + childOffset, worksheet);
            return childOffset;
        }

        // Remove the entire block including tags
        _tracker?.RecordDelete(worksheet, blockStartRow, ifNode.EndRow - ifNode.Row + 1);
        worksheet.DeleteRow(blockStartRow, ifNode.EndRow - ifNode.Row + 1);
        ReconcileConditionalFormatting(ifNode.ConditionalFormattingRules, ifNode.Row, ifNode.EndRow, 0, -1, worksheet);
        return rowOffset - (ifNode.EndRow - ifNode.Row + 1);
    }

    private int RenderGroup(GroupNode groupNode, RenderContext context, ExcelWorksheet worksheet, int rowOffset)
    {
        IEnumerable collection = null;
        var blockStartRow = groupNode.Row + rowOffset;
        var blockStartCol = groupNode.Column;
        var blockStartAddress = worksheet.Cells[blockStartRow, blockStartCol].Address;

        if (context.Variables != null && context.Variables.TryGetValue(groupNode.CollectionName, out var varValue))
        {
            collection = varValue as IEnumerable;
        }
        else if (context.Current != null)
        {
            collection = TryEvaluate(groupNode.CollectionName, context.Current, worksheet, blockStartRow, blockStartCol,
                blockStartAddress) as IEnumerable;
        }

        var items = collection?.Cast<object>().ToList() ?? [];

        var blockEndRow = groupNode.EndRow + rowOffset;

        if (items.Count == 0)
        {
            _tracker?.RecordDelete(worksheet, blockStartRow, groupNode.EndRow - groupNode.Row + 1);
            worksheet.DeleteRow(blockStartRow, groupNode.EndRow - groupNode.Row + 1);
            return rowOffset - (groupNode.EndRow - groupNode.Row + 1);
        }

        // Remove start and end tags
        worksheet.Cells[blockStartRow, groupNode.Column].Value = null;
        worksheet.Cells[blockEndRow, groupNode.Column].Value = null;

        // Sort and group
        var groups =
            GroupRenderer.SortAndGroup(items, groupNode.GroupByPaths, _evaluator, groupNode.Options.Descending);

        var hasSubtotal = groupNode.SubtotalTemplate.Count > 0;

        // Delete original subtotal template rows - they serve as templates but should not appear in output
        if (hasSubtotal)
        {
            var firstSubtotalRow = groupNode.SubtotalTemplate.Min(n => n.Row) + rowOffset;
            var lastSubtotalRow = groupNode.SubtotalTemplate.Max(n => n.Row) + rowOffset;
            var subtotalRowCount = lastSubtotalRow - firstSubtotalRow + 1;
            _tracker?.RecordDelete(worksheet, firstSubtotalRow, subtotalRowCount);
            worksheet.DeleteRow(firstSubtotalRow, subtotalRowCount);
            blockEndRow -= subtotalRowCount;
        }

        var currentOffset = rowOffset;
        var dataTemplateHeight = blockEndRow - blockStartRow - 1;

        for (var g = 0; g < groups.Count; g++)
        {
            var group = groups[g];

            int groupFirstDataRow;
            if (g == 0)
            {
                groupFirstDataRow = blockStartRow + 1;
            }
            else
            {
                var insertAtRow = blockEndRow + currentOffset - rowOffset;
                groupFirstDataRow = insertAtRow;
            }

            for (var i = 0; i < group.Items.Count; i++)
            {
                var item = group.Items[i];
                var itemContext = new RenderContext
                {
                    Current = item,
                    Variables = context.Variables != null ? new Dictionary<string, object>(context.Variables) : null,
                    CurrentCollection = collection,
                    IsNamedRangeLoop = false,
                    CurrentIndex = i
                };

                if (g == 0 && i == 0)
                {
                    currentOffset = RenderNodes(groupNode.Children, itemContext, worksheet, currentOffset);
                }
                else
                {
                    var insertAtRow = blockEndRow + currentOffset - rowOffset;
                    if (dataTemplateHeight > 0)
                    {
                        _tracker?.RecordInsert(worksheet, insertAtRow, dataTemplateHeight);
                        worksheet.InsertRow(insertAtRow, dataTemplateHeight,
                            blockStartRow + 1 + currentOffset - rowOffset);
                    }

                    var childOffset = insertAtRow - groupNode.Row - 1;
                    currentOffset = RenderNodes(groupNode.Children, itemContext, worksheet, childOffset);
                }
            }

            var groupLastDataRow = groupFirstDataRow + group.Items.Count * dataTemplateHeight - 1;

            // MergeLabels support
            if (groupNode.Options.MergeLabels != MergeMode.None && groupNode.GroupByPaths.Count > 0
                                                                && groupLastDataRow > groupFirstDataRow)
            {
                var groupKeyCol = groupNode.Children
                                      .OfType<ExpressionNode>()
                                      .FirstOrDefault(e => e.ExpressionPath == groupNode.GroupByPaths[0])?.Column
                                  ?? groupNode.Column;

                var range = worksheet.Cells[groupFirstDataRow, groupKeyCol, groupLastDataRow, groupKeyCol];
                range.Merge = true;
                range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                if (groupNode.Options.MergeLabels is MergeMode.Merge1 or MergeMode.Merge2)
                {
                    for (var r = groupFirstDataRow + 1; r <= groupLastDataRow; r++)
                    {
                        worksheet.Cells[r, groupKeyCol].Value = null;
                    }
                }
            }

            // Render subtotal for this group (unless disabled)
            if (hasSubtotal && !groupNode.Options.DisableSubtotals)
            {
                var subtotalContext = new RenderContext
                {
                    Current = group.Items.Last(),
                    Variables = context.Variables != null ? new Dictionary<string, object>(context.Variables) : null,
                    CurrentCollection = group.Items,
                    IsNamedRangeLoop = false,
                    CurrentIndex = group.Items.Count - 1
                };

                var subtotalTemplateRow = groupNode.SubtotalTemplate[0].Row;
                var insertAtRow = blockEndRow + currentOffset - rowOffset;
                _tracker?.RecordInsert(worksheet, insertAtRow, 1);
                worksheet.InsertRow(insertAtRow, 1, blockStartRow + 1 + currentOffset - rowOffset);
                var subtotalOffset = insertAtRow - subtotalTemplateRow;
                currentOffset = RenderNodes(groupNode.SubtotalTemplate, subtotalContext, worksheet, subtotalOffset);
                currentOffset++;
            }
        }

        // Grand total
        if (!groupNode.Options.DisableGrandTotal && hasSubtotal)
        {
            var grandTotalContext = new RenderContext
            {
                Current = items.Last(),
                Variables = context.Variables != null ? new Dictionary<string, object>(context.Variables) : null,
                CurrentCollection = items,
                IsNamedRangeLoop = false,
                CurrentIndex = items.Count - 1
            };

            var subtotalTemplateRow = groupNode.SubtotalTemplate[0].Row;
            var insertAtRow = blockEndRow + currentOffset - rowOffset;
            _tracker?.RecordInsert(worksheet, insertAtRow, 1);
            worksheet.InsertRow(insertAtRow, 1, blockStartRow + 1 + currentOffset - rowOffset);
            var grandOffset = insertAtRow - subtotalTemplateRow;
            currentOffset = RenderNodes(groupNode.SubtotalTemplate, grandTotalContext, worksheet, grandOffset);
            currentOffset++;
        }

        ReconcileConditionalFormatting(groupNode.ConditionalFormattingRules, groupNode.Row + rowOffset,
            groupNode.EndRow + rowOffset, groupNode.Row + rowOffset, groupNode.EndRow + currentOffset, worksheet);
        return currentOffset;
    }

    private object ResolveNamedRangeExpression(string expression, RenderContext context, ExcelWorksheet worksheet,
        int row, int column, string cellAddress)
    {
        if (expression == "item")
        {
            return context.Current;
        }

        if (expression == "index")
        {
            return context.CurrentIndex;
        }

        if (expression == "items")
        {
            return context.CurrentCollection;
        }

        if (expression.StartsWith("item."))
        {
            var propertyPath = expression.Substring(5);
            return TryEvaluate(propertyPath, context.Current, worksheet, row, column, cellAddress);
        }

        return null;
    }

    private static void ReconcileConditionalFormatting(List<ConditionalFormattingRule> rules, int originalStartRow,
        int originalEndRow, int finalStartRow, int finalEndRow, ExcelWorksheet worksheet)
    {
        if (rules == null || rules.Count == 0)
        {
            return;
        }

        // Remove old rules that intersected the original block range
        var rulesToRemove = new List<IExcelConditionalFormattingRule>();
        foreach (var cf in worksheet.ConditionalFormatting)
        {
            if (cf.Address.Start.Row <= originalEndRow && cf.Address.End.Row >= originalStartRow)
            {
                rulesToRemove.Add(cf);
            }
        }

        foreach (var cf in rulesToRemove)
        {
            worksheet.ConditionalFormatting.Remove(cf);
        }

        if (finalEndRow < finalStartRow)
        {
            return; // block was deleted
        }

        // Re-apply rules to the final range
        foreach (var rule in rules)
        {
            var newAddress = worksheet.Cells[finalStartRow, 1, finalEndRow, worksheet.Dimension?.End.Column ?? 1]
                .Address;
            var cf = worksheet.ConditionalFormatting.AddExpression(newAddress);
            cf.Formula = rule.Formula;
            cf.Style.Fill.BackgroundColor.Color = Color.Red; // placeholder
            cf.Priority = rule.Priority;
            cf.StopIfTrue = rule.StopIfTrue;
        }
    }

    private void ApplyServiceTag(ServiceTag tag, int firstDataRow, int lastDataRow, ExcelWorksheet worksheet, int row)
    {
        var cell = worksheet.Cells[row, tag.Column];
        var rangeAddress = new ExcelAddress(firstDataRow, tag.Column, lastDataRow, tag.Column).Address;

        switch (tag.TagName)
        {
            case "sum":
                cell.Formula = $"SUBTOTAL(9,{rangeAddress})";
                break;
            case "count":
                cell.Formula = $"SUBTOTAL(3,{rangeAddress})";
                break;
            case "avg":
                cell.Formula = $"SUBTOTAL(1,{rangeAddress})";
                break;
        }
    }

    private object TryEvaluate(string expression, object context, ExcelWorksheet worksheet, int row, int column,
        string cellAddress, string functionName = null)
    {
        try
        {
            if (string.IsNullOrEmpty(functionName))
            {
                return _evaluator.Evaluate(expression, context);
            }

            if (_evaluator is ExpressionEvaluator ee)
            {
                return ee.Evaluate(expression, context, functionName);
            }

            throw new InvalidOperationException("Evaluator does not support functions.");
        }
        catch (PropertyNotFoundException ex)
        {
            if (_warnings != null)
            {
                _warnings.Add(new TemplateError
                {
                    Message = ex.Message,
                    Type = ErrorType.Warning,
                    WorksheetName = worksheet.Name,
                    Row = row,
                    Column = column,
                    CellAddress = cellAddress,
                    Expression = expression
                });
                return null;
            }

            if (_renderingErrors != null)
            {
                _renderingErrors.Add(new TemplateError
                {
                    Message = ex.Message,
                    Type = ErrorType.Evaluation,
                    WorksheetName = worksheet.Name,
                    Row = row,
                    Column = column,
                    CellAddress = cellAddress,
                    Expression = expression
                });
                return null;
            }

            throw;
        }
        catch (Exception ex) when (ex is ArgumentException || ex is NullReferenceException ||
                                   ex is InvalidOperationException)
        {
            if (_renderingErrors == null)
            {
                throw;
            }

            var message = ex is NullReferenceException
                ? $"Null reference evaluating '{expression}': {ex.Message}"
                : ex.Message;

            _renderingErrors.Add(new TemplateError
            {
                Message = message,
                Type = ErrorType.Evaluation,
                WorksheetName = worksheet.Name,
                Row = row,
                Column = column,
                CellAddress = cellAddress,
                Expression = expression
            });
            return null;
        }
    }

    private object ApplyFunction(string functionName, object value, ExcelWorksheet worksheet, int row, int column,
        string cellAddress)
    {
        try
        {
            if (_evaluator is ExpressionEvaluator ee)
            {
                return ee.ApplyFunction(functionName, value);
            }

            throw new InvalidOperationException("Evaluator does not support functions.");
        }
        catch (Exception ex) when (ex is ArgumentException or NullReferenceException or InvalidOperationException)
        {
            if (_renderingErrors == null)
            {
                throw;
            }

            _renderingErrors.Add(new TemplateError
            {
                Message = ex.Message,
                Type = ErrorType.Evaluation,
                WorksheetName = worksheet.Name,
                Row = row,
                Column = column,
                CellAddress = cellAddress,
                Expression = functionName
            });
            return null;
        }
    }
}