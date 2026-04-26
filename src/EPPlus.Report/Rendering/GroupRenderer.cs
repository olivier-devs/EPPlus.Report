using System;
using System.Collections.Generic;
using System.Linq;
using EPPlus.Report.Evaluation;

namespace EPPlus.Report.Rendering;

/// <summary>
///     Provides utilities for sorting and grouping data items during template rendering.
/// </summary>
public static class GroupRenderer
{
    /// <summary>
    ///     Sorts the items by the specified group-by paths and groups consecutive items with equal keys.
    /// </summary>
    /// <param name="items">The items to sort and group.</param>
    /// <param name="groupByPaths">The property paths used to build group keys.</param>
    /// <param name="evaluator">The expression evaluator used to resolve property values.</param>
    /// <param name="descending">Whether to sort in descending order.</param>
    /// <returns>A list of <see cref="GroupResult" /> representing the grouped items.</returns>
    public static List<GroupResult> SortAndGroup(IEnumerable<object> items, List<string> groupByPaths,
        IExpressionEvaluator evaluator, bool descending = false)
    {
        var itemList = items.ToList();
        if (itemList.Count == 0)
        {
            return [];
        }

        var sorted = descending
            ? itemList.OrderByDescending(item => item, new GroupKeyComparer(groupByPaths, evaluator)).ToList()
            : itemList.OrderBy(item => item, new GroupKeyComparer(groupByPaths, evaluator)).ToList();

        var groups = new List<GroupResult>();
        GroupResult currentGroup = null;

        foreach (var item in sorted)
        {
            var key = BuildKey(item, groupByPaths, evaluator);
            if (currentGroup == null || !KeysEqual(currentGroup.Key, key))
            {
                currentGroup = new GroupResult { Key = key, Items = new List<object>() };
                groups.Add(currentGroup);
            }

            currentGroup.Items.Add(item);
        }

        return groups;
    }

    private static List<object> BuildKey(object item, List<string> groupByPaths, IExpressionEvaluator evaluator)
    {
        return groupByPaths.Select(path => evaluator.Evaluate(path, item)).ToList();
    }

    private static bool KeysEqual(List<object> a, List<object> b)
    {
        if (a.Count != b.Count)
        {
            return false;
        }

        for (var i = 0; i < a.Count; i++)
        {
            if (!Equals(a[i], b[i]))
            {
                return false;
            }
        }

        return true;
    }

    private class GroupKeyComparer : IComparer<object>
    {
        private readonly IExpressionEvaluator _evaluator;
        private readonly List<string> _paths;

        public GroupKeyComparer(List<string> paths, IExpressionEvaluator evaluator)
        {
            _paths = paths;
            _evaluator = evaluator;
        }

        public int Compare(object x, object y)
        {
            foreach (var path in _paths)
            {
                var valX = _evaluator.Evaluate(path, x);
                var valY = _evaluator.Evaluate(path, y);
                var cmp = CompareValues(valX, valY);
                if (cmp != 0)
                {
                    return cmp;
                }
            }

            return 0;
        }

        private static int CompareValues(object a, object b)
        {
            if (a == null && b == null)
            {
                return 0;
            }

            if (a == null)
            {
                return -1;
            }

            if (b == null)
            {
                return 1;
            }

            if (a is IComparable comparableA)
            {
                return comparableA.CompareTo(b);
            }

            return string.Compare(a.ToString(), b.ToString(), StringComparison.Ordinal);
        }
    }
}