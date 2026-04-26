using System.Collections;
using System.Collections.Generic;

namespace EPPlus.Report.Model;

/// <summary>
///     Provides the data context used during template rendering, including the current item and variables.
/// </summary>
public class RenderContext
{
    /// <summary>
    ///     Gets or sets the current data object being rendered.
    /// </summary>
    public object Current { get; set; }

    /// <summary>
    ///     Gets or sets the named variables available for expression evaluation.
    /// </summary>
    public Dictionary<string, object> Variables { get; set; }

    /// <summary>
    ///     Gets or sets the current collection being iterated over.
    /// </summary>
    public IEnumerable CurrentCollection { get; set; }

    /// <summary>
    ///     Gets or sets a value indicating whether the current rendering context is inside a named range loop.
    /// </summary>
    public bool IsNamedRangeLoop { get; set; }

    /// <summary>
    ///     Gets or sets the zero-based index of the current item within the collection.
    /// </summary>
    public int CurrentIndex { get; set; }
}