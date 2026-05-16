using System;

namespace EPPlus.Report.Model;

/// <summary>
///     Marks a property as visible and accessible in Excel templates.
///     When this attribute is applied to a property, it is automatically
///     added to the allowlist of the expression evaluator.
/// </summary>
/// <remarks>
///     <para>This attribute can be applied to properties in model classes to explicitly
///     declare which properties should be accessible in template expressions.</para>
///     <para>It supports inheritance - properties decorated on base classes are also recognized.</para>
///     <para>For nested objects, only the decorated properties at each level are accessible.</para>
/// </remarks>
[AttributeUsage(AttributeTargets.Property, Inherited = true, AllowMultiple = false)]
public sealed class TemplateVisibleAttribute : Attribute
{
}