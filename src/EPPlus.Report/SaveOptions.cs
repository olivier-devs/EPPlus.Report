#nullable enable

namespace EPPlus.Report;

/// <summary>
///     Specifies options that control how the rendered workbook is saved.
/// </summary>
public class SaveOptions
{
    /// <summary>
    ///     Gets or sets a value indicating whether formulas should be evaluated before saving the workbook.
    /// </summary>
    public bool EvaluateFormulasBeforeSave { get; set; } = false;

    /// <summary>
    ///     Gets or sets the password used to encrypt the workbook when saving to a file.
    ///     When <c>null</c> or empty, the workbook is saved without encryption.
    /// </summary>
    /// <remarks>
    ///     <para>Encryption is only supported when saving to a file (string path
    ///     or <see cref="System.IO.FileInfo"/>). Saving to a <see cref="System.IO.Stream"/>
    ///     with a non-empty password will throw a <see cref="System.NotSupportedException"/>.</para>
    ///     <para>The password is held in memory as plain text, matching EPPlus's own API
    ///     contract. Callers are responsible for managing the lifecycle and security of
    ///     the password.</para>
    /// </remarks>
    public string? Password { get; set; }
}