using Microsoft.Vbe.Interop;

#nullable enable

namespace ExcelFusion
{
    /// <summary>
    /// Represents information about a VBA reference.
    /// </summary>
    internal class ReferenceInfo
    {
        /// <summary>
        /// Gets or sets the <see cref="Guid"/> of the reference.
        /// </summary>
        public System.Guid? Guid { get; set; }

        /// <summary>
        /// Gets or sets the name of the reference.
        /// </summary>
        public string? Name { get; set; }

        /// <summary>
        /// Gets or sets the full path that this reference refers to.
        /// </summary>
        public string FullPath { get; set; } = string.Empty;

        /// <summary>
        /// Gets or sets the version of the reference.
        /// </summary>
        public float? Version { get; set; }

        /// <summary>
        /// Gets or sets the type of the reference.
        /// </summary>
        public vbext_RefKind? Type { get; internal set; }
    }
}