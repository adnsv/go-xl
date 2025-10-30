package xl

// Font represents font formatting properties for cell content.
// These properties correspond to the OpenXML font element as defined in ECMA-376.
type Font struct {
	Size          float64       // Font size in points (0 = use default of 11)
	Bold          bool          // Bold text
	Italic        bool          // Italic text
	Underline     UnderlineType // Underline style
	Strikethrough bool          // Strikethrough text
}

// UnderlineType represents the type of underline formatting.
type UnderlineType string

// Underline type constants as defined in ECMA-376 (ST_UnderlineValues).
const (
	UnderlineNone              UnderlineType = ""                    // No underline (default)
	UnderlineSingle            UnderlineType = "single"              // Single underline
	UnderlineDouble            UnderlineType = "double"              // Double underline
	UnderlineSingleAccounting  UnderlineType = "singleAccounting"   // Single accounting underline
	UnderlineDoubleAccounting  UnderlineType = "doubleAccounting"   // Double accounting underline
)

// IsDefault returns true if the font uses all default properties.
func (f *Font) IsDefault() bool {
	return f.Size == 0 && !f.Bold && !f.Italic &&
		f.Underline == UnderlineNone && !f.Strikethrough
}

// Empty returns true if the font has no custom properties set.
// This is an alias for IsDefault for consistency with other Empty() methods.
func (f *Font) Empty() bool {
	return f.IsDefault()
}
