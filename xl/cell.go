package xl

import "fmt"

// Cell represents a single cell in a worksheet.
// It contains the cell's value, type, formatting (XF), and position information.
type Cell struct {
	row          *Row
	columnNumber int // 1-based
	coord        string
	typ          CellType
	v            string
	picture      *PictureInfo

	XF
}

// PictureInfo contains image data and metadata for embedding images in cells.
// Supported formats are PNG and JPEG (specified via Extension field).
type PictureInfo struct {
	Extension string // File extension including dot (e.g., ".png", ".jpg", ".jpeg")
	Blob      []byte // Raw image data
}

// CellType is the type of cell value type.
type CellType int

// Cell value types enumeration.
const (
	CellTypeUnset CellType = iota
	CellTypeBool
	CellTypeDate
	CellTypeError
	CellTypeFormula
	CellTypeInlineString
	CellTypeNumber
	CellTypeSharedString

	// internal
	cellTypePicture
)

// XF (Extended Format) represents the complete formatting attributes for a cell.
// It includes alignment and font properties that define how the cell content appears.
type XF struct {
	Alignment Alignment
	Font      Font
}

// HorizontalAlignment represents the horizontal alignment of cell content.
type HorizontalAlignment string

// Horizontal alignment constants as defined in ECMA-376 (ST_HorizontalAlignment).
const (
	HAlignGeneral          HorizontalAlignment = "general"          // Default: numbers right-aligned, text left-aligned
	HAlignLeft             HorizontalAlignment = "left"             // Left aligned
	HAlignCenter           HorizontalAlignment = "center"           // Centered
	HAlignRight            HorizontalAlignment = "right"            // Right aligned
	HAlignFill             HorizontalAlignment = "fill"             // Fill/repeat content to fill column width
	HAlignJustify          HorizontalAlignment = "justify"          // Justified
	HAlignCenterContinuous HorizontalAlignment = "centerContinuous" // Center across selection
	HAlignDistributed      HorizontalAlignment = "distributed"      // Distributed alignment
)

// VerticalAlignment represents the vertical alignment of cell content.
type VerticalAlignment string

// Vertical alignment constants as defined in ECMA-376 (ST_VerticalAlignment).
const (
	VAlignTop         VerticalAlignment = "top"         // Top aligned
	VAlignCenter      VerticalAlignment = "center"      // Centered vertically
	VAlignBottom      VerticalAlignment = "bottom"      // Bottom aligned (default)
	VAlignJustify     VerticalAlignment = "justify"     // Justified
	VAlignDistributed VerticalAlignment = "distributed" // Distributed alignment
)

// Alignment represents the alignment properties for cell content.
// Both horizontal and vertical alignment can be set using type-safe constants.
type Alignment struct {
	Horizontal HorizontalAlignment
	Vertical   VerticalAlignment
}

// SetBool sets the cell value to a boolean.
// The value is stored as "1" (true) or "0" (false) in Excel format.
func (c *Cell) SetBool(v bool) {
	c.typ = CellTypeBool
	if v {
		c.v = "1"
	} else {
		c.v = "0"
	}
}

// SetInt sets the cell value to an integer number.
func (c *Cell) SetInt(v int64) {
	c.typ = CellTypeNumber
	c.v = fmt.Sprintf("%d", v)
}

// SetFloat sets the cell value to a floating-point number.
// The value is formatted using %g which chooses the most compact representation.
func (c *Cell) SetFloat(v float64) {
	c.typ = CellTypeNumber
	c.v = fmt.Sprintf("%g", v)
}

// SetStr sets the cell value to a string.
// The string will be stored in the shared string table for efficiency.
func (c *Cell) SetStr(v string) {
	c.typ = CellTypeSharedString
	c.v = v
}

// SetPicture sets the cell to display an image.
// The image data and extension must be provided via PictureInfo.
// Supported formats: PNG, JPEG.
func (c *Cell) SetPicture(p *PictureInfo) {
	c.typ = cellTypePicture
	c.picture = p
}

// Empty returns true if the alignment has no custom properties set.
// An empty alignment means both horizontal and vertical are using defaults.
func (a *Alignment) Empty() bool {
	return a.Horizontal == "" && a.Vertical == ""
}

// Empty returns true if the XF has no custom formatting properties set.
// This checks both alignment and font for default values.
func (xf *XF) Empty() bool {
	return xf.Alignment.Empty() && xf.Font.Empty()
}
