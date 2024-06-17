package xl

import "fmt"

type Cell struct {
	row          *Row
	columnNumber int // 1-based
	coord        string
	typ          CellType
	v            string
	picture      *PictureInfo
}

type PictureInfo struct {
	Extension string
	Blob      []byte
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

func (c *Cell) SetBool(v bool) {
	c.typ = CellTypeBool
	if v {
		c.v = "1"
	} else {
		c.v = "0"
	}
}

func (c *Cell) SetInt(v int64) {
	c.typ = CellTypeNumber
	c.v = fmt.Sprintf("%d", v)
}

func (c *Cell) SetFloat(v float64) {
	c.typ = CellTypeNumber

}

func (c *Cell) SetStr(v string) {
	c.typ = CellTypeSharedString
	c.v = v
}

func (c *Cell) SetPicture(p *PictureInfo) {
	c.typ = cellTypePicture
	c.picture = p
}
