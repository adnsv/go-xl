package xl

import "strconv"

// Row represents a single row in a worksheet.
// It contains cells and row-level formatting such as height.
type Row struct {
	Cells []*Cell

	Height float32 // Row height in points (0 = use default height)

	sheet            *Sheet
	rowNumber        int // 1-based
	nextColumnNumber int // 1-based, incremented as we add cells
}

// AddCell adds a new cell to the row and returns a pointer to it.
// Cells are added sequentially from left to right (A, B, C, etc.).
func (r *Row) AddCell() *Cell {
	c := &Cell{
		row:          r,
		columnNumber: r.nextColumnNumber,
		coord:        CellCoordAsString(r.nextColumnNumber, r.rowNumber),
	}
	r.nextColumnNumber++
	r.Cells = append(r.Cells, c)
	return c
}

// ColumnNumberAsLetters converts a 1-based column number to Excel column letters.
// For example: 1 -> "A", 26 -> "Z", 27 -> "AA", 702 -> "ZZ".
// Panics if n < 1.
func ColumnNumberAsLetters(n int) string {
	if n < 1 {
		panic("invalid column number")
	}
	var s string
	for n > 0 {
		s = string(rune((n-1)%26+65)) + s
		n = (n - 1) / 26
	}
	return s
}

// CellCoordAsString converts 1-based column and row numbers to an Excel cell reference.
// For example: (1, 1) -> "A1", (3, 5) -> "C5", (27, 10) -> "AA10".
// Panics if row < 0.
func CellCoordAsString(col, row int) string {
	if row < 0 {
		panic("invalid row number")
	}
	return ColumnNumberAsLetters(col) + strconv.Itoa(row)
}
