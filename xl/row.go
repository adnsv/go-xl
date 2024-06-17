package xl

import "strconv"

type Row struct {
	Cells []*Cell

	Height float32 // when Height=0, use default ~30?

	sheet            *Sheet
	rowNumber        int // 1-based
	nextColumnNumber int // 1-based, incremented as we add cells
}

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

func CellCoordAsString(col, row int) string {
	if row < 0 {
		panic("invalid row number")
	}
	return ColumnNumberAsLetters(col) + strconv.Itoa(row)
}
