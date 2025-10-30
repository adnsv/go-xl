package xl

import (
	"errors"
	"strconv"
	"strings"
	"unicode"
)

// Sheet represents a single worksheet in a workbook.
// It contains rows, column definitions, and merged cell ranges.
type Sheet struct {
	Name       string
	Rows       []*Row
	Columns    map[int]*Column // 1-based column index to column properties
	MergeCells []MergeCell     // List of merged cell ranges

	workbook      *Workbook
	nextRowNumber int // 1-based, incremented as we add rows
}

// Column represents column-level properties such as width.
type Column struct {
	Width float32 // Column width in Excel units
}

// MergeCell represents a range of cells that should be merged in the worksheet.
type MergeCell struct {
	Ref string // Cell range reference, e.g., "A1:B2"
}

// AddRow adds a new row to the sheet and returns a pointer to it.
// Rows are added sequentially starting from row 1.
func (s *Sheet) AddRow() *Row {
	r := &Row{
		sheet:            s,
		rowNumber:        s.nextRowNumber,
		nextColumnNumber: 1,
	}
	s.nextRowNumber++
	s.Rows = append(s.Rows, r)
	return r
}

// SetColumnWidth sets the width of a column (1-based index).
// Setting width <= 0 removes any custom width, reverting to default.
// Column 1 is "A", column 2 is "B", etc.
func (s *Sheet) SetColumnWidth(colNumber int, w float32) {
	if colNumber <= 0 {
		return
	}
	if w <= 0.0 {
		delete(s.Columns, colNumber)
	} else {
		c, exists := s.Columns[colNumber]
		if !exists {
			c = &Column{
				Width: w,
			}
		} else {
			c.Width = w
		}
		s.Columns[colNumber] = c
	}
}

// parseCellRef parses a cell reference like "A1" and returns the column and row numbers (1-based).
func parseCellRef(ref string) (col, row int, err error) {
	if ref == "" {
		return 0, 0, errors.New("empty cell reference")
	}

	// Split into letter part (column) and number part (row)
	i := 0
	for i < len(ref) && unicode.IsLetter(rune(ref[i])) {
		i++
	}

	if i == 0 || i == len(ref) {
		return 0, 0, errors.New("invalid cell reference format")
	}

	colLetters := strings.ToUpper(ref[:i])
	rowStr := ref[i:]

	// Parse column letters (A=1, B=2, ..., Z=26, AA=27, etc.)
	col = 0
	for _, ch := range colLetters {
		if ch < 'A' || ch > 'Z' {
			return 0, 0, errors.New("invalid column letter")
		}
		col = col*26 + int(ch-'A') + 1
	}

	// Parse row number
	row, err = strconv.Atoi(rowStr)
	if err != nil || row < 1 {
		return 0, 0, errors.New("invalid row number")
	}

	return col, row, nil
}

// parseMergeCellRef parses a merge cell range like "A1:B2" and returns start and end coordinates.
func parseMergeCellRef(ref string) (startCol, startRow, endCol, endRow int, err error) {
	parts := strings.Split(ref, ":")
	if len(parts) != 2 {
		return 0, 0, 0, 0, errors.New("invalid merge cell reference format, expected 'A1:B2'")
	}

	startCol, startRow, err = parseCellRef(parts[0])
	if err != nil {
		return 0, 0, 0, 0, err
	}

	endCol, endRow, err = parseCellRef(parts[1])
	if err != nil {
		return 0, 0, 0, 0, err
	}

	return startCol, startRow, endCol, endRow, nil
}

// Merge merges a range of cells specified by a cell reference like "A1:B2".
// Returns an error if the range is invalid or overlaps with existing merged cells.
func (s *Sheet) Merge(ref string) error {
	// Parse the reference
	startCol, startRow, endCol, endRow, err := parseMergeCellRef(ref)
	if err != nil {
		return err
	}

	// Validate the range
	if err := s.validateMergeRange(startCol, startRow, endCol, endRow); err != nil {
		return err
	}

	// Add to merge cells list
	s.MergeCells = append(s.MergeCells, MergeCell{Ref: ref})
	return nil
}

// MergeRange merges a range of cells specified by column and row coordinates (1-based).
// Returns an error if the range is invalid or overlaps with existing merged cells.
func (s *Sheet) MergeRange(startCol, startRow, endCol, endRow int) error {
	// Validate the range
	if err := s.validateMergeRange(startCol, startRow, endCol, endRow); err != nil {
		return err
	}

	// Normalize coordinates to ensure start <= end
	if startCol > endCol {
		startCol, endCol = endCol, startCol
	}
	if startRow > endRow {
		startRow, endRow = endRow, startRow
	}

	// Convert to cell reference format
	ref := CellCoordAsString(startCol, startRow) + ":" + CellCoordAsString(endCol, endRow)

	// Add to merge cells list
	s.MergeCells = append(s.MergeCells, MergeCell{Ref: ref})
	return nil
}

// validateMergeRange validates that a merge range is valid and doesn't overlap with existing merges.
func (s *Sheet) validateMergeRange(startCol, startRow, endCol, endRow int) error {
	// Ensure coordinates are in correct order
	if startCol > endCol {
		startCol, endCol = endCol, startCol
	}
	if startRow > endRow {
		startRow, endRow = endRow, startRow
	}

	// Ensure range spans at least 2 cells
	if startCol == endCol && startRow == endRow {
		return errors.New("merge range must span at least 2 cells")
	}

	// Check for overlaps with existing merge ranges
	for _, mc := range s.MergeCells {
		existStartCol, existStartRow, existEndCol, existEndRow, err := parseMergeCellRef(mc.Ref)
		if err != nil {
			continue // Skip invalid existing ranges
		}

		// Normalize existing range
		if existStartCol > existEndCol {
			existStartCol, existEndCol = existEndCol, existStartCol
		}
		if existStartRow > existEndRow {
			existStartRow, existEndRow = existEndRow, existStartRow
		}

		// Check for overlap
		if !(endCol < existStartCol || startCol > existEndCol ||
			endRow < existStartRow || startRow > existEndRow) {
			return errors.New("merge range overlaps with existing merged cells")
		}
	}

	return nil
}
