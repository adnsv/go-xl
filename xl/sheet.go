package xl

type Sheet struct {
	Name    string
	Rows    []*Row
	Columns map[int]*Column // 1-based

	workbook      *Workbook
	nextRowNumber int // 1-based, incremented as we add rows
}

type Column struct {
	Width float32
}

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
