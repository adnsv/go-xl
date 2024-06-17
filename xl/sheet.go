package xl

type Sheet struct {
	Name string
	Rows []*Row

	workbook      *Workbook
	nextRowNumber int // 1-based, incremented as we add rows
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
