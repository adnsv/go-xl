package xl

import (
	"errors"
	"fmt"
	"strings"
	"unicode/utf8"
)

type Workbook struct {
	AppName string
	Sheets  []*Sheet

	sheetMap map[string]*Sheet
	lastIdN  int
}

func NewWorkbook() *Workbook {
	return &Workbook{
		sheetMap: map[string]*Sheet{},
	}
}

func (wb *Workbook) AddSheet(name string) (*Sheet, error) {
	if _, exists := wb.sheetMap[name]; exists {
		return nil, fmt.Errorf("duplicate sheet name '%s'", name)
	}

	if err := validateSheetName(name); err != nil {
		return nil, err
	}

	sheet := &Sheet{
		workbook:      wb,
		Name:          name,
		Columns:       map[int]*Column{},
		nextRowNumber: 1,
	}

	wb.Sheets = append(wb.Sheets, sheet)
	wb.sheetMap[name] = sheet

	return sheet, nil
}

func validateSheetName(s string) error {
	n := utf8.RuneCountInString(s)
	if n == 0 {
		return errors.New("empty sheet name is not allowed")
	} else if n > 31 {
		return errors.New("the sheet name is too long")
	}
	if strings.HasPrefix(s, "'") || strings.HasSuffix(s, "'") {
		return errors.New("the first or last character of the sheet name can not be a single quote")
	}
	if strings.ContainsAny(s, ":\\/?*[]") {
		return errors.New("the sheet can not contain any of the characters :\\/?*[]")
	}
	return nil
}
