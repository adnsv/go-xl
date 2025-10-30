package xl

import (
	"errors"
	"fmt"
	"strings"
	"unicode/utf8"
)

// Workbook represents an Excel workbook containing one or more worksheets.
type Workbook struct {
	AppName string   // Optional application name that created the workbook
	Sheets  []*Sheet // List of worksheets in the workbook

	sheetMap map[string]*Sheet // Maps sheet name to sheet for duplicate detection
	lastIdN  int               // Counter for generating unique IDs
}

// NewWorkbook creates and initializes a new empty workbook.
func NewWorkbook() *Workbook {
	return &Workbook{
		sheetMap: map[string]*Sheet{},
	}
}

// AddSheet adds a new worksheet to the workbook with the specified name.
// Returns an error if a sheet with the same name already exists or if the name is invalid.
// Sheet names must be 1-31 characters, cannot start/end with single quotes,
// and cannot contain: : \ / ? * [ ]
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

// validateSheetName checks if a sheet name conforms to Excel's naming rules.
// Valid names must be 1-31 characters long, cannot start or end with single quotes,
// and cannot contain the characters: : \ / ? * [ ]
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
