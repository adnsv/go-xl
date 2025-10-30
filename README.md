# go-xl

A lightweight, type-safe Go library for generating Excel (XLSX) files with a focus on simplicity and standards compliance.

## Features

- **Simple API** - Intuitive, builder-style interface
- **Type-safe formatting** - Compile-time validation for alignment and font properties
- **OpenXML compliant** - Follows ECMA-376 specification
- **Cell formatting**
  - Horizontal and vertical alignment with type-safe constants
  - Font properties: bold, italic, underline (5 types), strikethrough, size
  - Merged cells support
- **Multiple data types** - Strings, numbers, booleans, formulas, errors
- **Images** - Embed PNG and JPEG images in cells
- **Column widths and row heights** - Custom sizing support
- **Flexible output** - Write to ZIP file or directory storage for debugging

## Installation

```bash
go get github.com/adnsv/go-xl/xl
```

## Quick Start

```go
package main

import (
    "log"
    "os"

    "github.com/adnsv/go-xl/xl"
)

func main() {
    // Create a new workbook
    wb := xl.NewWorkbook()
    wb.AppName = "My Application"

    // Add a sheet
    sheet, err := wb.AddSheet("sheet1")
    if err != nil {
        log.Fatal(err)
    }

    // Add a row with cells
    row := sheet.AddRow()
    cell := row.AddCell()
    cell.SetStr("Hello, Excel!")

    // Write to file
    f, err := os.Create("output.xlsx")
    if err != nil {
        log.Fatal(err)
    }
    defer f.Close()

    zs := xl.NewZipStorage(f)
    defer zs.Close()

    w := xl.NewWriter(zs)
    if err := w.Write(wb); err != nil {
        log.Fatal(err)
    }
}
```

## Usage Examples

### Creating Cells with Different Data Types

```go
row := sheet.AddRow()

// String
cell1 := row.AddCell()
cell1.SetStr("Product Name")

// Number
cell2 := row.AddCell()
cell2.SetInt(42)

// Float
cell3 := row.AddCell()
cell3.SetFloat(99.99)

// Boolean
cell4 := row.AddCell()
cell4.SetBool(true)
```

### Cell Alignment

```go
cell := row.AddCell()
cell.SetStr("Centered Text")

// Horizontal alignment (type-safe)
cell.XF.Alignment.Horizontal = xl.HAlignCenter  // or: Left, Right, General, Fill, Justify, CenterContinuous, Distributed

// Vertical alignment (type-safe)
cell.XF.Alignment.Vertical = xl.VAlignCenter    // or: Top, Bottom, Justify, Distributed
```

**Available alignment constants:**

Horizontal: `HAlignGeneral`, `HAlignLeft`, `HAlignCenter`, `HAlignRight`, `HAlignFill`, `HAlignJustify`, `HAlignCenterContinuous`, `HAlignDistributed`

Vertical: `VAlignTop`, `VAlignCenter`, `VAlignBottom`, `VAlignJustify`, `VAlignDistributed`

### Font Formatting

```go
cell := row.AddCell()
cell.SetStr("Formatted Text")

// Font properties
cell.XF.Font.Bold = true
cell.XF.Font.Italic = true
cell.XF.Font.Size = 14                               // Points
cell.XF.Font.Underline = xl.UnderlineSingle          // or: Double, SingleAccounting, DoubleAccounting, None
cell.XF.Font.Strikethrough = true
```

**Example: Bold header with custom size**

```go
headerCell := row.AddCell()
headerCell.SetStr("Sales Report")
headerCell.XF.Font.Bold = true
headerCell.XF.Font.Size = 18
headerCell.XF.Alignment.Horizontal = xl.HAlignCenter
```

### Merged Cells

```go
// Add cells that will be merged
row := sheet.AddRow()
cell := row.AddCell()
cell.SetStr("Merged Cell Content")
cell.XF.Alignment.Horizontal = xl.HAlignCenter
cell.XF.Alignment.Vertical = xl.VAlignCenter

row.AddCell() // Empty cell, part of merge
row.AddCell() // Empty cell, part of merge

// Merge using cell reference
err := sheet.Merge("A1:C1")

// Or merge using coordinates (1-based)
err = sheet.MergeRange(1, 1, 3, 1) // startCol, startRow, endCol, endRow
```

### Column Widths and Row Heights

```go
// Set column width (1-based index, width in Excel units)
sheet.SetColumnWidth(1, 20)   // Column A
sheet.SetColumnWidth(2, 30)   // Column B

// Set row height
row := sheet.AddRow()
row.Height = 40  // Height in points
```

### Adding Images

```go
import (
    "os"
    "path/filepath"
)

// Read image file
blob, err := os.ReadFile("logo.png")
if err != nil {
    log.Fatal(err)
}

// Create a tall row for the image
row := sheet.AddRow()
row.Height = 64

// Add image to cell
cell := row.AddCell()
cell.SetPicture(&xl.PictureInfo{
    Extension: filepath.Ext("logo.png"), // ".png" or ".jpg"/".jpeg"
    Blob:      blob,
})
```

**Supported image formats:** PNG, JPEG

### Complete Example: Formatted Table

```go
wb := xl.NewWorkbook()
sheet, _ := wb.AddSheet("Products")

// Set column widths
sheet.SetColumnWidth(1, 20) // Product name
sheet.SetColumnWidth(2, 10) // Price
sheet.SetColumnWidth(3, 10) // Stock

// Header row
headerRow := sheet.AddRow()
headerRow.Height = 30

headers := []string{"Product", "Price", "Stock"}
for _, header := range headers {
    cell := headerRow.AddCell()
    cell.SetStr(header)
    cell.XF.Font.Bold = true
    cell.XF.Font.Size = 12
    cell.XF.Alignment.Horizontal = xl.HAlignCenter
    cell.XF.Alignment.Vertical = xl.VAlignCenter
}

// Data rows
products := []struct{name string; price float64; stock int}{
    {"Widget A", 19.99, 100},
    {"Widget B", 24.99, 50},
    {"Widget C", 14.99, 75},
}

for _, p := range products {
    row := sheet.AddRow()

    nameCell := row.AddCell()
    nameCell.SetStr(p.name)

    priceCell := row.AddCell()
    priceCell.SetFloat(p.price)
    priceCell.XF.Alignment.Horizontal = xl.HAlignRight

    stockCell := row.AddCell()
    stockCell.SetInt(int64(p.stock))
    stockCell.XF.Alignment.Horizontal = xl.HAlignCenter
}
```

## Storage Options

### ZIP Storage (Standard XLSX File)

```go
f, err := os.Create("output.xlsx")
if err != nil {
    log.Fatal(err)
}
defer f.Close()

zs := xl.NewZipStorage(f)
defer zs.Close()

w := xl.NewWriter(zs)
err = w.Write(wb)
```

### Directory Storage (For Debugging)

Useful for inspecting the generated XML files:

```go
ds := xl.NewDirStorage("./output_dir")
w := xl.NewWriter(ds)
err := w.Write(wb)

// This creates: output_dir/xl/worksheets/sheet1.xml
//               output_dir/xl/styles.xml
//               output_dir/[Content_Types].xml
//               etc.
```

## API Reference

### Core Types

**Workbook**
```go
wb := xl.NewWorkbook()
wb.AppName = "My App"             // Optional application name
sheet, err := wb.AddSheet("name") // Add a sheet (unique names required)
```

**Sheet**
```go
row := sheet.AddRow()                           // Add a new row
sheet.SetColumnWidth(colNumber int, width float32) // Set column width (1-based)
sheet.Merge(ref string) error                   // Merge cells by range ("A1:B2")
sheet.MergeRange(startCol, startRow, endCol, endRow int) error // Merge by coordinates
```

**Row**
```go
cell := row.AddCell()      // Add a new cell
row.Height = 30            // Set row height in points
```

**Cell**
```go
cell.SetStr(s string)      // Set string value
cell.SetInt(i int64)       // Set integer value
cell.SetFloat(f float64)   // Set float value
cell.SetBool(b bool)       // Set boolean value
cell.SetPicture(p *PictureInfo) // Set image

// Formatting
cell.XF.Alignment.Horizontal = xl.HAlignCenter
cell.XF.Alignment.Vertical = xl.VAlignCenter
cell.XF.Font.Bold = true
cell.XF.Font.Italic = true
cell.XF.Font.Size = 14
cell.XF.Font.Underline = xl.UnderlineSingle
cell.XF.Font.Strikethrough = true
```

### Type-Safe Constants

**Horizontal Alignment:**
- `HAlignGeneral` (default)
- `HAlignLeft`
- `HAlignCenter`
- `HAlignRight`
- `HAlignFill`
- `HAlignJustify`
- `HAlignCenterContinuous`
- `HAlignDistributed`

**Vertical Alignment:**
- `VAlignTop`
- `VAlignCenter`
- `VAlignBottom` (default)
- `VAlignJustify`
- `VAlignDistributed`

**Underline Types:**
- `UnderlineNone` (default)
- `UnderlineSingle`
- `UnderlineDouble`
- `UnderlineSingleAccounting`
- `UnderlineDoubleAccounting`

## Design Principles

- **Type safety** - Use custom types with constants to prevent invalid values at compile time
- **Standards compliance** - Follows ECMA-376 OpenXML SpreadsheetML specification
- **Simplicity** - Minimal API surface with sensible defaults
- **No dependencies** - Only uses standard library (plus xml writer utility)
- **Lazy evaluation** - Deduplicates styles and fonts during write, not during construction

## Current Limitations

- Read-only generation (no reading/parsing of existing files)
- Limited formatting options (fonts, alignment, merged cells only)
- No formulas evaluation (formulas stored as text)
- No charts, pivot tables, or advanced features
- No color support yet (planned)
- No borders support yet (planned)
- No number formats beyond defaults (planned)

## License

See LICENSE file for details.

## Contributing

Contributions are welcome! Please ensure:
- Code follows Go conventions
- Changes maintain backward compatibility (or are clearly marked as breaking)
- OpenXML specification compliance is preserved
- Tests are included for new features

## Related Projects

This library is designed as a lightweight alternative to:
- [tealeg/xlsx](https://github.com/tealeg/xlsx)
- [qax-os/excelize](https://github.com/qax-os/excelize)

With a focus on type safety and minimal dependencies.
