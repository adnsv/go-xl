package xl

import (
	"bytes"
	"errors"
	"fmt"
	"slices"
	"strings"
	"time"

	"github.com/adnsv/srw/xml"

	"golang.org/x/exp/constraints"
	"golang.org/x/exp/maps"
)

// Writer is responsible for generating OpenXML SpreadsheetML files from a workbook.
// It manages shared strings, styles, fonts, media, and all XML part generation.
type Writer struct {
	out            Storage
	lastGlobalId   int
	lastWorkbookId int
	lastRichDataId int

	GlobalRels          map[string]RelInfo // maps id to absolute path
	WorkbookRels        map[string]RelInfo // maps id to absolute paths
	DefaultContentTypes map[string]string  // maps path extension to content-type
	PartContentTypes    map[string]string  // maps path partname to content-type

	sharedStrings   []string
	sharedStringMap map[string]int // 1-based index into sharedStrings

	media    []*MediaInfo
	mediaMap map[string]*MediaInfo // maps media name to media info

	xfs   []*XF
	fonts []*Font

	RichDataRels map[string]RelInfo
}

// RelInfo contains relationship information for OpenXML parts.
type RelInfo struct {
	Type   string // url to schema type
	Target string // relative path
}

// MediaInfo contains embedded media file information (images).
type MediaInfo struct {
	Name string // hashed blob + extension
	Blob []byte // raw file data
	IId  int    // internal ID
	RId  string // relationship ID
}

// NewWriter creates a new Writer that will output to the specified storage.
// The storage can be a ZIP file storage or directory storage for debugging.
func NewWriter(s Storage) *Writer {
	w := &Writer{
		out:                 s,
		GlobalRels:          map[string]RelInfo{},
		WorkbookRels:        map[string]RelInfo{},
		DefaultContentTypes: map[string]string{},
		PartContentTypes:    map[string]string{},

		sharedStringMap: map[string]int{},

		mediaMap: map[string]*MediaInfo{},

		RichDataRels: map[string]RelInfo{},
	}

	w.DefaultContentTypes["xml"] = "application/xml"
	w.DefaultContentTypes["rels"] = "application/vnd.openxmlformats-package.relationships+xml"

	return w
}

// SharedString adds a string to the shared string table and returns its index.
// If the string already exists, returns the existing index.
// This is used internally for efficient string storage in cells.
func (w *Writer) SharedString(s string) int {
	if i, ok := w.sharedStringMap[s]; ok {
		return i
	}
	i := len(w.sharedStrings)
	w.sharedStrings = append(w.sharedStrings, s)
	w.sharedStringMap[s] = i
	return i
}

func (w *Writer) nextGlobalID() (int, string) {
	w.lastGlobalId++
	return w.lastGlobalId, fmt.Sprintf("rId%d", w.lastGlobalId)
}
func (w *Writer) nextWorkbookID() (int, string) {
	w.lastWorkbookId++
	return w.lastWorkbookId, fmt.Sprintf("rId%d", w.lastWorkbookId)
}
func (w *Writer) nextRichDataID() (int, string) {
	w.lastRichDataId++
	return w.lastRichDataId, fmt.Sprintf("rId%d", w.lastRichDataId)
}

// Write generates a complete Excel workbook file from the given Workbook.
// It writes all necessary XML parts, relationships, and content types to the storage.
// Returns an error if any part of the generation fails.
func (w *Writer) Write(wb *Workbook) error {
	var err error

	err = w.writeWorkbook(wb)
	if err != nil {
		return err
	}

	if len(w.media) > 0 {

		err = w.writeMedia()
		if err != nil {
			return err
		}

		err = w.writeRichValueRel()
		if err != nil {
			return err
		}

		err = w.writeRels("/xl/richData/_rels/richValueRel.xml.rels", w.RichDataRels)
		if err != nil {
			return err
		}

		err = w.writeRichValueStructure()
		if err != nil {
			return err
		}

		/*
			err = w.writeRichValueTypes()
			if err != nil {
				return err
			}
		*/

		err = w.writeRichValueData()
		if err != nil {
			return err
		}

		err = w.writeMetadata()
		if err != nil {
			return err
		}
	}

	err = w.writeCoreProperties()
	if err != nil {
		return err
	}
	err = w.writeExtendedProperties(wb.AppName)
	if err != nil {
		return err
	}

	if len(w.sharedStrings) > 0 {
		err = w.writeSharedStrings()
		if err != nil {
			return err
		}
	}

	if len(w.xfs) > 0 {
		err = w.writeStyles()
		if err != nil {
			return err
		}
	}

	err = w.writeRels("/xl/_rels/workbook.xml.rels", w.WorkbookRels)
	if err != nil {
		return err
	}

	err = w.writeRels("/_rels/.rels", w.GlobalRels)
	if err != nil {
		return err
	}

	err = w.writeContentTypes()
	if err != nil {
		return err
	}

	return nil
}

func (w *Writer) writeCoreProperties() error {
	_, rid := w.nextGlobalID()

	relpath := "docProps/core.xml"
	abspath := "/" + relpath

	w.PartContentTypes[abspath] = "application/vnd.openxmlformats-package.core-properties+xml"
	w.GlobalRels[rid] = RelInfo{
		Type:   "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties",
		Target: relpath,
	}

	bb := bytes.Buffer{}
	x := xml.NewWriter(&bb, xml.WriterConfig{Indent: xml.Indent2Spaces})

	x.XmlStandaloneDecl()
	x.OTag("cp:coreProperties")
	x.Attr("xmlns:cp", "http://schemas.openxmlformats.org/package/2006/metadata/core-properties")
	x.Attr("xmlns:dc", "http://purl.org/dc/elements/1.1/")
	x.Attr("xmlns:dcterms", "http://purl.org/dc/terms/")
	x.Attr("xmlns:dcmitype", "http://purl.org/dc/dcmitype/")
	x.Attr("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance")

	x.OTag("+dcterms:created")
	x.Attr("xsi:type", "dcterms:W3CDTF")
	x.Write(time.Now().UTC().Format(time.RFC3339))
	x.CTag()

	x.CTag()

	return w.out.WriteBlob(abspath, bb.Bytes())
}

func (w *Writer) writeExtendedProperties(appname string) error {
	_, rid := w.nextGlobalID()

	relpath := "docProps/app.xml"
	abspath := "/" + relpath

	w.PartContentTypes[abspath] = "application/vnd.openxmlformats-officedocument.extended-properties+xml"
	w.GlobalRels[rid] = RelInfo{
		Type:   "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties",
		Target: relpath,
	}

	bb := bytes.Buffer{}
	x := xml.NewWriter(&bb, xml.WriterConfig{Indent: xml.Indent2Spaces})
	x.XmlStandaloneDecl()

	x.OTag("Properties")
	x.Attr("xmlns", "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties")
	x.Attr("xmlns:vt", "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes")

	if appname != "" {
		x.OTag("+Application").String(appname).CTag()
	}

	x.CTag()

	return w.out.WriteBlob(abspath, bb.Bytes())
}

func (w *Writer) writeContentTypes() error {
	bb := bytes.Buffer{}
	x := xml.NewWriter(&bb, xml.WriterConfig{Indent: xml.Indent2Spaces})

	x.XmlStandaloneDecl()
	x.OTag("Types")
	x.Attr("xmlns", "http://schemas.openxmlformats.org/package/2006/content-types")
	enumerate(w.DefaultContentTypes, func(ext, ctype string) error {
		x.OTag("+Default").Attr("Extension", ext).Attr("ContentType", ctype).CTag()
		return nil
	})
	enumerate(w.PartContentTypes, func(abspath, ctype string) error {
		x.OTag("+Override").Attr("PartName", abspath).Attr("ContentType", ctype).CTag()
		return nil
	})

	x.CTag()

	return w.out.WriteBlob("[Content_Types].xml", bb.Bytes())
}

func (w *Writer) writeStyles() error {
	_, rid := w.nextWorkbookID()

	relpath := "styles.xml"
	abspath := "/xl/" + relpath

	w.PartContentTypes[abspath] = "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"
	w.WorkbookRels[rid] = RelInfo{
		Type:   "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
		Target: relpath,
	}

	bb := bytes.Buffer{}
	x := xml.NewWriter(&bb, xml.WriterConfig{Indent: xml.Indent2Spaces})
	x.XmlStandaloneDecl()

	x.OTag("styleSheet")
	x.Attr("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")

	// Collect unique fonts from all xfs
	for _, xf := range w.xfs {
		if !xf.Font.IsDefault() {
			if w.FindFont(&xf.Font) < 0 {
				w.fonts = append(w.fonts, &xf.Font)
			}
		}
	}

	// Write fonts section
	fontCount := len(w.fonts) + 1 // +1 for default font at index 0
	x.OTag("+fonts").Attr("count", fontCount)

	// Font 0: Default font
	x.OTag("+font")
	x.OTag("sz").Attr("val", 11).CTag()
	x.OTag("name").Attr("val", "Calibri").CTag()
	x.OTag("family").Attr("val", 2).CTag()
	x.CTag() // font

	// Custom fonts
	for _, font := range w.fonts {
		x.OTag("+font")

		// Element order: b, i, strike, u, sz, color, name, family
		if font.Bold {
			x.OTag("b").CTag()
		}
		if font.Italic {
			x.OTag("i").CTag()
		}
		if font.Strikethrough {
			x.OTag("strike").CTag()
		}
		if font.Underline != UnderlineNone {
			if font.Underline == UnderlineSingle {
				x.OTag("u").CTag() // Empty element for single underline
			} else {
				x.OTag("u").Attr("val", string(font.Underline)).CTag()
			}
		}

		// Size (use 11 if not specified)
		size := font.Size
		if size == 0 {
			size = 11
		}
		x.OTag("sz").Attr("val", size).CTag()

		// Basic font properties for compatibility
		x.OTag("name").Attr("val", "Calibri").CTag()
		x.OTag("family").Attr("val", 2).CTag()

		x.CTag() // font
	}
	x.CTag() // fonts

	x.OTag("+fills").Attr("count", 1)
	x.OTag("+fill")
	x.OTag("patternFill").Attr("patternType", "none").CTag()
	x.CTag() // fill
	x.CTag() // fills

	x.OTag("+borders").Attr("count", 1)
	x.OTag("+border")
	x.OTag("+left").CTag()
	x.OTag("+right").CTag()
	x.OTag("+top").CTag()
	x.OTag("+bottom").CTag()
	x.OTag("+diagonal").CTag()
	x.CTag() // border
	x.CTag() // borders

	x.OTag("+cellStyleXfs").Attr("count", 1)
	x.OTag("+xf")
	x.Attr("numFmtId", "0")
	x.Attr("fontId", "0")
	x.Attr("fillId", "0")
	x.Attr("borderId", "0")
	x.CTag()
	x.CTag() //cellStyleXfs

	x.OTag("+cellXfs").Attr("count", len(w.xfs)+1)
	// Default xf (index 0)
	x.OTag("+xf")
	x.Attr("numFmtId", "0")
	x.Attr("fontId", "0")
	x.Attr("fillId", "0")
	x.Attr("borderId", "0")
	x.Attr("xfId", "0")
	x.CTag()
	// Custom xfs collected from cells
	for _, xf := range w.xfs {
		x.OTag("+xf")
		x.Attr("numFmtId", "0")

		// Determine font ID
		fontId := 0 // Default font
		if !xf.Font.IsDefault() {
			fontIdx := w.FindFont(&xf.Font)
			if fontIdx >= 0 {
				fontId = fontIdx + 1 // +1 because default font is at index 0
			}
		}
		x.Attr("fontId", fontId)

		x.Attr("fillId", "0")
		x.Attr("borderId", "0")
		x.Attr("xfId", "0")

		// Set applyFont if using custom font
		if !xf.Font.IsDefault() {
			x.Attr("applyFont", "1")
		}

		// Set applyAlignment if using custom alignment
		if !xf.Alignment.Empty() {
			x.Attr("applyAlignment", "1")
		}

		// Write alignment element if not empty
		if !xf.Alignment.Empty() {
			x.OTag("alignment")
			if xf.Alignment.Horizontal != "" {
				x.Attr("horizontal", xf.Alignment.Horizontal)
			}
			if xf.Alignment.Vertical != "" {
				x.Attr("vertical", xf.Alignment.Vertical)
			}
			x.CTag() // alignment
		}

		x.CTag() // xf
	}
	x.CTag() // cellXfs

	x.CTag()

	return w.out.WriteBlob(abspath, bb.Bytes())
}

func (w *Writer) writeWorkbook(wb *Workbook) error {
	_, rid := w.nextGlobalID()

	relpath := "xl/workbook.xml"
	abspath := "/" + relpath

	w.PartContentTypes[abspath] = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
	w.GlobalRels[rid] = RelInfo{
		Type:   "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
		Target: relpath,
	}

	bb := bytes.Buffer{}
	x := xml.NewWriter(&bb, xml.WriterConfig{Indent: xml.Indent2Spaces})
	x.XmlStandaloneDecl()

	x.OTag("workbook")
	x.Attr("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
	x.Attr("xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")

	/*
		if wb.AppName != "" {
			x.OTag("+fileVersion")
			x.Attr("appName", wb.AppName)
			x.CTag()
		}

		x.OTag("+workbookPr")
		x.Attr("showObjects", "all")
		x.Attr("date1904", "false")
		x.CTag()

		x.OTag("+<workbookProtection")
		x.CTag()

		x.OTag("+bookViews")
		{
			x.OTag("+workbookView")
			x.Attr("showHorizontalScroll", "true")
			x.Attr("showVerticalScroll", "true")
			x.Attr("showSheetTabs", "true")
			x.Attr("tabRatio", "204")
			x.Attr("windowHeight", "8192")
			x.Attr("windowWidth", "16384")
			x.Attr("xWindow", "0")
			x.Attr("yWindow", "0")
			x.CTag()
		}
		x.CTag()
	*/

	x.OTag("+sheets")
	for _, sheet := range wb.Sheets {
		sheet_id, sheet_rid := w.nextWorkbookID()
		{
			x.OTag("+sheet")
			x.Attr("name", sheet.Name)
			x.Attr("sheetId", sheet_id)
			x.Attr("r:id", sheet_rid)
			x.CTag()
		}

		err := w.writeSheet(sheet, sheet_rid)
		if err != nil {
			return err
		}
	}
	x.CTag()

	/*

		x.OTag("+definedNames")
		x.CTag()

		x.OTag("+calcPr")
		x.Attr("iterateCount", "100")
		x.Attr("refMode", "A1")
		x.Attr("iterateDelta", "0.001")
		x.CTag()
	*/

	x.CTag()

	return w.out.WriteBlob(abspath, bb.Bytes())
}

func (w *Writer) FindXF(xf *XF) int {
	for i, v := range w.xfs {
		if *v == *xf {
			return i
		}
	}
	return -1
}

// FindFont returns the index of a matching font in the fonts slice, or -1 if not found.
func (w *Writer) FindFont(font *Font) int {
	for i, f := range w.fonts {
		if *f == *font {
			return i
		}
	}
	return -1
}

func (w *Writer) writeSheet(sh *Sheet, rid string) error {
	relpath := "worksheets/" + sh.Name + ".xml"
	abspath := "/xl/" + relpath

	w.PartContentTypes[abspath] = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
	w.WorkbookRels[rid] = RelInfo{
		Type:   "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
		Target: relpath,
	}

	bb := bytes.Buffer{}
	x := xml.NewWriter(&bb, xml.WriterConfig{Indent: xml.Indent2Spaces})
	x.XmlStandaloneDecl()

	x.OTag("worksheet")
	x.Attr("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
	x.Attr("xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")

	if len(sh.Columns) > 0 {
		x.OTag("+cols")
		enumerate(sh.Columns, func(n int, v *Column) error {
			x.OTag("+col").Attr("min", n).Attr("max", n)
			if v.Width > 0 {
				x.Attr("width", v.Width).Attr("customWidth", 1)
			}
			x.CTag()
			return nil
		})
		x.CTag()
	}

	x.OTag("+sheetData")
	for _, row := range sh.Rows {
		x.OTag("+row").Attr("r", row.rowNumber)
		if row.Height > 0 {
			x.Attr("ht", row.Height).Attr("customHeight", 1)
		}

		for _, cell := range row.Cells {
			x.OTag("+c").Attr("r", cell.coord)

			if !cell.XF.Empty() {
				i := w.FindXF(&cell.XF)
				if i < 0 {
					w.xfs = append(w.xfs, &cell.XF)
					i = len(w.xfs) - 1
				}
				// Style index is xfs array index + 1 (because default xf is at index 0)
				x.Attr("s", i+1)
			}

			switch cell.typ {
			case CellTypeBool:
				x.Attr("t", "b")
				x.OTag("v").Write(cell.v).CTag()
			case CellTypeNumber:
				x.Attr("t", "n")
				x.OTag("v").Write(cell.v).CTag()
			case CellTypeError:
				x.Attr("t", "e")
				x.OTag("v").Write(cell.v).CTag()
			case CellTypeSharedString:
				x.Attr("t", "s")
				x.OTag("v").Write(w.SharedString(cell.v)).CTag()
			case cellTypePicture:
				if cell.picture == nil {
					return errors.New("missing picture data")
				}
				ext := strings.ToLower(cell.picture.Extension)
				if ext == ".jpg" {
					ext = ".jpeg"
				}
				if ext == ".jpeg" {
					w.DefaultContentTypes["jpeg"] = "image/jpeg"
				} else if ext == ".png" {
					w.DefaultContentTypes["png"] = "image/png"
				} else {
					return fmt.Errorf("unsupported image extension %s", ext)
				}
				n := fmt.Sprintf("%.16x%s", BlobHash(cell.picture.Blob), ext)
				info, ok := w.mediaMap[n]
				if !ok {
					_, rid := w.nextRichDataID()
					info = &MediaInfo{
						Name: n,
						Blob: cell.picture.Blob,
						IId:  len(w.media),
						RId:  rid,
					}
					w.mediaMap[n] = info
					w.media = append(w.media, info)
				}
				if len(info.Blob) == 0 {
					return errors.New("empty picture data")
				}

				x.Attr("t", "e").Attr("vm", info.IId+1)
				x.OTag("v").Write("#VALUE!").CTag()
			}
			x.CTag() // c
		}

		x.CTag() // row
	}
	x.CTag() // sheetData

	// Write mergeCells if any exist
	if len(sh.MergeCells) > 0 {
		x.OTag("+mergeCells").Attr("count", len(sh.MergeCells))
		for _, mc := range sh.MergeCells {
			x.OTag("+mergeCell").Attr("ref", mc.Ref).CTag()
		}
		x.CTag() // mergeCells
	}

	x.CTag() // worksheet

	return w.out.WriteBlob(abspath, bb.Bytes())
}

func (w *Writer) writeSharedStrings() error {
	_, rid := w.nextWorkbookID()

	relpath := "sharedStrings.xml"
	abspath := "/xl/" + relpath

	w.PartContentTypes[abspath] = "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"
	w.WorkbookRels[rid] = RelInfo{
		Type:   "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings",
		Target: relpath,
	}

	bb := bytes.Buffer{}
	x := xml.NewWriter(&bb, xml.WriterConfig{Indent: xml.Indent2Spaces})
	x.XmlStandaloneDecl()

	x.OTag("sst")
	x.Attr("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
	x.Attr("count", len(w.sharedStrings))
	x.Attr("uniqueCount", len(w.sharedStrings))

	for _, s := range w.sharedStrings {
		x.OTag("+si")
		x.OTag("t").Write(s).CTag()
		x.CTag()
	}

	x.CTag()

	return w.out.WriteBlob(abspath, bb.Bytes())
}

func (w *Writer) writeMedia() error {
	if len(w.media) == 0 {
		return nil
	}

	for _, m := range w.media {
		fn := "/xl/media/" + m.Name
		err := w.out.WriteBlob(fn, m.Blob)
		if err != nil {
			return err
		}
		w.RichDataRels[m.RId] = RelInfo{
			Type:   "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
			Target: "../media/" + m.Name,
		}
	}
	return nil
}

func (w *Writer) writeMetadata() error {
	_, rid := w.nextWorkbookID()

	relpath := "metadata.xml"
	abspath := "/xl/" + relpath

	w.PartContentTypes[abspath] = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"
	w.WorkbookRels[rid] = RelInfo{
		Type:   "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata",
		Target: relpath,
	}

	bb := bytes.Buffer{}
	x := xml.NewWriter(&bb, xml.WriterConfig{Indent: xml.Indent2Spaces})
	x.XmlStandaloneDecl()

	x.OTag("metadata")
	x.Attr("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
	x.Attr("xmlns:xlrd", "http://schemas.microsoft.com/office/spreadsheetml/2017/richdata")

	x.OTag("+metadataTypes").Attr("count", 1)
	x.OTag("+metadataType")
	x.Attr("name", "XLRICHVALUE")
	x.Attr("minSupportedVersion", "120000")
	for _, s := range []xml.NameString{"copy", "pasteAll", "pasteValues",
		"merge", "splitFirst", "rowColShift", "clearFormats",
		"clearComments", "assign", "coerce"} {
		x.Attr(s, 1)
	}
	x.CTag() // metadataType
	x.CTag() // metadataTypes

	x.OTag("futureMetadata").Attr("name", "XLRICHVALUE").Attr("count", len(w.media))
	for _, m := range w.media {
		x.OTag("+bk")
		x.OTag("extLst")
		x.OTag("ext").Attr("uri", "{3e2802c4-a4d2-4d8b-9148-e3be6c30e623}")
		x.OTag("xlrd:rvb").Attr("i", m.IId).CTag()
		x.CTag() // ext
		x.CTag() // extLst
		x.CTag() // bk
	}
	x.CTag() // futureMetadata

	x.OTag("valueMetadata").Attr("count", len(w.media))
	for _, m := range w.media {
		x.OTag("+bk")
		x.OTag("rc").Attr("t", 1).Attr("v", m.IId).CTag()
		x.CTag() // bk
	}
	x.CTag() // valueMetadata

	x.CTag() // metadata

	return w.out.WriteBlob(abspath, bb.Bytes())
}

func (w *Writer) writeRichValueRel() error {
	_, rid := w.nextWorkbookID()

	relpath := "richData/richValueRel.xml"
	abspath := "/xl/" + relpath

	w.PartContentTypes[abspath] = "application/vnd.ms-excel.richvaluerel+xml"
	w.WorkbookRels[rid] = RelInfo{
		Type:   "http://schemas.microsoft.com/office/2022/10/relationships/richValueRel",
		Target: relpath,
	}

	bb := bytes.Buffer{}
	x := xml.NewWriter(&bb, xml.WriterConfig{Indent: xml.Indent2Spaces})
	x.XmlStandaloneDecl()

	x.OTag("richValueRels")
	x.Attr("xmlns", "http://schemas.microsoft.com/office/spreadsheetml/2022/richvaluerel")
	x.Attr("xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")

	for _, m := range w.media {
		x.OTag("+rel")
		x.Attr("r:id", m.RId)
		x.CTag()
	}

	x.CTag()

	return w.out.WriteBlob(abspath, bb.Bytes())
}

func (w *Writer) writeRichValueStructure() error {
	_, rid := w.nextWorkbookID()

	relpath := "richData/rdrichvaluestructure.xml"
	abspath := "/xl/" + relpath

	w.PartContentTypes[abspath] = "application/vnd.ms-excel.rdrichvaluestructure+xml"
	w.WorkbookRels[rid] = RelInfo{
		Type:   "http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueStructure",
		Target: relpath,
	}

	bb := bytes.Buffer{}
	x := xml.NewWriter(&bb, xml.WriterConfig{Indent: xml.Indent2Spaces})
	x.XmlStandaloneDecl()

	x.OTag("rvStructures")
	x.Attr("xmlns", "http://schemas.microsoft.com/office/spreadsheetml/2017/richdata")
	x.Attr("count", 1)

	// define _localImage{Id, CalcOrigin}
	x.OTag("+s").Attr("t", "_localImage")
	x.OTag("+k").Attr("n", "_rvRel:LocalImageIdentifier").Attr("t", "i").CTag()
	x.OTag("+k").Attr("n", "CalcOrigin").Attr("t", "i").CTag()
	x.CTag()

	x.CTag()

	return w.out.WriteBlob(abspath, bb.Bytes())
}

func (w *Writer) writeRichValueData() error {
	_, rid := w.nextWorkbookID()

	relpath := "richData/rdrichvalue.xml"
	abspath := "/xl/" + relpath

	w.PartContentTypes[abspath] = "application/vnd.ms-excel.rdrichvalue+xml"
	w.WorkbookRels[rid] = RelInfo{
		Type:   "http://schemas.microsoft.com/office/2017/06/relationships/rdRichValue",
		Target: relpath,
	}

	bb := bytes.Buffer{}
	x := xml.NewWriter(&bb, xml.WriterConfig{Indent: xml.Indent2Spaces})
	x.XmlStandaloneDecl()

	x.OTag("rvData")

	x.Attr("xmlns", "http://schemas.microsoft.com/office/spreadsheetml/2017/richdata")
	x.Attr("count", len(w.media))

	for _, m := range w.media {
		x.OTag("+rv").Attr("s", 0)
		x.OTag("v").Write(m.IId).CTag() // image resource numeric id
		x.OTag("v").Write(5).CTag()
		x.CTag()
	}

	x.CTag()

	return w.out.WriteBlob(abspath, bb.Bytes())
}

func (w *Writer) writeRichValueTypes() error {
	_, rid := w.nextWorkbookID()

	relpath := "richData/rdRichValueTypes.xml"
	abspath := "/xl/" + relpath

	w.PartContentTypes[abspath] = "application/vnd.ms-excel.rdrichvaluetypes+xml"
	w.WorkbookRels[rid] = RelInfo{
		Type:   "http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueTypes",
		Target: relpath,
	}

	bb := bytes.Buffer{}
	x := xml.NewWriter(&bb, xml.WriterConfig{Indent: xml.Indent2Spaces})
	x.XmlStandaloneDecl()

	x.OTag("rvTypesInfo")
	x.Attr("xmlns", "http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2")
	x.Attr("xmlns:mc", "http://schemas.openxmlformats.org/markup-compatibility/2006")
	x.Attr("xmlns:x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
	x.Attr("mc:Ignorable", "x")

	x.OTag("global")

	x.OTag("+key").Attr("name", "_Self")
	x.OTag("+flag").Attr("name", "ExcludeFromFile").Attr("value", 1).CTag()
	x.OTag("+flag").Attr("name", "ExcludeFromCalcComparison").Attr("value", 1).CTag()
	x.CTag()

	for _, s := range []string{
		"_DisplayString", "_Flags", "_Format", "_SubLabel", "_Attribution",
		"_Icon", "_Display", "_CanonicalPropertyNames", "_ClassificationId"} {

		x.OTag("+key").Attr("name", s)
		x.OTag("+flag").Attr("name", "ExcludeFromCalcComparison").Attr("value", 1).CTag()
		x.CTag()
	}

	x.CTag() // global

	x.CTag() // rvTypesInfo

	return w.out.WriteBlob(abspath, bb.Bytes())
}

func (w *Writer) writeRels(path string, rels map[string]RelInfo) error {
	bb := bytes.Buffer{}
	x := xml.NewWriter(&bb, xml.WriterConfig{Indent: xml.Indent2Spaces})
	x.XmlStandaloneDecl()

	x.OTag("Relationships")
	x.Attr("xmlns", "http://schemas.openxmlformats.org/package/2006/relationships")
	err := enumerate(rels, func(rid string, info RelInfo) error {
		x.OTag("+Relationship").Attr("Id", rid).Attr("Type", info.Type).Attr("Target", info.Target)
		x.CTag()

		return nil
	})
	if err != nil {
		return err
	}
	x.CTag()

	return w.out.WriteBlob(path, bb.Bytes())
}

func enumerate[M ~map[K]V, K constraints.Ordered, V any](m M, callback func(k K, v V) error) error {
	keys := maps.Keys(m)
	slices.Sort(keys)
	for _, k := range keys {
		err := callback(k, m[k])
		if err != nil {
			return err
		}
	}
	return nil
}
