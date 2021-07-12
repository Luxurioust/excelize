package excelize

import (
	"bytes"
	"compress/gzip"
	"encoding/xml"
	"fmt"
	"image/color"
	_ "image/gif"
	_ "image/jpeg"
	_ "image/png"
	"io/ioutil"
	"math"
	"os"
	"path/filepath"
	"strconv"
	"strings"
	"testing"
	"time"

	"github.com/stretchr/testify/assert"
)

func TestOpenFile(t *testing.T) {
	// Test update the spreadsheet file.
	f, err := OpenFile(filepath.Join("test", "Book1.xlsx"))
	assert.NoError(t, err)

	// Test get all the rows in a not exists worksheet.
	_, err = f.GetRows("Sheet4")
	assert.EqualError(t, err, "sheet Sheet4 is not exist")
	// Test get all the rows in a worksheet.
	rows, err := f.GetRows("Sheet2")
	assert.NoError(t, err)
	for _, row := range rows {
		for _, cell := range row {
			t.Log(cell, "\t")
		}
		t.Log("\r\n")
	}
	assert.NoError(t, f.UpdateLinkedValue())

	assert.NoError(t, f.SetCellDefault("Sheet2", "A1", strconv.FormatFloat(float64(100.1588), 'f', -1, 32)))
	assert.NoError(t, f.SetCellDefault("Sheet2", "A1", strconv.FormatFloat(float64(-100.1588), 'f', -1, 64)))

	// Test set cell value with illegal row number.
	assert.EqualError(t, f.SetCellDefault("Sheet2", "A", strconv.FormatFloat(float64(-100.1588), 'f', -1, 64)),
		`cannot convert cell "A" to coordinates: invalid cell name "A"`)

	assert.NoError(t, f.SetCellInt("Sheet2", "A1", 100))

	// Test set cell integer value with illegal row number.
	assert.EqualError(t, f.SetCellInt("Sheet2", "A", 100), `cannot convert cell "A" to coordinates: invalid cell name "A"`)

	assert.NoError(t, f.SetCellStr("Sheet2", "C11", "Knowns"))
	// Test max characters in a cell.
	assert.NoError(t, f.SetCellStr("Sheet2", "D11", strings.Repeat("c", TotalCellChars+2)))
	f.NewSheet(":\\/?*[]Maximum 31 characters allowed in sheet title.")
	// Test set worksheet name with illegal name.
	f.SetSheetName("Maximum 31 characters allowed i", "[Rename]:\\/?* Maximum 31 characters allowed in sheet title.")
	assert.EqualError(t, f.SetCellInt("Sheet3", "A23", 10), "sheet Sheet3 is not exist")
	assert.EqualError(t, f.SetCellStr("Sheet3", "b230", "10"), "sheet Sheet3 is not exist")
	assert.EqualError(t, f.SetCellStr("Sheet10", "b230", "10"), "sheet Sheet10 is not exist")

	// Test set cell string value with illegal row number.
	assert.EqualError(t, f.SetCellStr("Sheet1", "A", "10"), `cannot convert cell "A" to coordinates: invalid cell name "A"`)

	f.SetActiveSheet(2)
	// Test get cell formula with given rows number.
	_, err = f.GetCellFormula("Sheet1", "B19")
	assert.NoError(t, err)
	// Test get cell formula with illegal worksheet name.
	_, err = f.GetCellFormula("Sheet2", "B20")
	assert.NoError(t, err)
	_, err = f.GetCellFormula("Sheet1", "B20")
	assert.NoError(t, err)

	// Test get cell formula with illegal rows number.
	_, err = f.GetCellFormula("Sheet1", "B")
	assert.EqualError(t, err, `cannot convert cell "B" to coordinates: invalid cell name "B"`)
	// Test get shared cell formula
	_, err = f.GetCellFormula("Sheet2", "H11")
	assert.NoError(t, err)
	_, err = f.GetCellFormula("Sheet2", "I11")
	assert.NoError(t, err)
	getSharedForumula(&xlsxWorksheet{}, "", "")

	// Test read cell value with given illegal rows number.
	_, err = f.GetCellValue("Sheet2", "a-1")
	assert.EqualError(t, err, `cannot convert cell "A-1" to coordinates: invalid cell name "A-1"`)
	_, err = f.GetCellValue("Sheet2", "A")
	assert.EqualError(t, err, `cannot convert cell "A" to coordinates: invalid cell name "A"`)

	// Test read cell value with given lowercase column number.
	_, err = f.GetCellValue("Sheet2", "a5")
	assert.NoError(t, err)
	_, err = f.GetCellValue("Sheet2", "C11")
	assert.NoError(t, err)
	_, err = f.GetCellValue("Sheet2", "D11")
	assert.NoError(t, err)
	_, err = f.GetCellValue("Sheet2", "D12")
	assert.NoError(t, err)
	// Test SetCellValue function.
	assert.NoError(t, f.SetCellValue("Sheet2", "F1", " Hello"))
	assert.NoError(t, f.SetCellValue("Sheet2", "G1", []byte("World")))
	assert.NoError(t, f.SetCellValue("Sheet2", "F2", 42))
	assert.NoError(t, f.SetCellValue("Sheet2", "F3", int8(1<<8/2-1)))
	assert.NoError(t, f.SetCellValue("Sheet2", "F4", int16(1<<16/2-1)))
	assert.NoError(t, f.SetCellValue("Sheet2", "F5", int32(1<<32/2-1)))
	assert.NoError(t, f.SetCellValue("Sheet2", "F6", int64(1<<32/2-1)))
	assert.NoError(t, f.SetCellValue("Sheet2", "F7", float32(42.65418)))
	assert.NoError(t, f.SetCellValue("Sheet2", "F8", float64(-42.65418)))
	assert.NoError(t, f.SetCellValue("Sheet2", "F9", float32(42)))
	assert.NoError(t, f.SetCellValue("Sheet2", "F10", float64(42)))
	assert.NoError(t, f.SetCellValue("Sheet2", "F11", uint(1<<32-1)))
	assert.NoError(t, f.SetCellValue("Sheet2", "F12", uint8(1<<8-1)))
	assert.NoError(t, f.SetCellValue("Sheet2", "F13", uint16(1<<16-1)))
	assert.NoError(t, f.SetCellValue("Sheet2", "F14", uint32(1<<32-1)))
	assert.NoError(t, f.SetCellValue("Sheet2", "F15", uint64(1<<32-1)))
	assert.NoError(t, f.SetCellValue("Sheet2", "F16", true))
	assert.NoError(t, f.SetCellValue("Sheet2", "F17", complex64(5+10i)))

	// Test on not exists worksheet.
	assert.EqualError(t, f.SetCellDefault("SheetN", "A1", ""), "sheet SheetN is not exist")
	assert.EqualError(t, f.SetCellFloat("SheetN", "A1", 42.65418, 2, 32), "sheet SheetN is not exist")
	assert.EqualError(t, f.SetCellBool("SheetN", "A1", true), "sheet SheetN is not exist")
	assert.EqualError(t, f.SetCellFormula("SheetN", "A1", ""), "sheet SheetN is not exist")
	assert.EqualError(t, f.SetCellHyperLink("SheetN", "A1", "Sheet1!A40", "Location"), "sheet SheetN is not exist")

	// Test boolean write
	booltest := []struct {
		value    bool
		expected string
	}{
		{false, "0"},
		{true, "1"},
	}
	for _, test := range booltest {
		assert.NoError(t, f.SetCellValue("Sheet2", "F16", test.value))
		val, err := f.GetCellValue("Sheet2", "F16")
		assert.NoError(t, err)
		assert.Equal(t, test.expected, val)
	}

	assert.NoError(t, f.SetCellValue("Sheet2", "G2", nil))

	assert.NoError(t, f.SetCellValue("Sheet2", "G4", time.Now()))

	assert.NoError(t, f.SetCellValue("Sheet2", "G4", time.Now().UTC()))
	// 02:46:40
	assert.NoError(t, f.SetCellValue("Sheet2", "G5", time.Duration(1e13)))
	// Test completion column.
	assert.NoError(t, f.SetCellValue("Sheet2", "M2", nil))
	// Test read cell value with given axis large than exists row.
	_, err = f.GetCellValue("Sheet2", "E231")
	assert.NoError(t, err)
	// Test get active worksheet of spreadsheet and get worksheet name of spreadsheet by given worksheet index.
	f.GetSheetName(f.GetActiveSheetIndex())
	// Test get worksheet index of spreadsheet by given worksheet name.
	f.GetSheetIndex("Sheet1")
	// Test get worksheet name of spreadsheet by given invalid worksheet index.
	f.GetSheetName(4)
	// Test get worksheet map of workbook.
	f.GetSheetMap()
	for i := 1; i <= 300; i++ {
		assert.NoError(t, f.SetCellStr("Sheet2", "c"+strconv.Itoa(i), strconv.Itoa(i)))
	}
	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestOpenFile.xlsx")))
	assert.EqualError(t, f.SaveAs(filepath.Join("test", strings.Repeat("c", 199), ".xlsx")), ErrMaxFileNameLength.Error())
}

func TestSaveFile(t *testing.T) {
	f, err := OpenFile(filepath.Join("test", "Book1.xlsx"))
	if !assert.NoError(t, err) {
		t.FailNow()
	}
	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestSaveFile.xlsx")))
	f, err = OpenFile(filepath.Join("test", "TestSaveFile.xlsx"))
	if !assert.NoError(t, err) {
		t.FailNow()
	}
	assert.NoError(t, f.Save())
}

func TestSaveAsWrongPath(t *testing.T) {
	f, err := OpenFile(filepath.Join("test", "Book1.xlsx"))
	if assert.NoError(t, err) {
		// Test write file to not exist directory.
		err = f.SaveAs("")
		if assert.Error(t, err) {
			assert.True(t, os.IsNotExist(err), "Error: %v: Expected os.IsNotExists(err) == true", err)
		}
	}
}

func TestCharsetTranscoder(t *testing.T) {
	f := NewFile()
	f.CharsetTranscoder(*new(charsetTranscoderFn))
}

func TestOpenReader(t *testing.T) {
	_, err := OpenReader(strings.NewReader(""))
	assert.EqualError(t, err, "zip: not a valid zip file")
	_, err = OpenReader(bytes.NewReader(oleIdentifier), Options{Password: "password"})
	assert.EqualError(t, err, "decrypted file failed")

	// Test open password protected spreadsheet created by Microsoft Office Excel 2010.
	f, err := OpenFile(filepath.Join("test", "encryptSHA1.xlsx"), Options{Password: "password"})
	assert.NoError(t, err)
	val, err := f.GetCellValue("Sheet1", "A1")
	assert.NoError(t, err)
	assert.Equal(t, "SECRET", val)

	// Test open password protected spreadsheet created by LibreOffice 7.0.0.3.
	f, err = OpenFile(filepath.Join("test", "encryptAES.xlsx"), Options{Password: "password"})
	assert.NoError(t, err)
	val, err = f.GetCellValue("Sheet1", "A1")
	assert.NoError(t, err)
	assert.Equal(t, "SECRET", val)

	// Test unexpected EOF.
	var b bytes.Buffer
	w := gzip.NewWriter(&b)
	defer w.Close()
	w.Flush()

	r, _ := gzip.NewReader(&b)
	defer r.Close()

	_, err = OpenReader(r)
	assert.EqualError(t, err, "unexpected EOF")

	_, err = OpenReader(bytes.NewReader([]byte{
		0x50, 0x4b, 0x03, 0x04, 0x0a, 0x00, 0x09, 0x00, 0x63, 0x00, 0x47, 0xa3, 0xb6, 0x50, 0x00, 0x00,
		0x00, 0x00, 0x1c, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x08, 0x00, 0x0b, 0x00, 0x70, 0x61,
		0x73, 0x73, 0x77, 0x6f, 0x72, 0x64, 0x01, 0x99, 0x07, 0x00, 0x02, 0x00, 0x41, 0x45, 0x03, 0x00,
		0x00, 0x21, 0x06, 0x59, 0xc0, 0x12, 0xf3, 0x19, 0xc7, 0x51, 0xd1, 0xc9, 0x31, 0xcb, 0xcc, 0x8a,
		0xe1, 0x44, 0xe1, 0x56, 0x20, 0x24, 0x1f, 0xba, 0x09, 0xda, 0x53, 0xd5, 0xef, 0x50, 0x4b, 0x07,
		0x08, 0x00, 0x00, 0x00, 0x00, 0x1c, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x50, 0x4b, 0x01,
		0x02, 0x1f, 0x00, 0x0a, 0x00, 0x09, 0x00, 0x63, 0x00, 0x47, 0xa3, 0xb6, 0x50, 0x00, 0x00, 0x00,
		0x00, 0x1c, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x08, 0x00, 0x0b, 0x00, 0x00, 0x00, 0x00,
		0x00, 0x00, 0x00, 0x20, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x70, 0x61, 0x73, 0x73, 0x77,
		0x6f, 0x72, 0x64, 0x01, 0x99, 0x07, 0x00, 0x02, 0x00, 0x41, 0x45, 0x03, 0x00, 0x00, 0x50, 0x4b,
		0x05, 0x06, 0x00, 0x00, 0x00, 0x00, 0x01, 0x00, 0x01, 0x00, 0x41, 0x00, 0x00, 0x00, 0x5d, 0x00,
		0x00, 0x00, 0x00, 0x00,
	}))
	assert.EqualError(t, err, "zip: unsupported compression algorithm")
}

func TestBrokenFile(t *testing.T) {
	// Test write file with broken file struct.
	f := File{}

	t.Run("SaveWithoutName", func(t *testing.T) {
		assert.EqualError(t, f.Save(), "no path defined for file, consider File.WriteTo or File.Write")
	})

	t.Run("SaveAsEmptyStruct", func(t *testing.T) {
		// Test write file with broken file struct with given path.
		assert.NoError(t, f.SaveAs(filepath.Join("test", "BadWorkbook.SaveAsEmptyStruct.xlsx")))
	})

	t.Run("OpenBadWorkbook", func(t *testing.T) {
		// Test set active sheet without BookViews and Sheets maps in xl/workbook.xml.
		f3, err := OpenFile(filepath.Join("test", "BadWorkbook.xlsx"))
		f3.GetActiveSheetIndex()
		f3.SetActiveSheet(1)
		assert.NoError(t, err)
	})

	t.Run("OpenNotExistsFile", func(t *testing.T) {
		// Test open a spreadsheet file with given illegal path.
		_, err := OpenFile(filepath.Join("test", "NotExistsFile.xlsx"))
		if assert.Error(t, err) {
			assert.True(t, os.IsNotExist(err), "Expected os.IsNotExists(err) == true")
		}
	})
}

func TestNewFile(t *testing.T) {
	// Test create a spreadsheet file.
	f := NewFile()
	f.NewSheet("Sheet1")
	f.NewSheet("XLSXSheet2")
	f.NewSheet("XLSXSheet3")
	assert.NoError(t, f.SetCellInt("XLSXSheet2", "A23", 56))
	assert.NoError(t, f.SetCellStr("Sheet1", "B20", "42"))
	f.SetActiveSheet(0)

	// Test add picture to sheet with scaling and positioning.
	err := f.AddPicture("Sheet1", "H2", filepath.Join("test", "images", "excel.gif"),
		`{"x_scale": 0.5, "y_scale": 0.5, "positioning": "absolute"}`)
	if !assert.NoError(t, err) {
		t.FailNow()
	}

	// Test add picture to worksheet without formatset.
	err = f.AddPicture("Sheet1", "C2", filepath.Join("test", "images", "excel.png"), "")
	if !assert.NoError(t, err) {
		t.FailNow()
	}

	// Test add picture to worksheet with invalid formatset.
	err = f.AddPicture("Sheet1", "C2", filepath.Join("test", "images", "excel.png"), `{`)
	if !assert.Error(t, err) {
		t.FailNow()
	}

	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestNewFile.xlsx")))
	assert.NoError(t, f.Save())
}

func TestAddDrawingVML(t *testing.T) {
	// Test addDrawingVML with illegal cell coordinates.
	f := NewFile()
	assert.EqualError(t, f.addDrawingVML(0, "", "*", 0, 0), `cannot convert cell "*" to coordinates: invalid cell name "*"`)
}

func TestSetCellHyperLink(t *testing.T) {
	f, err := OpenFile(filepath.Join("test", "Book1.xlsx"))
	if err != nil {
		t.Log(err)
	}
	// Test set cell hyperlink in a work sheet already have hyperlinks.
	assert.NoError(t, f.SetCellHyperLink("Sheet1", "B19", "https://github.com/360EntSecGroup-Skylar/excelize", "External"))
	// Test add first hyperlink in a work sheet.
	assert.NoError(t, f.SetCellHyperLink("Sheet2", "C1", "https://github.com/360EntSecGroup-Skylar/excelize", "External"))
	// Test add Location hyperlink in a work sheet.
	assert.NoError(t, f.SetCellHyperLink("Sheet2", "D6", "Sheet1!D8", "Location"))
	// Test add Location hyperlink with display & tooltip in a work sheet.
	display := "Display value"
	tooltip := "Hover text"
	assert.NoError(t, f.SetCellHyperLink("Sheet2", "D7", "Sheet1!D9", "Location", HyperlinkOpts{
		Display: &display,
		Tooltip: &tooltip,
	}))

	assert.EqualError(t, f.SetCellHyperLink("Sheet2", "C3", "Sheet1!D8", ""), `invalid link type ""`)

	assert.EqualError(t, f.SetCellHyperLink("Sheet2", "", "Sheet1!D60", "Location"), `invalid cell name ""`)

	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestSetCellHyperLink.xlsx")))

	f = NewFile()
	_, err = f.workSheetReader("Sheet1")
	assert.NoError(t, err)
	ws, ok := f.Sheet.Load("xl/worksheets/sheet1.xml")
	assert.True(t, ok)
	ws.(*xlsxWorksheet).Hyperlinks = &xlsxHyperlinks{Hyperlink: make([]xlsxHyperlink, 65530)}
	assert.EqualError(t, f.SetCellHyperLink("Sheet1", "A65531", "https://github.com/360EntSecGroup-Skylar/excelize", "External"), ErrTotalSheetHyperlinks.Error())

	f = NewFile()
	_, err = f.workSheetReader("Sheet1")
	assert.NoError(t, err)
	ws, ok = f.Sheet.Load("xl/worksheets/sheet1.xml")
	assert.True(t, ok)
	ws.(*xlsxWorksheet).MergeCells = &xlsxMergeCells{Cells: []*xlsxMergeCell{{Ref: "A:A"}}}
	err = f.SetCellHyperLink("Sheet1", "A1", "https://github.com/360EntSecGroup-Skylar/excelize", "External")
	assert.EqualError(t, err, `cannot convert cell "A" to coordinates: invalid cell name "A"`)
}

func TestGetCellHyperLink(t *testing.T) {
	f, err := OpenFile(filepath.Join("test", "Book1.xlsx"))
	if !assert.NoError(t, err) {
		t.FailNow()
	}

	_, _, err = f.GetCellHyperLink("Sheet1", "")
	assert.EqualError(t, err, `invalid cell name ""`)

	link, target, err := f.GetCellHyperLink("Sheet1", "A22")
	assert.NoError(t, err)
	t.Log(link, target)
	link, target, err = f.GetCellHyperLink("Sheet2", "D6")
	assert.NoError(t, err)
	t.Log(link, target)
	link, target, err = f.GetCellHyperLink("Sheet3", "H3")
	assert.EqualError(t, err, "sheet Sheet3 is not exist")
	t.Log(link, target)

	f = NewFile()
	_, err = f.workSheetReader("Sheet1")
	assert.NoError(t, err)
	ws, ok := f.Sheet.Load("xl/worksheets/sheet1.xml")
	assert.True(t, ok)
	ws.(*xlsxWorksheet).Hyperlinks = &xlsxHyperlinks{
		Hyperlink: []xlsxHyperlink{{Ref: "A1"}},
	}
	link, target, err = f.GetCellHyperLink("Sheet1", "A1")
	assert.NoError(t, err)
	assert.Equal(t, link, true)
	assert.Equal(t, target, "")

	ws, ok = f.Sheet.Load("xl/worksheets/sheet1.xml")
	assert.True(t, ok)
	ws.(*xlsxWorksheet).MergeCells = &xlsxMergeCells{Cells: []*xlsxMergeCell{{Ref: "A:A"}}}
	link, target, err = f.GetCellHyperLink("Sheet1", "A1")
	assert.EqualError(t, err, `cannot convert cell "A" to coordinates: invalid cell name "A"`)
	assert.Equal(t, link, false)
	assert.Equal(t, target, "")

}

func TestSetCellFormula(t *testing.T) {
	f, err := OpenFile(filepath.Join("test", "Book1.xlsx"))
	if !assert.NoError(t, err) {
		t.FailNow()
	}

	assert.NoError(t, f.SetCellFormula("Sheet1", "B19", "SUM(Sheet2!D2,Sheet2!D11)"))
	assert.NoError(t, f.SetCellFormula("Sheet1", "C19", "SUM(Sheet2!D2,Sheet2!D9)"))

	// Test set cell formula with illegal rows number.
	assert.EqualError(t, f.SetCellFormula("Sheet1", "C", "SUM(Sheet2!D2,Sheet2!D9)"), `cannot convert cell "C" to coordinates: invalid cell name "C"`)

	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestSetCellFormula1.xlsx")))

	f, err = OpenFile(filepath.Join("test", "CalcChain.xlsx"))
	if !assert.NoError(t, err) {
		t.FailNow()
	}
	// Test remove cell formula.
	assert.NoError(t, f.SetCellFormula("Sheet1", "A1", ""))
	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestSetCellFormula2.xlsx")))
	// Test remove all cell formula.
	assert.NoError(t, f.SetCellFormula("Sheet1", "B1", ""))
	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestSetCellFormula3.xlsx")))
}

func TestSetSheetBackground(t *testing.T) {
	f, err := OpenFile(filepath.Join("test", "Book1.xlsx"))
	if !assert.NoError(t, err) {
		t.FailNow()
	}

	err = f.SetSheetBackground("Sheet2", filepath.Join("test", "images", "background.jpg"))
	if !assert.NoError(t, err) {
		t.FailNow()
	}

	err = f.SetSheetBackground("Sheet2", filepath.Join("test", "images", "background.jpg"))
	if !assert.NoError(t, err) {
		t.FailNow()
	}

	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestSetSheetBackground.xlsx")))
}

func TestSetSheetBackgroundErrors(t *testing.T) {
	f, err := OpenFile(filepath.Join("test", "Book1.xlsx"))
	if !assert.NoError(t, err) {
		t.FailNow()
	}

	err = f.SetSheetBackground("Sheet2", filepath.Join("test", "not_exists", "not_exists.png"))
	if assert.Error(t, err) {
		assert.True(t, os.IsNotExist(err), "Expected os.IsNotExists(err) == true")
	}

	err = f.SetSheetBackground("Sheet2", filepath.Join("test", "Book1.xlsx"))
	assert.EqualError(t, err, ErrImgExt.Error())
}

// TestWriteArrayFormula tests the extended options of SetCellFormula by writing an array function
// to a workbook. In the resulting file, the lines 2 and 3 as well as 4 and 5 should have matching
// contents.
func TestWriteArrayFormula(t *testing.T) {
	cell := func(col, row int) string {
		c, err := CoordinatesToCellName(col, row)
		if err != nil {
			t.Fatal(err)
		}

		return c
	}

	f := NewFile()

	sample := []string{"Sample 1", "Sample 2", "Sample 3"}
	values := []int{1855, 1709, 1462, 1115, 1524, 625, 773, 126, 1027, 1696, 1078, 1917, 1109, 1753, 1884, 659, 994, 1911, 1925, 899, 196, 244, 1488, 1056, 1986, 66, 784, 725, 767, 1722, 1541, 1026, 1455, 264, 1538, 877, 1581, 1098, 383, 762, 237, 493, 29, 1923, 474, 430, 585, 688, 308, 200, 1259, 622, 798, 1048, 996, 601, 582, 332, 377, 805, 250, 1860, 1360, 840, 911, 1346, 1651, 1651, 665, 584, 1057, 1145, 925, 1752, 202, 149, 1917, 1398, 1894, 818, 714, 624, 1085, 1566, 635, 78, 313, 1686, 1820, 494, 614, 1913, 271, 1016, 338, 1301, 489, 1733, 1483, 1141}
	assoc := []int{2, 0, 0, 0, 0, 1, 1, 0, 0, 1, 2, 2, 2, 1, 1, 1, 1, 0, 0, 0, 1, 0, 2, 0, 2, 1, 2, 2, 2, 1, 0, 1, 0, 1, 1, 2, 0, 2, 1, 0, 2, 1, 0, 1, 0, 0, 2, 0, 2, 2, 1, 2, 2, 1, 2, 2, 1, 2, 1, 2, 2, 1, 1, 1, 0, 1, 0, 2, 0, 0, 1, 2, 1, 0, 1, 0, 0, 2, 1, 1, 2, 0, 2, 1, 0, 2, 2, 2, 1, 0, 0, 1, 1, 1, 2, 0, 2, 0, 1, 1}
	if len(values) != len(assoc) {
		t.Fatal("values and assoc must be of same length")
	}

	// Average calculates the average of the n-th sample (0 <= n < len(sample)).
	average := func(n int) int {
		sum := 0
		count := 0
		for i := 0; i != len(values); i++ {
			if assoc[i] == n {
				sum += values[i]
				count++
			}
		}

		return int(math.Round(float64(sum) / float64(count)))
	}

	// Stdev calculates the standard deviation of the n-th sample (0 <= n < len(sample)).
	stdev := func(n int) int {
		avg := average(n)

		sum := 0
		count := 0
		for i := 0; i != len(values); i++ {
			if assoc[i] == n {
				sum += (values[i] - avg) * (values[i] - avg)
				count++
			}
		}

		return int(math.Round(math.Sqrt(float64(sum) / float64(count))))
	}

	// Line 2 contains the results of AVERAGEIF
	assert.NoError(t, f.SetCellStr("Sheet1", "A2", "Average"))

	// Line 3 contains the average that was calculated in Go
	assert.NoError(t, f.SetCellStr("Sheet1", "A3", "Average (calculated)"))

	// Line 4 contains the results of the array function that calculates the standard deviation
	assert.NoError(t, f.SetCellStr("Sheet1", "A4", "Std. deviation"))

	// Line 5 contains the standard deviations calculated in Go
	assert.NoError(t, f.SetCellStr("Sheet1", "A5", "Std. deviation (calculated)"))

	assert.NoError(t, f.SetCellStr("Sheet1", "B1", sample[0]))
	assert.NoError(t, f.SetCellStr("Sheet1", "C1", sample[1]))
	assert.NoError(t, f.SetCellStr("Sheet1", "D1", sample[2]))

	firstResLine := 8
	assert.NoError(t, f.SetCellStr("Sheet1", cell(1, firstResLine-1), "Result Values"))
	assert.NoError(t, f.SetCellStr("Sheet1", cell(2, firstResLine-1), "Sample"))

	for i := 0; i != len(values); i++ {
		valCell := cell(1, i+firstResLine)
		assocCell := cell(2, i+firstResLine)

		assert.NoError(t, f.SetCellInt("Sheet1", valCell, values[i]))
		assert.NoError(t, f.SetCellStr("Sheet1", assocCell, sample[assoc[i]]))
	}

	valRange := fmt.Sprintf("$A$%d:$A$%d", firstResLine, len(values)+firstResLine-1)
	assocRange := fmt.Sprintf("$B$%d:$B$%d", firstResLine, len(values)+firstResLine-1)

	for i := 0; i != len(sample); i++ {
		nameCell := cell(i+2, 1)
		avgCell := cell(i+2, 2)
		calcAvgCell := cell(i+2, 3)
		stdevCell := cell(i+2, 4)
		calcStdevCell := cell(i+2, 5)

		assert.NoError(t, f.SetCellInt("Sheet1", calcAvgCell, average(i)))
		assert.NoError(t, f.SetCellInt("Sheet1", calcStdevCell, stdev(i)))

		// Average can be done with AVERAGEIF
		assert.NoError(t, f.SetCellFormula("Sheet1", avgCell, fmt.Sprintf("ROUND(AVERAGEIF(%s,%s,%s),0)", assocRange, nameCell, valRange)))

		ref := stdevCell + ":" + stdevCell
		arr := STCellFormulaTypeArray
		// Use an array formula for standard deviation
		assert.NoError(t, f.SetCellFormula("Sheet1", stdevCell, fmt.Sprintf("ROUND(STDEVP(IF(%s=%s,%s)),0)", assocRange, nameCell, valRange),
			FormulaOpts{}, FormulaOpts{Type: &arr}, FormulaOpts{Ref: &ref}))
	}

	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestWriteArrayFormula.xlsx")))
}

func TestSetCellStyleAlignment(t *testing.T) {
	f, err := prepareTestBook1()
	if !assert.NoError(t, err) {
		t.FailNow()
	}

	var style int
	style, err = f.NewStyle(`{"alignment":{"horizontal":"center","ident":1,"justify_last_line":true,"reading_order":0,"relative_indent":1,"shrink_to_fit":true,"text_rotation":45,"vertical":"top","wrap_text":true}}`)
	if !assert.NoError(t, err) {
		t.FailNow()
	}

	assert.NoError(t, f.SetCellStyle("Sheet1", "A22", "A22", style))

	// Test set cell style with given illegal rows number.
	assert.EqualError(t, f.SetCellStyle("Sheet1", "A", "A22", style), `cannot convert cell "A" to coordinates: invalid cell name "A"`)
	assert.EqualError(t, f.SetCellStyle("Sheet1", "A22", "A", style), `cannot convert cell "A" to coordinates: invalid cell name "A"`)

	// Test get cell style with given illegal rows number.
	index, err := f.GetCellStyle("Sheet1", "A")
	assert.Equal(t, 0, index)
	assert.EqualError(t, err, `cannot convert cell "A" to coordinates: invalid cell name "A"`)

	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestSetCellStyleAlignment.xlsx")))
}

func TestSetCellStyleBorder(t *testing.T) {
	f, err := prepareTestBook1()
	if !assert.NoError(t, err) {
		t.FailNow()
	}

	var style int

	// Test set border on overlapping area with vertical variants shading styles gradient fill.
	style, err = f.NewStyle(`{"border":[{"type":"left","color":"0000FF","style":2},{"type":"top","color":"00FF00","style":12},{"type":"bottom","color":"FFFF00","style":5},{"type":"right","color":"FF0000","style":6},{"type":"diagonalDown","color":"A020F0","style":9},{"type":"diagonalUp","color":"A020F0","style":8}]}`)
	if !assert.NoError(t, err) {
		t.FailNow()
	}
	assert.NoError(t, f.SetCellStyle("Sheet1", "J21", "L25", style))

	style, err = f.NewStyle(`{"border":[{"type":"left","color":"0000FF","style":2},{"type":"top","color":"00FF00","style":3},{"type":"bottom","color":"FFFF00","style":4},{"type":"right","color":"FF0000","style":5},{"type":"diagonalDown","color":"A020F0","style":6},{"type":"diagonalUp","color":"A020F0","style":7}],"fill":{"type":"gradient","color":["#FFFFFF","#E0EBF5"],"shading":1}}`)
	if !assert.NoError(t, err) {
		t.FailNow()
	}
	assert.NoError(t, f.SetCellStyle("Sheet1", "M28", "K24", style))

	style, err = f.NewStyle(`{"border":[{"type":"left","color":"0000FF","style":2},{"type":"top","color":"00FF00","style":3},{"type":"bottom","color":"FFFF00","style":4},{"type":"right","color":"FF0000","style":5},{"type":"diagonalDown","color":"A020F0","style":6},{"type":"diagonalUp","color":"A020F0","style":7}],"fill":{"type":"gradient","color":["#FFFFFF","#E0EBF5"],"shading":4}}`)
	if !assert.NoError(t, err) {
		t.FailNow()
	}
	assert.NoError(t, f.SetCellStyle("Sheet1", "M28", "K24", style))

	// Test set border and solid style pattern fill for a single cell.
	style, err = f.NewStyle(&Style{
		Border: []Border{
			{
				Type:  "left",
				Color: "0000FF",
				Style: 8,
			},
			{
				Type:  "top",
				Color: "00FF00",
				Style: 9,
			},
			{
				Type:  "bottom",
				Color: "FFFF00",
				Style: 10,
			},
			{
				Type:  "right",
				Color: "FF0000",
				Style: 11,
			},
			{
				Type:  "diagonalDown",
				Color: "A020F0",
				Style: 12,
			},
			{
				Type:  "diagonalUp",
				Color: "A020F0",
				Style: 13,
			},
		},
		Fill: Fill{
			Type:    "pattern",
			Color:   []string{"#E0EBF5"},
			Pattern: 1,
		},
	})
	if !assert.NoError(t, err) {
		t.FailNow()
	}

	assert.NoError(t, f.SetCellStyle("Sheet1", "O22", "O22", style))

	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestSetCellStyleBorder.xlsx")))
}

func TestSetCellStyleBorderErrors(t *testing.T) {
	f, err := prepareTestBook1()
	if !assert.NoError(t, err) {
		t.FailNow()
	}

	// Set border with invalid style parameter.
	_, err = f.NewStyle("")
	if !assert.EqualError(t, err, "unexpected end of JSON input") {
		t.FailNow()
	}

	// Set border with invalid style index number.
	_, err = f.NewStyle(`{"border":[{"type":"left","color":"0000FF","style":-1},{"type":"top","color":"00FF00","style":14},{"type":"bottom","color":"FFFF00","style":5},{"type":"right","color":"FF0000","style":6},{"type":"diagonalDown","color":"A020F0","style":9},{"type":"diagonalUp","color":"A020F0","style":8}]}`)
	if !assert.NoError(t, err) {
		t.FailNow()
	}
}

func TestSetCellStyleNumberFormat(t *testing.T) {
	f, err := prepareTestBook1()
	if !assert.NoError(t, err) {
		t.FailNow()
	}

	// Test only set fill and number format for a cell.
	col := []string{"L", "M", "N", "O", "P"}
	data := []int{0, 1, 2, 3, 4, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 37, 38, 39, 40, 41, 42, 43, 44, 45, 46, 47, 48, 49}
	value := []string{"37947.7500001", "-37947.7500001", "0.007", "2.1", "String"}
	for i, v := range value {
		for k, d := range data {
			c := col[i] + strconv.Itoa(k+1)
			var val float64
			val, err = strconv.ParseFloat(v, 64)
			if err != nil {
				assert.NoError(t, f.SetCellValue("Sheet2", c, v))
			} else {
				assert.NoError(t, f.SetCellValue("Sheet2", c, val))
			}
			style, err := f.NewStyle(`{"fill":{"type":"gradient","color":["#FFFFFF","#E0EBF5"],"shading":5},"number_format": ` + strconv.Itoa(d) + `}`)
			if !assert.NoError(t, err) {
				t.FailNow()
			}
			assert.NoError(t, f.SetCellStyle("Sheet2", c, c, style))
			t.Log(f.GetCellValue("Sheet2", c))
		}
	}
	var style int
	style, err = f.NewStyle(`{"number_format":-1}`)
	if !assert.NoError(t, err) {
		t.FailNow()
	}
	assert.NoError(t, f.SetCellStyle("Sheet2", "L33", "L33", style))

	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestSetCellStyleNumberFormat.xlsx")))
}

func TestSetCellStyleCurrencyNumberFormat(t *testing.T) {
	t.Run("TestBook3", func(t *testing.T) {
		f, err := prepareTestBook3()
		if !assert.NoError(t, err) {
			t.FailNow()
		}

		assert.NoError(t, f.SetCellValue("Sheet1", "A1", 56))
		assert.NoError(t, f.SetCellValue("Sheet1", "A2", -32.3))
		var style int
		style, err = f.NewStyle(`{"number_format": 188, "decimal_places": -1}`)
		if !assert.NoError(t, err) {
			t.FailNow()
		}

		assert.NoError(t, f.SetCellStyle("Sheet1", "A1", "A1", style))
		style, err = f.NewStyle(`{"number_format": 188, "decimal_places": 31, "negred": true}`)
		if !assert.NoError(t, err) {
			t.FailNow()
		}

		assert.NoError(t, f.SetCellStyle("Sheet1", "A2", "A2", style))

		assert.NoError(t, f.SaveAs(filepath.Join("test", "TestSetCellStyleCurrencyNumberFormat.TestBook3.xlsx")))
	})

	t.Run("TestBook4", func(t *testing.T) {
		f, err := prepareTestBook4()
		if !assert.NoError(t, err) {
			t.FailNow()
		}
		assert.NoError(t, f.SetCellValue("Sheet1", "A1", 42920.5))
		assert.NoError(t, f.SetCellValue("Sheet1", "A2", 42920.5))

		_, err = f.NewStyle(`{"number_format": 26, "lang": "zh-tw"}`)
		if !assert.NoError(t, err) {
			t.FailNow()
		}

		style, err := f.NewStyle(`{"number_format": 27}`)
		if !assert.NoError(t, err) {
			t.FailNow()
		}

		assert.NoError(t, f.SetCellStyle("Sheet1", "A1", "A1", style))
		style, err = f.NewStyle(`{"number_format": 31, "lang": "ko-kr"}`)
		if !assert.NoError(t, err) {
			t.FailNow()
		}

		assert.NoError(t, f.SetCellStyle("Sheet1", "A2", "A2", style))

		style, err = f.NewStyle(`{"number_format": 71, "lang": "th-th"}`)
		if !assert.NoError(t, err) {
			t.FailNow()
		}
		assert.NoError(t, f.SetCellStyle("Sheet1", "A2", "A2", style))

		assert.NoError(t, f.SaveAs(filepath.Join("test", "TestSetCellStyleCurrencyNumberFormat.TestBook4.xlsx")))
	})
}

func TestSetCellStyleCustomNumberFormat(t *testing.T) {
	f := NewFile()
	assert.NoError(t, f.SetCellValue("Sheet1", "A1", 42920.5))
	assert.NoError(t, f.SetCellValue("Sheet1", "A2", 42920.5))
	style, err := f.NewStyle(`{"custom_number_format": "[$-380A]dddd\\,\\ dd\" de \"mmmm\" de \"yyyy;@"}`)
	assert.NoError(t, err)
	assert.NoError(t, f.SetCellStyle("Sheet1", "A1", "A1", style))
	style, err = f.NewStyle(`{"custom_number_format": "[$-380A]dddd\\,\\ dd\" de \"mmmm\" de \"yyyy;@","font":{"color":"#9A0511"}}`)
	assert.NoError(t, err)
	assert.NoError(t, f.SetCellStyle("Sheet1", "A2", "A2", style))

	_, err = f.NewStyle(`{"custom_number_format": "[$-380A]dddd\\,\\ dd\" de \"mmmm\" de \"yy;@"}`)
	assert.NoError(t, err)
	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestSetCellStyleCustomNumberFormat.xlsx")))
}

func TestSetCellStyleFill(t *testing.T) {
	f, err := prepareTestBook1()
	if !assert.NoError(t, err) {
		t.FailNow()
	}

	var style int
	// Test set fill for cell with invalid parameter.
	style, err = f.NewStyle(`{"fill":{"type":"gradient","color":["#FFFFFF","#E0EBF5"],"shading":6}}`)
	assert.NoError(t, err)
	assert.NoError(t, f.SetCellStyle("Sheet1", "O23", "O23", style))

	style, err = f.NewStyle(`{"fill":{"type":"gradient","color":["#FFFFFF"],"shading":1}}`)
	assert.NoError(t, err)
	assert.NoError(t, f.SetCellStyle("Sheet1", "O23", "O23", style))

	style, err = f.NewStyle(`{"fill":{"type":"pattern","color":[],"pattern":1}}`)
	assert.NoError(t, err)
	assert.NoError(t, f.SetCellStyle("Sheet1", "O23", "O23", style))

	style, err = f.NewStyle(`{"fill":{"type":"pattern","color":["#E0EBF5"],"pattern":19}}`)
	if !assert.NoError(t, err) {
		t.FailNow()
	}
	assert.NoError(t, f.SetCellStyle("Sheet1", "O23", "O23", style))

	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestSetCellStyleFill.xlsx")))
}

func TestSetCellStyleFont(t *testing.T) {
	f, err := prepareTestBook1()
	if !assert.NoError(t, err) {
		t.FailNow()
	}

	var style int
	style, err = f.NewStyle(`{"font":{"bold":true,"italic":true,"family":"Times New Roman","size":36,"color":"#777777","underline":"single"}}`)
	if !assert.NoError(t, err) {
		t.FailNow()
	}

	assert.NoError(t, f.SetCellStyle("Sheet2", "A1", "A1", style))

	style, err = f.NewStyle(`{"font":{"italic":true,"underline":"double"}}`)
	if !assert.NoError(t, err) {
		t.FailNow()
	}

	assert.NoError(t, f.SetCellStyle("Sheet2", "A2", "A2", style))

	style, err = f.NewStyle(`{"font":{"bold":true}}`)
	if !assert.NoError(t, err) {
		t.FailNow()
	}

	assert.NoError(t, f.SetCellStyle("Sheet2", "A3", "A3", style))

	style, err = f.NewStyle(`{"font":{"bold":true,"family":"","size":0,"color":"","underline":""}}`)
	if !assert.NoError(t, err) {
		t.FailNow()
	}

	assert.NoError(t, f.SetCellStyle("Sheet2", "A4", "A4", style))

	style, err = f.NewStyle(`{"font":{"color":"#777777","strike":true}}`)
	if !assert.NoError(t, err) {
		t.FailNow()
	}

	assert.NoError(t, f.SetCellStyle("Sheet2", "A5", "A5", style))

	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestSetCellStyleFont.xlsx")))
}

func TestSetCellStyleProtection(t *testing.T) {
	f, err := prepareTestBook1()
	if !assert.NoError(t, err) {
		t.FailNow()
	}

	var style int
	style, err = f.NewStyle(`{"protection":{"hidden":true, "locked":true}}`)
	if !assert.NoError(t, err) {
		t.FailNow()
	}

	assert.NoError(t, f.SetCellStyle("Sheet2", "A6", "A6", style))
	err = f.SaveAs(filepath.Join("test", "TestSetCellStyleProtection.xlsx"))
	if !assert.NoError(t, err) {
		t.FailNow()
	}
}

func TestSetDeleteSheet(t *testing.T) {
	t.Run("TestBook3", func(t *testing.T) {
		f, err := prepareTestBook3()
		if !assert.NoError(t, err) {
			t.FailNow()
		}

		f.DeleteSheet("XLSXSheet3")
		assert.NoError(t, f.SaveAs(filepath.Join("test", "TestSetDeleteSheet.TestBook3.xlsx")))
	})

	t.Run("TestBook4", func(t *testing.T) {
		f, err := prepareTestBook4()
		if !assert.NoError(t, err) {
			t.FailNow()
		}
		f.DeleteSheet("Sheet1")
		assert.EqualError(t, f.AddComment("Sheet1", "A1", ""), "unexpected end of JSON input")
		assert.NoError(t, f.AddComment("Sheet1", "A1", `{"author":"Excelize: ","text":"This is a comment."}`))
		assert.NoError(t, f.SaveAs(filepath.Join("test", "TestSetDeleteSheet.TestBook4.xlsx")))
	})
}

func TestSheetVisibility(t *testing.T) {
	f, err := prepareTestBook1()
	if !assert.NoError(t, err) {
		t.FailNow()
	}

	assert.NoError(t, f.SetSheetVisible("Sheet2", false))
	assert.NoError(t, f.SetSheetVisible("Sheet1", false))
	assert.NoError(t, f.SetSheetVisible("Sheet1", true))
	assert.Equal(t, true, f.GetSheetVisible("Sheet1"))

	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestSheetVisibility.xlsx")))
}

func TestCopySheet(t *testing.T) {
	f, err := prepareTestBook1()
	if !assert.NoError(t, err) {
		t.FailNow()
	}

	idx := f.NewSheet("CopySheet")
	assert.NoError(t, f.CopySheet(0, idx))

	assert.NoError(t, f.SetCellValue("CopySheet", "F1", "Hello"))
	val, err := f.GetCellValue("Sheet1", "F1")
	assert.NoError(t, err)
	assert.NotEqual(t, "Hello", val)

	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestCopySheet.xlsx")))
}

func TestCopySheetError(t *testing.T) {
	f, err := prepareTestBook1()
	if !assert.NoError(t, err) {
		t.FailNow()
	}

	assert.EqualError(t, f.copySheet(-1, -2), "sheet  is not exist")
	if !assert.EqualError(t, f.CopySheet(-1, -2), "invalid worksheet index") {
		t.FailNow()
	}

	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestCopySheetError.xlsx")))
}

func TestGetSheetComments(t *testing.T) {
	f := NewFile()
	assert.Equal(t, "", f.getSheetComments("sheet0"))
}

func TestSetSheetVisible(t *testing.T) {
	f := NewFile()
	f.WorkBook.Sheets.Sheet[0].Name = "SheetN"
	assert.EqualError(t, f.SetSheetVisible("Sheet1", false), "sheet SheetN is not exist")
}

func TestGetActiveSheetIndex(t *testing.T) {
	f := NewFile()
	f.WorkBook.BookViews = nil
	assert.Equal(t, 0, f.GetActiveSheetIndex())
}

func TestRelsWriter(t *testing.T) {
	f := NewFile()
	f.Relationships.Store("xl/worksheets/sheet/rels/sheet1.xml.rel", &xlsxRelationships{})
	f.relsWriter()
}

func TestGetSheetView(t *testing.T) {
	f := NewFile()
	_, err := f.getSheetView("SheetN", 0)
	assert.EqualError(t, err, "sheet SheetN is not exist")
}

func TestConditionalFormat(t *testing.T) {
	f := NewFile()
	sheet1 := f.GetSheetName(0)

	fillCells(f, sheet1, 10, 15)

	var format1, format2, format3, format4 int
	var err error
	// Rose format for bad conditional.
	format1, err = f.NewConditionalStyle(`{"font":{"color":"#9A0511"},"fill":{"type":"pattern","color":["#FEC7CE"],"pattern":1}}`)
	if !assert.NoError(t, err) {
		t.FailNow()
	}

	// Light yellow format for neutral conditional.
	format2, err = f.NewConditionalStyle(`{"fill":{"type":"pattern","color":["#FEEAA0"],"pattern":1}}`)
	if !assert.NoError(t, err) {
		t.FailNow()
	}

	// Light green format for good conditional.
	format3, err = f.NewConditionalStyle(`{"font":{"color":"#09600B"},"fill":{"type":"pattern","color":["#C7EECF"],"pattern":1}}`)
	if !assert.NoError(t, err) {
		t.FailNow()
	}

	// conditional style with align and left border.
	format4, err = f.NewConditionalStyle(`{"alignment":{"wrap_text":true},"border":[{"type":"left","color":"#000000","style":1}]}`)
	if !assert.NoError(t, err) {
		t.FailNow()
	}

	// Color scales: 2 color.
	assert.NoError(t, f.SetConditionalFormat(sheet1, "A1:A10", `[{"type":"2_color_scale","criteria":"=","min_type":"min","max_type":"max","min_color":"#F8696B","max_color":"#63BE7B"}]`))
	// Color scales: 3 color.
	assert.NoError(t, f.SetConditionalFormat(sheet1, "B1:B10", `[{"type":"3_color_scale","criteria":"=","min_type":"min","mid_type":"percentile","max_type":"max","min_color":"#F8696B","mid_color":"#FFEB84","max_color":"#63BE7B"}]`))
	// Hightlight cells rules: between...
	assert.NoError(t, f.SetConditionalFormat(sheet1, "C1:C10", fmt.Sprintf(`[{"type":"cell","criteria":"between","format":%d,"minimum":"6","maximum":"8"}]`, format1)))
	// Hightlight cells rules: Greater Than...
	assert.NoError(t, f.SetConditionalFormat(sheet1, "D1:D10", fmt.Sprintf(`[{"type":"cell","criteria":">","format":%d,"value":"6"}]`, format3)))
	// Hightlight cells rules: Equal To...
	assert.NoError(t, f.SetConditionalFormat(sheet1, "E1:E10", fmt.Sprintf(`[{"type":"top","criteria":"=","format":%d}]`, format3)))
	// Hightlight cells rules: Not Equal To...
	assert.NoError(t, f.SetConditionalFormat(sheet1, "F1:F10", fmt.Sprintf(`[{"type":"unique","criteria":"=","format":%d}]`, format2)))
	// Hightlight cells rules: Duplicate Values...
	assert.NoError(t, f.SetConditionalFormat(sheet1, "G1:G10", fmt.Sprintf(`[{"type":"duplicate","criteria":"=","format":%d}]`, format2)))
	// Top/Bottom rules: Top 10%.
	assert.NoError(t, f.SetConditionalFormat(sheet1, "H1:H10", fmt.Sprintf(`[{"type":"top","criteria":"=","format":%d,"value":"6","percent":true}]`, format1)))
	// Top/Bottom rules: Above Average...
	assert.NoError(t, f.SetConditionalFormat(sheet1, "I1:I10", fmt.Sprintf(`[{"type":"average","criteria":"=","format":%d, "above_average": true}]`, format3)))
	// Top/Bottom rules: Below Average...
	assert.NoError(t, f.SetConditionalFormat(sheet1, "J1:J10", fmt.Sprintf(`[{"type":"average","criteria":"=","format":%d, "above_average": false}]`, format1)))
	// Data Bars: Gradient Fill.
	assert.NoError(t, f.SetConditionalFormat(sheet1, "K1:K10", `[{"type":"data_bar", "criteria":"=", "min_type":"min","max_type":"max","bar_color":"#638EC6"}]`))
	// Use a formula to determine which cells to format.
	assert.NoError(t, f.SetConditionalFormat(sheet1, "L1:L10", fmt.Sprintf(`[{"type":"formula", "criteria":"L2<3", "format":%d}]`, format1)))
	// Alignment/Border cells rules.
	assert.NoError(t, f.SetConditionalFormat(sheet1, "M1:M10", fmt.Sprintf(`[{"type":"cell","criteria":">","format":%d,"value":"0"}]`, format4)))

	// Test set invalid format set in conditional format.
	assert.EqualError(t, f.SetConditionalFormat(sheet1, "L1:L10", ""), "unexpected end of JSON input")
	// Set conditional format on not exists worksheet.
	assert.EqualError(t, f.SetConditionalFormat("SheetN", "L1:L10", "[]"), "sheet SheetN is not exist")

	err = f.SaveAs(filepath.Join("test", "TestConditionalFormat.xlsx"))
	if !assert.NoError(t, err) {
		t.FailNow()
	}

	// Set conditional format with illegal valid type.
	assert.NoError(t, f.SetConditionalFormat(sheet1, "K1:K10", `[{"type":"", "criteria":"=", "min_type":"min","max_type":"max","bar_color":"#638EC6"}]`))
	// Set conditional format with illegal criteria type.
	assert.NoError(t, f.SetConditionalFormat(sheet1, "K1:K10", `[{"type":"data_bar", "criteria":"", "min_type":"min","max_type":"max","bar_color":"#638EC6"}]`))

	// Set conditional format with file without dxfs element shold not return error.
	f, err = OpenFile(filepath.Join("test", "Book1.xlsx"))
	if !assert.NoError(t, err) {
		t.FailNow()
	}

	_, err = f.NewConditionalStyle(`{"font":{"color":"#9A0511"},"fill":{"type":"pattern","color":["#FEC7CE"],"pattern":1}}`)
	if !assert.NoError(t, err) {
		t.FailNow()
	}
}

func TestConditionalFormatError(t *testing.T) {
	f := NewFile()
	sheet1 := f.GetSheetName(0)

	fillCells(f, sheet1, 10, 15)

	// Set conditional format with illegal JSON string should return error.
	_, err := f.NewConditionalStyle("")
	if !assert.EqualError(t, err, "unexpected end of JSON input") {
		t.FailNow()
	}
}

func TestSharedStrings(t *testing.T) {
	f, err := OpenFile(filepath.Join("test", "SharedStrings.xlsx"))
	if !assert.NoError(t, err) {
		t.FailNow()
	}
	rows, err := f.GetRows("Sheet1")
	if !assert.NoError(t, err) {
		t.FailNow()
	}
	assert.Equal(t, "A", rows[0][0])
	rows, err = f.GetRows("Sheet2")
	if !assert.NoError(t, err) {
		t.FailNow()
	}
	assert.Equal(t, "Test Weight (Kgs)", rows[0][0])
}

func TestSetSheetRow(t *testing.T) {
	f, err := OpenFile(filepath.Join("test", "Book1.xlsx"))
	if !assert.NoError(t, err) {
		t.FailNow()
	}

	assert.NoError(t, f.SetSheetRow("Sheet1", "B27", &[]interface{}{"cell", nil, int32(42), float64(42), time.Now().UTC()}))

	assert.EqualError(t, f.SetSheetRow("Sheet1", "", &[]interface{}{"cell", nil, 2}),
		`cannot convert cell "" to coordinates: invalid cell name ""`)

	assert.EqualError(t, f.SetSheetRow("Sheet1", "B27", []interface{}{}), ErrParameterInvalid.Error())
	assert.EqualError(t, f.SetSheetRow("Sheet1", "B27", &f), ErrParameterInvalid.Error())
	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestSetSheetRow.xlsx")))
}

func TestHSL(t *testing.T) {
	var hsl HSL
	r, g, b, a := hsl.RGBA()
	assert.Equal(t, uint32(0), r)
	assert.Equal(t, uint32(0), g)
	assert.Equal(t, uint32(0), b)
	assert.Equal(t, uint32(0xffff), a)
	assert.Equal(t, HSL{0, 0, 0}, hslModel(hsl))
	assert.Equal(t, HSL{0, 0, 0}, hslModel(color.Gray16{Y: uint16(1)}))
	R, G, B := HSLToRGB(0, 1, 0.4)
	assert.Equal(t, uint8(204), R)
	assert.Equal(t, uint8(0), G)
	assert.Equal(t, uint8(0), B)
	R, G, B = HSLToRGB(0, 1, 0.6)
	assert.Equal(t, uint8(255), R)
	assert.Equal(t, uint8(51), G)
	assert.Equal(t, uint8(51), B)
	assert.Equal(t, 0.0, hueToRGB(0, 0, -1))
	assert.Equal(t, 0.0, hueToRGB(0, 0, 2))
	assert.Equal(t, 0.0, hueToRGB(0, 0, 1.0/7))
	assert.Equal(t, 0.0, hueToRGB(0, 0, 0.4))
	assert.Equal(t, 0.0, hueToRGB(0, 0, 2.0/4))
	t.Log(RGBToHSL(255, 255, 0))
	h, s, l := RGBToHSL(0, 255, 255)
	assert.Equal(t, float64(0.5), h)
	assert.Equal(t, float64(1), s)
	assert.Equal(t, float64(0.5), l)
	t.Log(RGBToHSL(250, 100, 50))
	t.Log(RGBToHSL(50, 100, 250))
	t.Log(RGBToHSL(250, 50, 100))
}

func TestProtectSheet(t *testing.T) {
	f := NewFile()
	assert.NoError(t, f.ProtectSheet("Sheet1", nil))
	assert.NoError(t, f.ProtectSheet("Sheet1", &FormatSheetProtection{
		Password:      "password",
		EditScenarios: false,
	}))

	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestProtectSheet.xlsx")))
	// Test protect not exists worksheet.
	assert.EqualError(t, f.ProtectSheet("SheetN", nil), "sheet SheetN is not exist")
}

func TestUnprotectSheet(t *testing.T) {
	f, err := OpenFile(filepath.Join("test", "Book1.xlsx"))
	if !assert.NoError(t, err) {
		t.FailNow()
	}
	// Test unprotect not exists worksheet.
	assert.EqualError(t, f.UnprotectSheet("SheetN"), "sheet SheetN is not exist")

	assert.NoError(t, f.UnprotectSheet("Sheet1"))
	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestUnprotectSheet.xlsx")))
}

func TestSetDefaultTimeStyle(t *testing.T) {
	f := NewFile()
	// Test set default time style on not exists worksheet.
	assert.EqualError(t, f.setDefaultTimeStyle("SheetN", "", 0), "sheet SheetN is not exist")

	// Test set default time style on invalid cell
	assert.EqualError(t, f.setDefaultTimeStyle("Sheet1", "", 42), "cannot convert cell \"\" to coordinates: invalid cell name \"\"")
}

func TestAddVBAProject(t *testing.T) {
	f := NewFile()
	assert.NoError(t, f.SetSheetPrOptions("Sheet1", CodeName("Sheet1")))
	assert.EqualError(t, f.AddVBAProject("macros.bin"), "stat macros.bin: no such file or directory")
	assert.EqualError(t, f.AddVBAProject(filepath.Join("test", "Book1.xlsx")), ErrAddVBAProject.Error())
	assert.NoError(t, f.AddVBAProject(filepath.Join("test", "vbaProject.bin")))
	// Test add VBA project twice.
	assert.NoError(t, f.AddVBAProject(filepath.Join("test", "vbaProject.bin")))
	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestAddVBAProject.xlsm")))
}

func TestContentTypesReader(t *testing.T) {
	// Test unsupported charset.
	f := NewFile()
	f.ContentTypes = nil
	f.Pkg.Store("[Content_Types].xml", MacintoshCyrillicCharset)
	f.contentTypesReader()
}

func TestWorkbookReader(t *testing.T) {
	// Test unsupported charset.
	f := NewFile()
	f.WorkBook = nil
	f.Pkg.Store("xl/workbook.xml", MacintoshCyrillicCharset)
	f.workbookReader()
}

func TestWorkSheetReader(t *testing.T) {
	// Test unsupported charset.
	f := NewFile()
	f.Sheet.Delete("xl/worksheets/sheet1.xml")
	f.Pkg.Store("xl/worksheets/sheet1.xml", MacintoshCyrillicCharset)
	_, err := f.workSheetReader("Sheet1")
	assert.EqualError(t, err, "xml decode error: XML syntax error on line 1: invalid UTF-8")

	// Test on no checked worksheet.
	f = NewFile()
	f.Sheet.Delete("xl/worksheets/sheet1.xml")
	f.Pkg.Store("xl/worksheets/sheet1.xml", []byte(`<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>`))
	f.checked = nil
	_, err = f.workSheetReader("Sheet1")
	assert.NoError(t, err)
}

func TestRelsReader(t *testing.T) {
	// Test unsupported charset.
	f := NewFile()
	rels := "xl/_rels/workbook.xml.rels"
	f.Relationships.Store(rels, nil)
	f.Pkg.Store(rels, MacintoshCyrillicCharset)
	f.relsReader(rels)
}

func TestDeleteSheetFromWorkbookRels(t *testing.T) {
	f := NewFile()
	rels := "xl/_rels/workbook.xml.rels"
	f.Relationships.Store(rels, nil)
	assert.Equal(t, f.deleteSheetFromWorkbookRels("rID"), "")
}

func TestAttrValToInt(t *testing.T) {
	_, err := attrValToInt("r", []xml.Attr{
		{Name: xml.Name{Local: "r"}, Value: "s"}})
	assert.EqualError(t, err, `strconv.Atoi: parsing "s": invalid syntax`)
}

func prepareTestBook1() (*File, error) {
	f, err := OpenFile(filepath.Join("test", "Book1.xlsx"))
	if err != nil {
		return nil, err
	}

	err = f.AddPicture("Sheet2", "I9", filepath.Join("test", "images", "excel.jpg"),
		`{"x_offset": 140, "y_offset": 120, "hyperlink": "#Sheet2!D8", "hyperlink_type": "Location"}`)
	if err != nil {
		return nil, err
	}

	// Test add picture to worksheet with offset, external hyperlink and positioning.
	err = f.AddPicture("Sheet1", "F21", filepath.Join("test", "images", "excel.png"),
		`{"x_offset": 10, "y_offset": 10, "hyperlink": "https://github.com/360EntSecGroup-Skylar/excelize", "hyperlink_type": "External", "positioning": "oneCell"}`)
	if err != nil {
		return nil, err
	}

	file, err := ioutil.ReadFile(filepath.Join("test", "images", "excel.jpg"))
	if err != nil {
		return nil, err
	}

	err = f.AddPictureFromBytes("Sheet1", "Q1", "", "Excel Logo", ".jpg", file)
	if err != nil {
		return nil, err
	}

	return f, nil
}

func prepareTestBook3() (*File, error) {
	f := NewFile()
	f.NewSheet("Sheet1")
	f.NewSheet("XLSXSheet2")
	f.NewSheet("XLSXSheet3")
	if err := f.SetCellInt("XLSXSheet2", "A23", 56); err != nil {
		return nil, err
	}
	if err := f.SetCellStr("Sheet1", "B20", "42"); err != nil {
		return nil, err
	}
	f.SetActiveSheet(0)

	err := f.AddPicture("Sheet1", "H2", filepath.Join("test", "images", "excel.gif"),
		`{"x_scale": 0.5, "y_scale": 0.5, "positioning": "absolute"}`)
	if err != nil {
		return nil, err
	}

	err = f.AddPicture("Sheet1", "C2", filepath.Join("test", "images", "excel.png"), "")
	if err != nil {
		return nil, err
	}

	return f, nil
}

func prepareTestBook4() (*File, error) {
	f := NewFile()
	if err := f.SetColWidth("Sheet1", "B", "A", 12); err != nil {
		return f, err
	}
	if err := f.SetColWidth("Sheet1", "A", "B", 12); err != nil {
		return f, err
	}
	if _, err := f.GetColWidth("Sheet1", "A"); err != nil {
		return f, err
	}
	if _, err := f.GetColWidth("Sheet1", "C"); err != nil {
		return f, err
	}

	return f, nil
}

func fillCells(f *File, sheet string, colCount, rowCount int) {
	for col := 1; col <= colCount; col++ {
		for row := 1; row <= rowCount; row++ {
			cell, _ := CoordinatesToCellName(col, row)
			if err := f.SetCellStr(sheet, cell, cell); err != nil {
				fmt.Println(err)
			}
		}
	}
}

func BenchmarkOpenFile(b *testing.B) {
	for i := 0; i < b.N; i++ {
		if _, err := OpenFile(filepath.Join("test", "Book1.xlsx")); err != nil {
			b.Error(err)
		}
	}
}
