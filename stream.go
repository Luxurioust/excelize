// Copyright 2016 - 2019 The excelize Authors. All rights reserved. Use of
// this source code is governed by a BSD-style license that can be found in
// the LICENSE file.
//
// Package excelize providing a set of functions that allow you to write to
// and read from XLSX files. Support reads and writes XLSX file generated by
// Microsoft Excel™ 2007 and later. Support save file without losing original
// charts of XLSX. This library needs Go version 1.10 or later.

package excelize

import (
	"bytes"
	"encoding/xml"
	"errors"
	"fmt"
	"io/ioutil"
	"os"
	"reflect"
	"time"
)

// StreamWriter defined the type of stream writer.
type StreamWriter struct {
	tmpFile   *os.File
	File      *File
	Sheet     string
	SheetID   int
	SheetData bytes.Buffer
	encoder   *xml.Encoder
}

// NewStreamWriter return stream writer struct by given worksheet name for
// generate new worksheet with large amounts of data. Note that after set
// rows, you must call the 'Flush' method to end the streaming writing
// process and ensure that the order of line numbers is ascending. For
// example, set data for worksheet of size 102400 rows x 50 columns with
// numbers:
//
//    file := excelize.NewFile()
//    streamWriter, err := file.NewStreamWriter("Sheet1")
//    if err != nil {
//        panic(err)
//    }
//    for rowID := 1; rowID <= 102400; rowID++ {
//        row := make([]interface{}, 50)
//        for colID := 0; colID < 50; colID++ {
//            row[colID] = rand.Intn(640000)
//        }
//        cell, _ := excelize.CoordinatesToCellName(1, rowID)
//        if err := streamWriter.SetRow(cell, &row); err != nil {
//            panic(err)
//        }
//    }
//    if err := streamWriter.Flush(); err != nil {
//        panic(err)
//    }
//    if err := file.SaveAs("Book1.xlsx"); err != nil {
//        panic(err)
//    }
//
func (f *File) NewStreamWriter(sheet string) (*StreamWriter, error) {
	sheetID := f.GetSheetIndex(sheet)
	if sheetID == 0 {
		return nil, fmt.Errorf("sheet %s is not exist", sheet)
	}
	rsw := &StreamWriter{
		File:    f,
		Sheet:   sheet,
		SheetID: sheetID,
	}
	rsw.encoder = xml.NewEncoder(&rsw.SheetData)
	rsw.SheetData.WriteString("<sheetData>")
	return rsw, nil
}

// SetRow writes an array to stream rows by giving a worksheet name, starting
// coordinate and a pointer to an array of values. If styles is non-nil, then
// the styles must be the same size as the values and will be applied to each
// corresponding cell. Note that you must call the 'Flush' method to end the
// streaming writing process.
func (sw *StreamWriter) SetRow(axis string, values []interface{}, styles []int) error {
	col, row, err := CellNameToCoordinates(axis)
	if err != nil {
		return err
	}
	if styles == nil {
		styles = make([]int, len(values))
	}
	if len(styles) != len(values) {
		return errors.New("incorrect number of styles for this row")
	}
	sw.SheetData.WriteString(fmt.Sprintf(`<row r="%d">`, row))
	for i, val := range values {
		axis, err := CoordinatesToCellName(col+i, row)
		if err != nil {
			return err
		}
		c := xlsxC{R: axis, S: styles[i]}
		switch val := val.(type) {
		case int:
			c.T, c.V = setCellInt(val)
		case int8:
			c.T, c.V = setCellInt(int(val))
		case int16:
			c.T, c.V = setCellInt(int(val))
		case int32:
			c.T, c.V = setCellInt(int(val))
		case int64:
			c.T, c.V = setCellInt(int(val))
		case uint:
			c.T, c.V = setCellInt(int(val))
		case uint8:
			c.T, c.V = setCellInt(int(val))
		case uint16:
			c.T, c.V = setCellInt(int(val))
		case uint32:
			c.T, c.V = setCellInt(int(val))
		case uint64:
			c.T, c.V = setCellInt(int(val))
		case float32:
			c.T, c.V = setCellFloat(float64(val), -1, 32)
		case float64:
			c.T, c.V = setCellFloat(val, -1, 64)
		case string:
			c.T, c.V, c.XMLSpace = setCellStr(val)
		case []byte:
			c.T, c.V, c.XMLSpace = setCellStr(string(val))
		case time.Duration:
			c.T, c.V = setCellDuration(val)
		case time.Time:
			c.T, c.V, _, err = setCellTime(val)
		case bool:
			c.T, c.V = setCellBool(val)
		case nil:
			c.T, c.V, c.XMLSpace = setCellStr("")
		default:
			c.T, c.V, c.XMLSpace = setCellStr(fmt.Sprint(val))
		}
		sw.encoder.Encode(c)
	}
	sw.SheetData.WriteString(`</row>`)
	// Try to use local storage
	chunk := 1 << 24
	if sw.SheetData.Len() >= chunk {
		if sw.tmpFile == nil {
			err := sw.createTmp()
			if err != nil {
				// can not use local storage
				return nil
			}
		}
		// use local storage
		_, err := sw.tmpFile.Write(sw.SheetData.Bytes())
		if err != nil {
			return nil
		}
		sw.SheetData.Reset()
	}
	return err
}

// Flush ending the streaming writing process.
func (sw *StreamWriter) Flush() error {
	sw.SheetData.WriteString(`</sheetData>`)

	ws, err := sw.File.workSheetReader(sw.Sheet)
	if err != nil {
		return err
	}
	sheetXML := fmt.Sprintf("xl/worksheets/sheet%d.xml", sw.SheetID)
	delete(sw.File.Sheet, sheetXML)
	delete(sw.File.checked, sheetXML)
	var sheetDataByte []byte
	if sw.tmpFile != nil {
		// close the local storage file
		if err = sw.tmpFile.Close(); err != nil {
			return err
		}

		file, err := os.Open(sw.tmpFile.Name())
		if err != nil {
			return err
		}

		sheetDataByte, err = ioutil.ReadAll(file)
		if err != nil {
			return err
		}

		if err := file.Close(); err != nil {
			return err
		}

		err = os.Remove(sw.tmpFile.Name())
		if err != nil {
			return err
		}
	}

	sheetDataByte = append(sheetDataByte, sw.SheetData.Bytes()...)
	replaceMap := map[string][]byte{
		"XMLName":   []byte{},
		"SheetData": sheetDataByte,
	}
	sw.SheetData.Reset()
	sw.File.XLSX[fmt.Sprintf("xl/worksheets/sheet%d.xml", sw.SheetID)] =
		StreamMarshalSheet(ws, replaceMap)
	return err
}

// createTmp creates a temporary file in the operating system default
// temporary directory.
func (sw *StreamWriter) createTmp() (err error) {
	sw.tmpFile, err = ioutil.TempFile(os.TempDir(), "excelize-")
	return err
}

// StreamMarshalSheet provides method to serialization worksheets by field as
// streaming.
func StreamMarshalSheet(ws *xlsxWorksheet, replaceMap map[string][]byte) []byte {
	s := reflect.ValueOf(ws).Elem()
	typeOfT := s.Type()
	var marshalResult []byte
	marshalResult = append(marshalResult, []byte(XMLHeader+`<worksheet`+templateNamespaceIDMap)...)
	for i := 0; i < s.NumField(); i++ {
		f := s.Field(i)
		content, ok := replaceMap[typeOfT.Field(i).Name]
		if ok {
			marshalResult = append(marshalResult, content...)
			continue
		}
		out, _ := xml.Marshal(f.Interface())
		marshalResult = append(marshalResult, out...)
	}
	marshalResult = append(marshalResult, []byte(`</worksheet>`)...)
	return marshalResult
}

// setCellStr provides a function to set string type value of a cell as
// streaming. Total number of characters that a cell can contain 32767
// characters.
func (sw *StreamWriter) setCellStr(axis, value string) string {
	if len(value) > 32767 {
		value = value[0:32767]
	}
	// Leading and ending space(s) character detection.
	if len(value) > 0 && (value[0] == 32 || value[len(value)-1] == 32) {
		return fmt.Sprintf(`<c xml:space="preserve" r="%s" t="str"><v>%s</v></c>`, axis, value)
	}
	return fmt.Sprintf(`<c r="%s" t="str"><v>%s</v></c>`, axis, value)
}
