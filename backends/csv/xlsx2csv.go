package csv

import (
	"bytes"
	stdcsv "encoding/csv"
	"errors"
	"io"
	"io/ioutil"
	"path/filepath"

	"github.com/tealeg/xlsx/v2"
)

func generateCSVFromXLSXFile(fileName string) (io.ReadCloser, error) {
	xlFile, err := xlsx.OpenFile(fileName)
	if err != nil {
		return nil, err
	}
	if len(xlFile.Sheets) == 0 {
		return nil, errors.New("This XLSX file contains no sheets")
	}
	sheet := xlFile.Sheets[0]

	var buf bytes.Buffer
	csvWriter := stdcsv.NewWriter(&buf)

	var firstRowSize int

	for i := 0; i < sheet.MaxRow; i++ {
		row := sheet.Row(i)

		if row.Hidden {
			continue
		}

		var record []string

		for j := 0; j < sheet.MaxCol; j++ {
			cell := sheet.Cell(i, j)
			record = append(record, cell.Value)
		}

		if len(record) == 0 {
			continue
		}

		if firstRowSize == 0 {
			firstRowSize = len(record)
		}

		if firstRowSize != len(record) {
		}

		err = csvWriter.Write(record)
		if err != nil {
		}
	}

	if err != nil {
		return nil, err
	}

	csvWriter.Flush()
	err = csvWriter.Error()
	if err != nil {
		return nil, err
	}

	return ioutil.NopCloser(&buf), nil
}

func isXLSXFile(fileName string) bool {
	return filepath.Ext(fileName) == ".xlsx"
}
