package goexcel

import (
	"strings"

	"github.com/xuri/excelize/v2"
)

type ExcelFunc func(*ExcelHandle) error

type ExcelHandle struct {
	*excelize.File
}

func (f *ExcelHandle) GetRowNumber() int {
	idx := f.GetActiveSheetIndex()
	sheetName := f.GetSheetName(idx)
	rows, _ := f.GetRows(sheetName)
	return len(rows)
}

func (f *ExcelHandle) GetColNumber() int {
	idx := f.GetActiveSheetIndex()
	sheetName := f.GetSheetName(idx)
	cols, _ := f.GetCols(sheetName)
	return len(cols)
}

func (f *ExcelHandle) GetAllRows() [][]string {
	idx := f.GetActiveSheetIndex()
	sheetName := f.GetSheetName(idx)
	rows, _ := f.GetRows(sheetName)
	return rows
}

func (f *ExcelHandle) GetCell(key string) string {
	idx := f.GetActiveSheetIndex()
	sheetName := f.GetSheetName(idx)
	str, _ := f.GetCellValue(sheetName, key)
	return str
}

func (f *ExcelHandle) GetRowCell(key string) []string {
	_, y, _ := excelize.CellNameToCoordinates(key)
	return f.GetRowCellByXY(y)
}

func (f *ExcelHandle) GetRowCellByXY(y int) []string {
	idx := f.GetActiveSheetIndex()
	sheetName := f.GetSheetName(idx)
	var vals []string
	cols, _ := f.GetCols(sheetName)
	for i := 0; i < len(cols); i++ {
		key, _ := excelize.CoordinatesToCellName(i, y)
		str, _ := f.GetCellValue(sheetName, key)
		vals = append(vals, str)
	}
	return vals
}

func (f *ExcelHandle) SetCell(key string, value interface{}) {
	idx := f.GetActiveSheetIndex()
	sheetName := f.GetSheetName(idx)
	f.SetCellValue(sheetName, key, value)
}

func (f *ExcelHandle) AppendRowCell(value ...interface{}) {
	row := f.GetRowNumber()
	f.SetRowCellByXY(1, row+1, value...)
}

func (f *ExcelHandle) SetRowCell(key string, value ...interface{}) {
	x, y, _ := excelize.CellNameToCoordinates(key)
	f.SetRowCellByXY(x, y, value...)
}

func (f *ExcelHandle) SetRowCellByXY(x, y int, value ...interface{}) {
	idx := f.GetActiveSheetIndex()
	sheetName := f.GetSheetName(idx)
	for i, v := range value {
		key, _ := excelize.CoordinatesToCellName(x+i, y)
		f.SetCellValue(sheetName, key, v)
	}
}

// err = f.SaveAs(fileName)
func ExcelInit(fileName, sheet string) (file *ExcelHandle, err error) {
	const fileNotExsitSuff = "The system cannot find the file specified."
	var f *excelize.File
	f, err = excelize.OpenFile(fileName)
	if err != nil && strings.HasSuffix(err.Error(), fileNotExsitSuff) {
		f = excelize.NewFile()
	}

	idx, err := f.NewSheet(sheet)
	if err != nil {
		return
	}

	file = &ExcelHandle{f}
	f.SetActiveSheet(idx)
	return
}

func ExcelRun(fileName, sheet string, fn ExcelFunc) (err error) {
	const fileNotExsitSuff = "The system cannot find the file specified."
	var f *excelize.File
	f, err = excelize.OpenFile(fileName)
	if err != nil && strings.HasSuffix(err.Error(), fileNotExsitSuff) {
		f = excelize.NewFile()
	}

	idx, err := f.NewSheet(sheet)
	if err != nil {
		return
	}

	file := &ExcelHandle{f}
	f.SetActiveSheet(idx)

	if err = fn(file); err != nil {
		return
	}

	return f.SaveAs(fileName)
}
