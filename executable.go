package main

import (
	logic "excel/logic"
	"excel/model"
	"strconv"
)

func main() {

	rows := 5
	columns := 5
	array := make([][]string, rows)
	for i := range array {
		array[i] = make([]string, columns)
		for iColumn, _ := range array[i] {
			array[i][iColumn] = "sii"
		}
	}
	createExcelFile("C:\\GoRepo\\Labs\\3\\goexcel", "PruebaOtro.xlsx", array)
}

func createExcelFile(path string, name string, array [][]string) {
	excel := new(logic.Excel)
	excelFile := new(model.ExcelFile)
	excelFile.SetFilePath(path)
	excelFile.SetFileName(name)
	addSheetToFile(excelFile, array)
	excel.WriteExcelFile(excelFile)
}

func addSheetToFile(excelFile *model.ExcelFile, array [][]string) {
	excelSheet := new(model.ExcelSheet)
	excelSheet.SetSheetName("Sheet1")

	for iRow, columns := range array {
		for iColumn, column := range columns {
			addCellToSheet(excelSheet, strconv.Itoa(iRow+1), string(iColumn+65), column)
		}
	}

	excelFile.AddExcelSheet(excelSheet)
}

func addCellToSheet(excelSheet *model.ExcelSheet,
	row string, column string, value string) {
	excelCell := new(model.ExcelCell)
	excelCell.SetRow(row)
	excelCell.SetColumn(column)
	excelCell.SetCellValue(value)
	excelSheet.AddSheetCell(excelCell)
}
