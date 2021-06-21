package main

import (
	logic "excel/logic"
	"excel/model"
)

func main() {
	createExcelFile("C:\\GoRepo\\Labs\\3\\Excel",
		"Prueba9.xlsx")
}

func createExcelFile(path string, name string) {
	excel := new(logic.Excel)
	excelFile := new(model.ExcelFile)
	excelFile.SetFilePath(path)
	excelFile.SetFileName(name)
	addSheetToFile(excelFile)
	excel.WriteExcelFile(excelFile)
}

func addSheetToFile(excelFile *model.ExcelFile) {
	excelSheet := new(model.ExcelSheet)
	excelSheet.SetSheetName("Sheet1")

	addCellToSheet(excelSheet, "A", "1", "Hablame")
	addCellToSheet(excelSheet, "B", "1", "xyz")

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
