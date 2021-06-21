package model

type ExcelSheet struct{
	active bool
	sheetName string
	cells []IExcelCell
}

func (sheet *ExcelSheet) SetActive(active bool) {
	sheet.active = active
}

func (sheet *ExcelSheet) SetSheetName(sheetName string) {
	sheet.sheetName = sheetName
}

func (sheet *ExcelSheet) AddSheetCell( cell IExcelCell) {
	if sheet.cells == nil {
		sheet.cells = make([]IExcelCell, 1)
	}
	sheet.cells = append(sheet.cells, cell)
}

func (sheet *ExcelSheet) GetSheetName() string {
	return sheet.sheetName
}

func (sheet *ExcelSheet) GetSheetCells() []IExcelCell {
	return sheet.cells
}

func (sheet *ExcelSheet) IsActive() bool {
	return sheet.active
}
