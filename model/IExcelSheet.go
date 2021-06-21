package model

type IExcelSheet interface {
	SetActive(active bool)
	SetSheetName(sheetName string)
	AddSheetCell( cell IExcelCell)
	GetSheetName() string
	GetSheetCells() []IExcelCell
	IsActive() bool
}

