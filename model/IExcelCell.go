package model

type IExcelCell interface {
	SetRow(row string)
	SetColumn(column string)
	SetCellValue(cellValue string)

	GetRow() string
	GetColumn() string
	GetCellValue() string
	GetAxis() string
}

