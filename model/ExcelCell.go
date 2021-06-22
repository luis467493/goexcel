package model

type ExcelCell struct {
	row    string
	column string
	value  string
}

func (cell *ExcelCell) SetRow(row string) {
	cell.row = row
}

func (cell *ExcelCell) SetColumn(column string) {
	cell.column = column
}

func (cell *ExcelCell) SetCellValue(cellValue string) {
	cell.value = cellValue
}

func (cell *ExcelCell) GetRow() string {
	return cell.row
}

func (cell *ExcelCell) GetColumn() string {
	return cell.column
}

func (cell *ExcelCell) GetCellValue() string {
	return cell.value
}

func (cell *ExcelCell) GetAxis() string {
	return cell.column + cell.row
}
