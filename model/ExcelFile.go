package model

type ExcelFile struct{
	fileName string
	filePath string
	sheets []IExcelSheet
}

func (file *ExcelFile) SetFilePath(filePath string) {
	file.filePath = filePath
}

func (file *ExcelFile) SetFileName(fileName string) {
	file.fileName = fileName
}

func (file *ExcelFile) AddExcelSheet(excelSheet IExcelSheet) {
	if file.sheets == nil {
		file.sheets = make([]IExcelSheet, 1)
	}
	file.sheets = append(file.sheets, excelSheet)
}

func (file *ExcelFile) GetFilePath() string {
	return file.filePath
}

func (file *ExcelFile) GetFileName() string {
	return file.fileName
}

func (file *ExcelFile) GetExcelSheets() []IExcelSheet {
	return file.sheets
}
