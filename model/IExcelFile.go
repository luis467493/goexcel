package model

type IExcelFile interface {
	SetFilePath(filePath string)
	SetFileName(fileName string)
	AddExcelSheet(excelSheet IExcelSheet)
	GetFilePath() string
	GetFileName() string
	GetExcelSheets() []IExcelSheet
}
