package logic

import "excel/model"

type IExcel interface {
	WriteExcelFile(file model.IExcelFile)
}
