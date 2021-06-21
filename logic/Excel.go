package logic

import (
	"excel/constants"
	"excel/model"
	"fmt"

	"github.com/360EntSecGroup-Skylar/excelize/v2"
)

type Excel struct{

}

func (Excel) WriteExcelFile(file model.IExcelFile) {
	f := excelize.NewFile()

	sheets := file.GetExcelSheets()
	for index, sheet := range sheets {
		if sheet != nil {
			f.NewSheet(sheet.GetSheetName())
			if sheet.IsActive() {
				f.SetActiveSheet(index)
			}
			for _, cell := range sheet.GetSheetCells() {
				if cell != nil {
					f.SetCellValue(sheet.GetSheetName(), cell.GetAxis(), cell.GetCellValue())
				}
			}
		}
	}

	path := file.GetFilePath() + constants.Separator + file.GetFileName()

	if err := f.SaveAs(path); err != nil {
		fmt.Println(err)
	}
}
