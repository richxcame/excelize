package main

import (
	"fmt"

	"github.com/xuri/excelize/v2"
)

func main() {
	f := excelize.NewFile()
	// Create a new worksheet.
	index := f.NewSheet("Sheet2")
	// Set value of a cell.
	f.SetCellValue("Sheet2", "A2", "Hello world.")
	f.SetCellValue("Sheet1", "B2", 100)
	// Set the active worksheet of the workbook.
	f.SetActiveSheet(index)
	// Save the spreadsheet by the given path.
	if err := f.SaveAs("excel.xlsx"); err != nil {
		fmt.Println(err)
	}
}
