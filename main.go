// package main

// import (
// 	"fmt"

// 	"github.com/xuri/excelize/v2"
// )

// func main() {
// 		f := excelize.NewFile()
// 		// Create a new worksheet.
// 		index := f.NewSheet("Sheet2")
// 		// Set value of a cell.
// 		f.SetCellValue("Sheet2", "A2", "Hello world.")
// 		f.SetCellValue("Sheet1", "B2", 100)
// 		// Set the active worksheet of the workbook.
// 		f.SetActiveSheet(index)
// 		// Save the spreadsheet by the given path.
// 		if err := f.SaveAs("excel.xlsx"); err != nil {
// 			fmt.Println(err)
// 		}
// }

package main

import (
	"fmt"

	"github.com/xuri/excelize/v2"
)

func main() {
	f, err := excelize.OpenFile("arza.xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}
	defer func() {
		// Close the spreadsheet.
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()

	// HSUK bca kody
	f.SetCellValue("Sheet4", "M15", 1)
	f.SetCellValue("Sheet4", "N15", 2)
	f.SetCellValue("Sheet4", "O15", 3)
	f.SetCellValue("Sheet4", "P15", 4)
	f.SetCellValue("Sheet4", "Q15", 5)
	f.SetCellValue("Sheet4", "R15", 6)
	f.SetCellValue("Sheet4", "S15", 7)
	f.SetCellValue("Sheet4", "T15", 8)

	// EEYBDK bca kody
	f.SetCellValue("Sheet4", "O26", 1)
	f.SetCellValue("Sheet4", "P26", 1)
	f.SetCellValue("Sheet4", "Q26", 1)
	f.SetCellValue("Sheet4", "R26", 1)
	f.SetCellValue("Sheet4", "S26", 1)
	f.SetCellValue("Sheet4", "T26", 1)

	// Yuridik salgy
	f.SetCellValue("Sheet4", "E17", "Gorogly 113, Ashgabat, Turkmenistan")

	// EEYBDK bca kody
	f.SetCellValue("Sheet4", "K23", 9)
	f.SetCellValue("Sheet4", "L23", 9)
	f.SetCellValue("Sheet4", "M23", 9)
	f.SetCellValue("Sheet4", "N23", 9)
	f.SetCellValue("Sheet4", "O23", 9)
	f.SetCellValue("Sheet4", "P23", 9)
	f.SetCellValue("Sheet4", "Q23", 9)
	f.SetCellValue("Sheet4", "R23", 9)
	f.SetCellValue("Sheet4", "S23", 9)
	f.SetCellValue("Sheet4", "T23", 9)

	// Ortaca ishgarlerin sany
	f.SetCellValue("Sheet4", "O28", 100)

	// Save the spreadsheet by the given path.
	if err := f.SaveAs("final-arza.xlsx"); err != nil {
		fmt.Println(err)
	}
}
