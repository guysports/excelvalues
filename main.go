package main

import (
	"flag"
	"fmt"

	"github.com/xuri/excelize/v2"
)

func main() {
	var filePath, savePath string
	hiddenCols := map[string]bool{}
	hiddenRows := map[int]bool{}

	flag.StringVar(&filePath, "in", "", "path to the file to convert to values")
	flag.StringVar(&savePath, "out", "", "path to save the values only file to")
	flag.Parse()
	// Load the XLSX file
	xlsx, err := excelize.OpenFile(filePath)
	if err != nil {
		fmt.Println("OpenFile: " + err.Error())
		return
	}
	defer func() {
		if err := xlsx.Close(); err != nil {
			fmt.Println("Close: " + err.Error())
		}
	}()

	// Get the names of all the sheets in the XLSX file
	sheets := xlsx.GetSheetMap()

	// Loop through each sheet and convert formulas to values while keeping cell formats and colors
	for idx, sheetName := range sheets {
		// Set the active sheet
		xlsx.SetActiveSheet(idx)

		// Loop through each row and column and convert formulas to values
		rows, err := xlsx.Rows(sheetName)
		if err != nil {
			fmt.Println("Rows: " + err.Error())
		}
		currRow := 1
		furthestCol := 1
		for rows.Next() {
			if rVis, _ := xlsx.GetRowVisible(sheetName, currRow); !rVis {
				hiddenRows[currRow] = true
			}
			row, err := rows.Columns()
			if err != nil {
				fmt.Println("Columns: " + err.Error())
			}
			if len(row)+1 > furthestCol {
				furthestCol = len(row) + 1
			}
			for col := range row {
				column, _ := excelize.ColumnNumberToName(col)
				if cVis, _ := xlsx.GetColVisible(sheetName, column); !cVis {
					hiddenCols[column] = true
				}

				cellCoords, err := excelize.CoordinatesToCellName(col+1, currRow, false)
				if err != nil {
					fmt.Println("CoordinatesToCellName: " + err.Error())
				}

				cellValue, _ := xlsx.GetCellValue(sheetName, cellCoords, excelize.Options{
					MaxCalcIterations: 10,
				})
				hcol, _ := hiddenCols[column]
				hrow, _ := hiddenRows[currRow]
				if !hcol && !hrow {
					xlsx.SetCellValue(sheetName, cellCoords, cellValue)
				}
			}
			currRow++
		}
		rows.Close()

		// Unhide hidden rows and columns
		for colName := range hiddenCols {
			err = xlsx.SetColVisible(sheetName, colName, true)
			if err != nil {
				fmt.Println("SetColVisble: " + err.Error())
			}
		}
		for rowNum := range hiddenRows {
			err = xlsx.SetRowVisible(sheetName, rowNum, true)
			if err != nil {
				fmt.Println("SetRowVisible: " + err.Error())
			}
		}

		// Blank hidden columns
		for colName := range hiddenCols {
			// if err = xlsx.RemoveCol(sheetName, colName); err != nil {
			// 	fmt.Println("RemoveCol: " + err.Error())
			// }
			for idx := 1; idx <= currRow; idx++ {
				// Blank the cell entry
				if err = xlsx.SetCellValue(sheetName, fmt.Sprintf("%s%d", colName, idx), ""); err != nil {
					fmt.Println("Blanking column: " + err.Error())
				}
			}
		}

		for rowNum := range hiddenRows {
			for col := 1; col <= furthestCol; col++ {
				cellCoords, err := excelize.CoordinatesToCellName(col, rowNum, false)
				if err != nil {
					fmt.Println("Blanking Row Coords: " + err.Error())
				}
				// Blank the cell entry
				if err = xlsx.SetCellValue(sheetName, cellCoords, ""); err != nil {
					fmt.Println("Blanking row: " + err.Error())
				}
			}
		}

		// re-hide rows and columns
		for colName := range hiddenCols {
			err = xlsx.SetColVisible(sheetName, colName, false)
			if err != nil {
				fmt.Println("SetColVisble: " + err.Error())
			}
		}
		for rowNum := range hiddenRows {
			err = xlsx.SetRowVisible(sheetName, rowNum, false)
			if err != nil {
				fmt.Println("SetRowVisible: " + err.Error())
			}
		}

		// Save the updated XLSX file with the same name and format as the original
		err = xlsx.SaveAs(savePath)
		if err != nil {
			fmt.Println(err)
			return
		}
	}
}
