package excel

import (
	"fmt"
	"strconv"

	"github.com/xuri/excelize/v2"
)


func ReadXLSX(filePath, sheetName string) ([][]string, error) {
	// Open the .xlsx file
	f, err := excelize.OpenFile(filePath)
	if err != nil {
		return nil, fmt.Errorf("failed to open .xlsx file: %v", err)
	}

	// Get all rows in the specified sheet
	rows, err := f.GetRows(sheetName)
	if err != nil {
		return nil, fmt.Errorf("failed to get rows from sheet '%s': %v", sheetName, err)
	}

	return rows, nil
}

func WriteXLSXFormatted(outputPath string, data [][]string) error {
	// Create a new Excel file
	f := excelize.NewFile()

	// Create a new sheet
	sheetName := "Combined"
	index, err := f.NewSheet(sheetName)
	if err != nil {
		return fmt.Errorf("failed to create sheet: %v", err)
	}
	f.SetActiveSheet(index)

	// Define header style
	headerStyle, err := f.NewStyle(&excelize.Style{
		Font:      &excelize.Font{Bold: true, Color: "#FFFFFF"},
		Fill:      excelize.Fill{Type: "pattern", Color: []string{"#4CAF50"}, Pattern: 1},
		Alignment: &excelize.Alignment{Horizontal: "center", Vertical: "center"},
	})
	if err != nil {
		return fmt.Errorf("failed to create header style: %v", err)
	}

	// Write the data to the sheet
	for rowIndex, row := range data {
		for colIndex, cellValue := range row {
			cell, err := excelize.CoordinatesToCellName(colIndex+1, rowIndex+1)
			if err != nil {
				return fmt.Errorf("failed to convert cell coordinates: %v", err)
			}

			// Write value
			if err := f.SetCellValue(sheetName, cell, cellValue); err != nil {
				return fmt.Errorf("failed to set cell value: %v", err)
			}

			// Apply header style for the first row
			if rowIndex == 0 {
				if err := f.SetCellStyle(sheetName, cell, cell, headerStyle); err != nil {
					return fmt.Errorf("failed to set header style: %v", err)
				}
			} else {
				// Set cell type as number if the value is numeric
				if _, err := strconv.ParseFloat(cellValue, 64); err == nil {
					f.SetCellValue(sheetName, cell, cellValue) // Overwrite as number
				}
			}
		}
	}

	// Set auto-filter for the header row
	lastColumn, err := excelize.CoordinatesToCellName(len(data[0]), 1) // Get the last column name
	if err != nil {
		return fmt.Errorf("failed to calculate last column: %v", err)
	}
	filterRange := fmt.Sprintf("A1:%s", lastColumn)
	if err := f.AutoFilter(sheetName, filterRange, nil); err != nil {
		return fmt.Errorf("failed to set auto filter: %v", err)
	}

	// Adjust column widths to fit content
	for i := 0; i < len(data[0]); i++ {
		colName, err := excelize.ColumnNumberToName(i + 1)
		if err != nil {
			return fmt.Errorf("failed to convert column index to name: %v", err)
		}
		if err := f.SetColWidth(sheetName, colName, colName, 15); err != nil {
			return fmt.Errorf("failed to auto-adjust column widths: %v", err)
		}
	}

	// Save the file
	if err := f.SaveAs(outputPath); err != nil {
		return fmt.Errorf("failed to save file: %v", err)
	}

	return nil
}


func GetXlsxSheetNames(filePath string) ([]string, error) {
	f, err := excelize.OpenFile(filePath)
	if err != nil {
		return nil, fmt.Errorf("failed to open .xlsx file: %s", err)
	}
	defer f.Close()

	return f.GetSheetList(), nil
}

