package main

import (
	"context"
	"errors"
	"fmt"
	"os"
	"path/filepath"
	"strings"

	"github.com/wailsapp/wails/v2/pkg/runtime"

	Excel "wails-excel-tools/excel"
)

type App struct {
	ctx context.Context
}

func NewApp() *App {
	return &App{}
}

// startup is called at application startup
func (a *App) startup(ctx context.Context) {
	// Perform your setup here
	a.ctx = ctx
}

// domReady is called after front-end resources have been loaded
func (a App) domReady(ctx context.Context) {
	// Add your action here
}

// beforeClose is called when the application is about to quit,
// either by clicking the window close button or calling runtime.Quit.
// Returning true will cause the application to continue, false will continue shutdown as normal.
func (a *App) beforeClose(ctx context.Context) (prevent bool) {
	return false
}

// shutdown is called at application termination
func (a *App) shutdown(ctx context.Context) {
	// Perform your teardown here
}

// Greet returns a greeting for the given name
func (a *App) Greet(name string) string {
	return fmt.Sprintf("Hello %s, It's show time!", name)
}

// OpenFile - Opens a file dialog and returns the selected file path
func (a *App) OpenExcelFile() (string, error) {
	filePath, err := runtime.OpenFileDialog(a.ctx, runtime.OpenDialogOptions{
		Title: "Select an Excel File",
		Filters: []runtime.FileFilter{
			{DisplayName: "Excel Files", Pattern: "*.xlsx;*.xls"},
		},
	})
	if err != nil {
		return "", err
	}
	if filePath == "" {
		return "", errors.New("no file selected")
	}
	return filePath, nil
}


type SheetResponse struct {
	File1Sheets []string `json:"file1Sheets"`
	File2Sheets []string `json:"file2Sheets"`
}


func (a *App) GetExcelFilesSheetNames(file1 string, file2 string) (SheetResponse, error) {
	// Check if files exist
	if _, err := os.Stat(file1); os.IsNotExist(err) {
		return SheetResponse{}, fmt.Errorf("file1 does not exist: %s", file1)
	}
	if _, err := os.Stat(file2); os.IsNotExist(err) {
		return SheetResponse{}, fmt.Errorf("file2 does not exist: %s", file2)
	}

	// Process the first file
	sheets1, err := a.GetFileSheetsNames(file1)
	if err != nil {
		return SheetResponse{}, fmt.Errorf("failed to process file1: %s", err)
	}

	// Process the second file
	sheets2, err := a.GetFileSheetsNames(file2)
	if err != nil {
		return SheetResponse{}, fmt.Errorf("failed to process file2: %s", err)
	}

	return SheetResponse{
		File1Sheets: sheets1,
		File2Sheets: sheets2,
	}, nil
}

func (a *App) GetFileSheetsNames(filePath string) ([]string, error) {
	ext := filepath.Ext(filePath)
	switch ext {
	case ".xlsx":
		return Excel.GetXlsxSheetNames(filePath)
	case ".xls":
		// Show warning dialog for .xls files
		runtime.MessageDialog(a.ctx, runtime.MessageDialogOptions{
			Type:    runtime.WarningDialog,
			Title:   "Формат не поддерживается",
			Message: "Формат .xls не поддерживается. Используйте .xlsx файлы.",
		})
		return nil, fmt.Errorf("unsupported file format: .xls")
	default:
		return nil, fmt.Errorf("unsupported file type: %s", ext)
	}
}

type PersonsData struct {
    Arr1 []string
    Arr2 []string
}

func (a *App) CombineExcelFiles(file1Path, file2Path, sheetName1, sheetName2 string) error {
	// Read File 1
	file1Data, err := Excel.ReadXLSX(file1Path, sheetName1)
	if err != nil {
		return fmt.Errorf("failed to read file1: %v", err)
	}

	// Read File 2
	file2Data, err := Excel.ReadXLSX(file2Path, sheetName2)
	if err != nil {
		return fmt.Errorf("failed to read file2: %v", err)
	}

	// Use the first row (headers) from File 1
	headers := file1Data[0]
	file1Data = file1Data[1:] // Skip headers
	file2Data = file2Data[1:] // Skip headers

	combinedData := [][]string{headers} // Start with headers
	nameIndex := make(map[string]int)   // Map to store the row index of names

	personsData := make(map[string]PersonsData) 

	for _, row := range file1Data {
		if len(row) > 0 {
			name := strings.ToLower(row[0]) // Column A is the name
			combinedData = append(combinedData, row)
			nameIndex[name] = len(combinedData) - 1 // Map name to its row index

			personsData[name] = PersonsData{
				Arr2: row,
			}
		}
	}

	for _, row := range file2Data {
		if len(row) > 0 {
			name := strings.ToLower(row[0]) // Column A is the name
			if index, exists := nameIndex[name]; exists {
				combinedData = append(combinedData[:index+1], append([][]string{row}, combinedData[index+1:]...)...)
				
			}
			personsData[name] = PersonsData{
				Arr2: personsData[name].Arr2,
				Arr1: row,
			}
		}
	}

	newCombinedData := [][]string{headers}
	var lastName string

	for name, data := range personsData {
		if lastName != "" && lastName != name {
			newCombinedData = append(newCombinedData, []string{})
		}
		if len(data.Arr1) > 0 {
			data.Arr1[0] = "1C: " + data.Arr1[0] // Modify the first value
			newCombinedData = append(newCombinedData, data.Arr1)
		}
		if len(data.Arr2) > 0 {
			// Add Arr2 row (the second row for the person)
			newCombinedData = append(newCombinedData, data.Arr2)
		}

		lastName = name
	}

	// Print newCombinedData for debugging
	for _, row := range newCombinedData {
		fmt.Println(row)
	}

	// for _, row := range file2Data {
	// 	if len(row) > 0 {
	// 		name := strings.ToLower(row[0]) // Column A is the name
	// 		if index, exists := nameIndex[name]; exists {
	// 			// Add data from File 2 as a second row
	// 			fmt.Println("Exist, ", name)
	// 			combinedData = append(combinedData[:index+1], append([][]string{row}, combinedData[index+1:]...)...)
	// 		} else {
	// 			fmt.Println("Doesn't exist, ", name)
	// 		}
	// 	}
	// }

	// Prompt the user to save the file
	savePath, err := runtime.SaveFileDialog(a.ctx, runtime.SaveDialogOptions{
		Title:           "Save Combined File",
		DefaultFilename: "combined.xlsx",
		Filters: []runtime.FileFilter{
			{DisplayName: "Excel Files", Pattern: "*.xlsx"},
		},
	})
	if err != nil || savePath == "" {
		return fmt.Errorf("file save dialog was canceled or failed: %v", err)
	}

	// Write the combined data to the new file
	if err := Excel.WriteXLSXFormatted(savePath, newCombinedData); err != nil {
		return fmt.Errorf("failed to write combined file: %v", err)
	}

	fmt.Printf("Combined file saved as %s\n", savePath)
	return nil
}