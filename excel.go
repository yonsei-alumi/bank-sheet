package main

import (
	"fmt"
	"os"
	"path/filepath"
	"strings"

	"github.com/xuri/excelize/v2"
)

// WriteExcel writes transactions to an Excel file.
func WriteExcel(transactions []Transaction, outputPath string) error {
	f := excelize.NewFile()
	defer f.Close()

	sheet := "Sheet1"

	headers := []string{"입금일자", "입금자", "입금액", "잔액"}
	for i, h := range headers {
		cell, _ := excelize.CoordinatesToCellName(i+1, 1)
		f.SetCellValue(sheet, cell, h)
	}

	headerStyle, _ := f.NewStyle(&excelize.Style{
		Font:      &excelize.Font{Bold: true},
		Alignment: &excelize.Alignment{Horizontal: "center"},
	})
	f.SetCellStyle(sheet, "A1", "D1", headerStyle)

	numberStyle, _ := f.NewStyle(&excelize.Style{
		NumFmt: 3, // #,##0
	})

	for i, txn := range transactions {
		row := i + 2
		f.SetCellValue(sheet, fmt.Sprintf("A%d", row), txn.DateTime.Format("2006-01-02 15:04:05"))
		f.SetCellValue(sheet, fmt.Sprintf("B%d", row), txn.Depositor)
		f.SetCellValue(sheet, fmt.Sprintf("C%d", row), txn.Amount)
		f.SetCellValue(sheet, fmt.Sprintf("D%d", row), txn.Balance)
		f.SetCellStyle(sheet, fmt.Sprintf("C%d", row), fmt.Sprintf("D%d", row), numberStyle)
	}

	f.SetColWidth(sheet, "A", "A", 22)
	f.SetColWidth(sheet, "B", "B", 18)
	f.SetColWidth(sheet, "C", "C", 15)
	f.SetColWidth(sheet, "D", "D", 15)

	return f.SaveAs(outputPath)
}

// uniqueFilename returns a unique filename by appending (1), (2), etc.
// if the file already exists, like browser download behavior.
func uniqueFilename(basePath string) string {
	if _, err := os.Stat(basePath); os.IsNotExist(err) {
		return basePath
	}

	ext := filepath.Ext(basePath)
	name := strings.TrimSuffix(basePath, ext)

	for i := 1; ; i++ {
		candidate := fmt.Sprintf("%s(%d)%s", name, i, ext)
		if _, err := os.Stat(candidate); os.IsNotExist(err) {
			return candidate
		}
	}
}
