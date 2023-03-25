package exceltool

import (
	"fmt"
	"strconv"

	"github.com/dsnet/try"
	"github.com/xuri/excelize/v2"
)

type ExcelTool struct {
	FileName string
	Excel    *excelize.File
}

func NewExcel(filename string) *ExcelTool {
	return &ExcelTool{
		FileName: filename,
		Excel:    excelize.NewFile(),
	}
}

func (tool *ExcelTool) AddSheet(name string) {
	try.E1(tool.Excel.NewSheet(name))
	try.E(tool.Excel.DeleteSheet("Sheet1")) //remove default after add
}

func (tool *ExcelTool) AddHeader(sheet string, header []string) {
	row := []any{}
	for _, v := range header {
		row = append(row, v)
	}
	try.E(tool.Excel.SetSheetRow(sheet, "A1", &row))
}

func (tool *ExcelTool) AddRow(sheet string, row int, data []any) {
	try.E(tool.Excel.SetSheetRow(sheet, "A"+strconv.Itoa(row), &data))
}

func (tool *ExcelTool) SetStyle(sheet string, cellRange string, style string) {
	//cell range A1:E100
	if style == "" {
		style = "TableStyleMedium14"
	}
	showRowStripes := true

	try.E(tool.Excel.AddTable(sheet, cellRange, &excelize.TableOptions{
		Name:              "table",
		StyleName:         style,
		ShowFirstColumn:   false,
		ShowLastColumn:    false,
		ShowRowStripes:    &showRowStripes,
		ShowColumnStripes: false,
	}))
}

func (tool *ExcelTool) Save() {
	try.E(tool.Excel.SaveAs(tool.FileName))
}

func (tool *ExcelTool) Close() {
	try.E(tool.Excel.Close())
}

func (tool *ExcelTool) NextColumn(header []string) string {
	return fmt.Sprintf("%c", 'A'+len(header))
}
