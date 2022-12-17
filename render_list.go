package xlsx_template

import (
	"regexp"
	"strings"

	"github.com/tealeg/xlsx/v2"
)

var (
	listRgx = regexp.MustCompile(`{{\s*.(\w+)\.\w+\s*}}`)
)

func findListProp(in *xlsx.Row, v any) string {
	for _, cell := range in.Cells {
		if cell.Value == "" {
			continue
		}
		if match := listRgx.FindAllStringSubmatch(cell.Value, -1); match != nil {
			for i := 0; i < len(match); i++ {
				for j := 0; j < len(match[i]); j++ {
					if flg, _ := isSliceOrArray(v, match[i][j]); flg {
						return match[i][j]
					}
				}
			}
		}
	}
	return ""
}

func prepareListProp(in *xlsx.Row, prop string) {
	for _, cell := range in.Cells {
		cell.Value = strings.Replace(cell.Value, "."+prop+".", ".", strings.Count(cell.Value, "."+prop+"."))
	}
}

// rendering list property slice or array {{.xxx.yyy}}
func renderList(sheet *xlsx.Sheet, row *xlsx.Row, v any) (IsRender bool, err error) {
	prop := findListProp(row, v)
	if prop == "" {
		return false, nil
	}

	arr, err := getField(v, prop)
	if err != nil {
		return false, err
	}
	for i := 0; i < len(arr); i++ {
		newRow := sheet.AddRow()
		cloneRow(row, newRow)
		prepareListProp(newRow, prop)
		if err := renderRow(newRow, arr[i]); err != nil {
			return true, err
		}
	}

	return true, nil
}
