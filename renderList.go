package xlsx_template

import (
	"github.com/tealeg/xlsx"
	"regexp"
	"strings"
)

var (
	listRgx = regexp.MustCompile(`{{\s*.(\w+)\.\w+\s*}}`)
)

func findListProp(in *xlsx.Row, ctx interface{}) string {
	for _, cell := range in.Cells {
		if cell.Value == "" {
			continue
		}
		if match := listRgx.FindAllStringSubmatch(cell.Value, -1); match != nil {
			for i := 0; i < len(match); i++ {
				for j := 0; j < len(match[i]); j++ {
					if flg, _ := isSliceOrArray(ctx, match[i][j]); flg == true {
						return match[i][j]
					}
				}
			}
		}
	}
	return ""
}

func prepareListProp(in *xlsx.Row, Prop string) {

	for _, cell := range in.Cells {
		cell.Value = strings.Replace(cell.Value, "."+Prop+".", ".", strings.Count(cell.Value, "."+Prop+"."))
	}
}

// rendering list property slice or array {{.xxx.yyy}}
func renderList(sheet *xlsx.Sheet, row *xlsx.Row, ctx interface{}) (IsRender bool, err error) {

	prop := findListProp(row, ctx)
	if prop == "" {
		return false, nil
	}

	arr, err := getField(ctx, prop)
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
