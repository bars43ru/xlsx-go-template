package xlsx_template

import (
	"errors"
	"fmt"
	"regexp"

	"github.com/tealeg/xlsx/v3"
)

var (
	rangeRgx    = regexp.MustCompile(`{{\s*range\s+.(\w+)\s*}}`)
	rangeEndRgx = regexp.MustCompile(`{{\s*end\s*}}`)
)

type rangeRows struct {
	PropName string
	BRow     int
	ERow     int
}

func getRangeRows(beginRow, endRow int, template *xlsx.Sheet) (*rangeRows, error) {
	// check begin range
	row, err := template.Row(beginRow)
	if err != nil {
		return nil, fmt.Errorf("get row: %w", err)
	}
	retValue := &rangeRows{
		PropName: getRangeRgx(row),
		BRow:     beginRow + 1,
	}

	if retValue.PropName == "" {
		return nil, nil
	}

	// find end range
	if ri, err := getRangeEndRgx(retValue.BRow, endRow, template); err != nil {
		return nil, fmt.Errorf("range '%s': %w", retValue.PropName, err)
	} else {
		retValue.ERow = ri
	}

	return retValue, nil
}

func getRangeRgx(in *xlsx.Row) string {
	cell := in.GetCell(0)
	match := rangeRgx.FindAllStringSubmatch(cell.Value, -1)
	if match != nil {
		return match[0][1]
	}
	return ""
}

func getRangeEndRgx(beginRow, endRow int, template *xlsx.Sheet) (int, error) {
	var nesting int

	for idx := beginRow; idx <= endRow; idx++ {
		row, err := template.Row(idx)
		if err != nil {
			return 0, fmt.Errorf("get row: %w", err)
		}
		cell := row.GetCell(0)
		if rangeEndRgx.MatchString(cell.Value) {
			if nesting == 0 {
				return idx, nil
			}

			nesting--
			continue
		}

		if rangeRgx.MatchString(cell.Value) {
			nesting++
		}
	}
	return -1, errors.New("not found end of range")
}

// template, destination *xlsx.Sheet, startRow, endRow int,
func renderRange(iRow *int, template, sheet *xlsx.Sheet, endRow int, v any) (IsRender bool, err error) {
	val, err := getRangeRows(*iRow, endRow, template)
	if err != nil || val == nil {
		return false, err
	}

	var flg bool
	flg, err = isSliceOrArray(v, val.PropName)
	if err != nil {
		return false, err
	}
	if !flg {
		return false, fmt.Errorf("range '%s' error: field '%s' in v is not slice or array", val.PropName, val.PropName)
	}

	var items []any
	items, err = getField(v, val.PropName)
	if err != nil {
		return false, err
	}

	for _, item := range items {
		err = renderRows(template, sheet, val.BRow, val.ERow-1, item)
		if err != nil {
			return false, err
		}
	}
	*iRow = val.ERow
	return true, nil
}
