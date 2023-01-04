package xlsx_template

import (
	"bytes"
	"errors"
	"fmt"
	"math"
	"strconv"
	"strings"
	"text/template"

	"github.com/tealeg/xlsx/v3"
)

func cloneSheet(from, to *xlsx.Sheet) {
	to.SheetFormat.DefaultColWidth = from.SheetFormat.DefaultColWidth
	to.SheetFormat.DefaultRowHeight = from.SheetFormat.DefaultRowHeight

	from.Cols.ForEach(func(idx int, col *xlsx.Col) {
		newCol := xlsx.Col{}
		style := col.GetStyle()
		newCol.SetStyle(style)
		newCol.Width = col.Width
		newCol.Hidden = col.Hidden
		newCol.Collapsed = col.Collapsed
		newCol.Min = col.Min
		newCol.Max = col.Max
		to.Cols.Add(&newCol)
	})
}

func cloneRow(from, to *xlsx.Row) {
	height := from.GetHeight()
	if height != 0 {
		to.SetHeight(height)
	}
	to.Hidden = from.Hidden
	_ = from.ForEachCell(func(cell *xlsx.Cell) error {
		newCell := to.AddCell()
		cloneCell(cell, newCell)
		return nil
	})
}

func cloneCell(from, to *xlsx.Cell) {
	to.Value = from.Value
	style := from.GetStyle()
	to.SetStyle(style)

	to.GetStyle().ApplyFill = style.Fill.BgColor != "" || style.Fill.FgColor != ""

	to.HMerge = from.HMerge
	to.VMerge = from.VMerge
	to.Hidden = from.Hidden
	to.NumFmt = from.NumFmt
}

func renderRow(in *xlsx.Row, v any) error {
	var (
		maxEntBefore float64
		maxEntAfter  float64
	)

	err := in.ForEachCell(func(cell *xlsx.Cell) error {
		countEnt := float64(strings.Count(cell.Value, "\n"))
		maxEntBefore = math.Max(maxEntBefore, countEnt)

		err := renderCell(cell, v)
		if err != nil {
			return err
		}

		countEnt = float64(strings.Count(cell.Value, "\n"))
		maxEntAfter = math.Max(maxEntAfter, countEnt)
		return nil
	})
	if err != nil {
		return err
	}

	maxEntAfter = (maxEntAfter + 1) / (maxEntBefore + 1)
	if maxEntAfter != 0 {
		height := in.GetHeight()
		if height != 0 {
			in.SetHeight(height * maxEntAfter)
		} else {
			in.SetHeight(in.Sheet.SheetFormat.DefaultRowHeight * maxEntAfter)
		}
	}
	return nil
}

func renderCell(cell *xlsx.Cell, v any) error {
	var buf bytes.Buffer
	tpl, err := template.New("").Parse(cell.Value)
	if err != nil {
		return err
	}
	buf.Reset()
	err = tpl.Execute(&buf, v)
	if err != nil {
		cell.Value = err.Error()
		return nil
	}

	if cell.NumFmt == "@" {
		cell.SetString(buf.String())
	} else if outFloat, err := strconv.ParseFloat(buf.String(), 64); err == nil {
		cFmt := cell.NumFmt
		cell.SetFloat(outFloat)
		cell.NumFmt = cFmt
	} else {
		cell.SetValue(buf.String())
	}

	return nil
}

func renderRows(template, destination *xlsx.Sheet, startRow, endRow int, v any) error {
	if isArray(v) {
		return errors.New("v can not be slice or array")
	}

	for ri := startRow; ri <= endRow; ri++ {
		row, err := template.Row(ri)
		if err != nil {
			return fmt.Errorf("get row: %w", err)
		}

		// rendering range property {{range .xxx}}
		flg, err := renderRange(&ri, template, destination, endRow, v)
		if err != nil {
			return err
		}
		if flg {
			// if render range execute, then go next row after range
			continue
		}
		// end rendering range property {{range .xxx}}

		// rendering list property slice or array {{.xxx.yyy}}
		flg, err = renderList(destination, row, v)
		if err != nil {
			return err
		}
		if flg {
			// if render range execute, then go next row after list property
			continue
		}
		// end rendering list property

		// rendering only property
		newRow := destination.AddRow()
		cloneRow(row, newRow)
		if err := renderRow(newRow, v); err != nil {
			return err
		}
		// end rendering only property
	}

	return nil
}
