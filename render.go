package xlsx_template

import (
	"bytes"
	"errors"
	"strconv"
	"strings"
	"text/template"

	"github.com/tealeg/xlsx/v2"
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
	if from.Height != 0 {
		to.SetHeight(from.Height)
	}
	to.Hidden = from.Hidden
	for _, cell := range from.Cells {
		newCell := to.AddCell()
		cloneCell(cell, newCell)
	}
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

func max(x, y float64) float64 {
	if x < y {
		return y
	}
	return x
}

func renderRow(in *xlsx.Row, ctx interface{}) error {
	var maxEntBefore float64
	var maxEntAfter float64

	for _, cell := range in.Cells {
		countEnt := float64(strings.Count(cell.Value, "\n"))
		maxEntBefore = max(maxEntBefore, countEnt)

		err := renderCell(cell, ctx)
		if err != nil {
			return err
		}

		countEnt = float64(strings.Count(cell.Value, "\n"))
		maxEntAfter = max(maxEntAfter, countEnt)
	}

	maxEntAfter = (maxEntAfter + 1) / (maxEntBefore + 1)
	if maxEntAfter != 0 {
		if in.Height != 0 {
			in.SetHeight(in.Height * maxEntAfter)
		} else {
			in.SetHeight(in.Sheet.SheetFormat.DefaultRowHeight * maxEntAfter)
		}
	}
	return nil
}

func renderCell(cell *xlsx.Cell, ctx interface{}) error {

	var buf bytes.Buffer
	tpl, err := template.New("").Parse(cell.Value)
	if err != nil {
		return err
	}
	buf.Reset()
	err = tpl.Execute(&buf, ctx)
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

func renderRows(sheet *xlsx.Sheet, rows []*xlsx.Row, ctx interface{}) error {

	if isArray(ctx) {
		return errors.New("сtx can not be slice or array")
	}

	for ri := 0; ri < len(rows); ri++ {
		row := rows[ri]

		// rendering range property {{range .xxx}}
		flg, err := renderRange(&ri, sheet, rows, ctx)
		if err != nil {
			return err
		}
		if flg {
			// if render range execute, then go next row after range
			continue
		}
		// end rendering range property {{range .xxx}}

		// rendering list property slice or array {{.xxx.yyy}}
		flg, err = renderList(sheet, row, ctx)
		if err != nil {
			return err
		}
		if flg {
			// if render range execute, then go next row after list property
			continue
		}
		// end rendering list property

		// rendering only property
		newRow := sheet.AddRow()
		cloneRow(row, newRow)
		if err := renderRow(newRow, ctx); err != nil {
			return err
		}
		// end rendering only property
	}

	return nil
}
