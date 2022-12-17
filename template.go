package xlsx_template

import (
	"errors"
	"io"

	"github.com/tealeg/xlsx/v2"
)

// Template struct
type Template struct {
	template *xlsx.File
	report   *xlsx.File
}

// Deprecated: using Template
type XlstTemplate = Template

// New Template
func New() *Template {
	return &Template{}
}

// Render report it v a struct
func (t *Template) Render(v any) error {
	report := xlsx.NewFile()
	for _, sheet := range t.template.Sheets {
		repSheet, err := report.AddSheet("NewSheet")
		if err != nil {
			return err
		}

		repSheet.Name = sheet.Name

		cloneSheet(sheet, repSheet)

		err = renderRows(repSheet, sheet.Rows, v)
		if err != nil {
			return err
		}
	}
	t.report = report
	return nil
}

// ReadTemplate reads template from disk
func (t *Template) ReadTemplate(path string) error {
	file, err := xlsx.OpenFile(path)
	if err != nil {
		return err
	}
	t.template = file
	return nil
}

// OpenBinary reads template from bytes
func (t *Template) OpenBinary(b []byte) error {
	file, err := xlsx.OpenBinary(b)
	if err != nil {
		return err
	}
	t.template = file
	return nil
}

// Save saves generated report to disk
func (t *Template) Save(path string) error {
	if t.report == nil {
		return errors.New("Report was not generated")
	}
	return t.report.Save(path)
}

// Write writes generated report to provided writer
func (t *Template) Write(writer io.Writer) error {
	if t.report == nil {
		return errors.New("Report was not generated")
	}
	return t.report.Write(writer)
}
