package xlsx_template

import (
	"errors"
	"io"

	"github.com/tealeg/xlsx/v2"
)

// Xlst template struct
type XlstTemplate struct {
	template *xlsx.File
	report   *xlsx.File
}

// New XlstTemplate
func New() *XlstTemplate {
	return &XlstTemplate{}
}

// Render report it ctx a struct
func (this *XlstTemplate) Render(ctx interface{}) error {
	report := xlsx.NewFile()
	for _, templSheet := range this.template.Sheets {
		repSheet, err := report.AddSheet("NewSheet")
		if err != nil {
			return err
		}

		repSheet.Name = templSheet.Name

		cloneSheet(templSheet, repSheet)

		err = renderRows(repSheet, templSheet.Rows, ctx)
		if err != nil {
			return err
		}
	}
	this.report = report
	return nil
}

// ReadTemplate reads template from disk
func (this *XlstTemplate) ReadTemplate(path string) error {
	file, err := xlsx.OpenFile(path)
	if err != nil {
		return err
	}
	this.template = file
	return nil
}

// OpenBinary reads template from bytes
func (this *XlstTemplate) OpenBinary(bs []byte) error {
	file, err := xlsx.OpenBinary(bs)
	if err != nil {
		return err
	}
	this.template = file
	return nil
}

// Save saves generated report to disk
func (this *XlstTemplate) Save(path string) error {
	if this.report == nil {
		return errors.New("Report was not generated")
	}
	return this.report.Save(path)
}

// Write writes generated report to provided writer
func (this *XlstTemplate) Write(writer io.Writer) error {
	if this.report == nil {
		return errors.New("Report was not generated")
	}
	return this.report.Write(writer)
}
