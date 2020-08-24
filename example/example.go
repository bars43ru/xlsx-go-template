package main

import (
	xlsx_template "github.com/bars43ru/xlsx-go-template"
	"time"
)

type report struct {
	Date       time.Time
	Items      []item
	TotalField totalField
}

type item struct {
	Field1 string
	Field2 int
	Field3 int
	Field4 float64
}

type totalField struct {
	Field string
	Value int
}

func (r report) TotalField4() (retValue float64) {
	retValue = 0
	for _, val := range r.Items {
		retValue += val.Field4
	}
	return
}

func main() {
	doc := xlsx_template.New()
	if err := doc.ReadTemplate("./template.xlsx"); err != nil {
		panic(err)
	}

	ctx := prepareTestData()

	err := doc.Render(ctx)
	if err != nil {
		panic(err)
	}

	err = doc.Save("./result.xlsx")
	if err != nil {
		panic(err)
	}
}

func prepareTestData() (retValue *report) {
	retValue = &report{
		Date: time.Now(),
		Items: []item{
			{
				Field1: "Value1",
				Field2: 1,
				Field3: 2,
				Field4: 45.45,
			},
			{
				Field1: "Value2",
				Field2: 2,
				Field3: 7,
				Field4: 459.987,
			},
			{
				Field1: "Value\nResult",
				Field2: 888,
				Field3: 0,
				Field4: 0,
			},
		},
		TotalField: totalField{
			Field: "Field3",
			Value: 9,
		},
	}

	return
}
