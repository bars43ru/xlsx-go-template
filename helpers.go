package xlsx_template

import (
	"fmt"
	"reflect"
	"strings"
)

func isField(v any, field string) error {
	field = strings.ToLower(field)
	tType := reflect.Indirect(reflect.ValueOf(v)).Type()
	for i := 0; i < tType.NumField(); i++ {
		if strings.ToLower(tType.Field(i).Name) == field {
			return nil
		}
	}
	return fmt.Errorf("field %s in type %s not found", tType.Name(), field)
}

func getField(v any, field string) ([]any, error) {
	if err := isField(v, field); err != nil {
		return nil, err
	}

	fValue := reflect.Indirect(reflect.ValueOf(v)).FieldByName(field).Interface()

	var items []any
	// convert interface to Slice
	vlist := reflect.Indirect(reflect.ValueOf(fValue))
	switch tp := vlist.Kind(); tp {
	case reflect.Slice, reflect.Array:
		for j := 0; j < vlist.Len(); j++ {
			items = append(items, vlist.Index(j).Interface())
		}
	}

	return items, nil
}

func isArray(v any) bool {

	switch tp := reflect.Indirect(reflect.ValueOf(v)).Kind(); tp {
	case reflect.Slice, reflect.Array:
		return true
	}
	return false
}

func isSliceOrArray(v any, field string) (bool, error) {

	if err := isField(v, field); err != nil {
		return false, err
	}

	fValue := reflect.Indirect(reflect.ValueOf(v)).FieldByName(field).Interface()

	return isArray(fValue), nil
}
