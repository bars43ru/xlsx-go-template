package xlsx_template

import (
	"errors"
	"fmt"
	"reflect"
	"strings"
)

func isField(ctx interface{}, field string) error {

	field = strings.ToLower(field)
	tType := reflect.Indirect(reflect.ValueOf(ctx)).Type()
	for i := 0; i < tType.NumField(); i++ {
		if strings.ToLower(tType.Field(i).Name) == field {
			return nil
		}
	}
	return errors.New(fmt.Sprint("Field %s in type %s not found.", tType.Name(), field))
}

func getField(v interface{}, field string) ([]interface{}, error) {

	if err := isField(v, field); err != nil {
		return nil, err
	}

	fValue := reflect.Indirect(reflect.ValueOf(v)).FieldByName(field).Interface()

	retVal := []interface{}{}

	// convert interface to Slice
	vlist := reflect.Indirect(reflect.ValueOf(fValue))
	switch tp := vlist.Kind(); tp {
	case reflect.Slice, reflect.Array:
		for j := 0; j < vlist.Len(); j++ {
			retVal = append(retVal, vlist.Index(j).Interface())
		}
	}

	return retVal, nil
}

func isArray(v interface{}) bool {

	switch tp := reflect.Indirect(reflect.ValueOf(v)).Kind(); tp {
	case reflect.Slice, reflect.Array:
		return true
	}
	return false
}

func isSliceOrArray(v interface{}, field string) (bool, error) {

	if err := isField(v, field); err != nil {
		return false, err
	}

	fValue := reflect.Indirect(reflect.ValueOf(v)).FieldByName(field).Interface()

	return isArray(fValue), nil
}
