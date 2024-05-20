package excelTemplate

import (
	"github.com/xuri/excelize/v2"
	"log"
	"os"
	"reflect"
	"slices"
	"strconv"
)

func ReaderExcel[T any](path string, sheetName string) []T {

	file, err := os.OpenFile(path, os.O_RDWR, 666)
	if err != nil {
		log.Fatal(err)
	}
	defer file.Close()
	reader, err := excelize.OpenReader(file, excelize.Options{
		RawCellValue: true,
	})
	defer reader.Close()
	if err != nil {
		log.Fatal("open excel file error")
	}
	if sheetName == "" {
		sheetName = "Sheet1"
	}

	rows, err := reader.Rows(sheetName)
	if err != nil {
		log.Fatal(err)
	}

	ts := make([]T, 0, 16)
	fieldNames := make([]string, 0, 16)
	rowIndex := 0
	for rows.Next() {
		columns, _ := rows.Columns()
		if rowIndex == 0 {
			fieldNames = append(fieldNames, columns...)
			rowIndex++
			continue
		}
		var tmp T
		value := reflect.ValueOf(&tmp).Elem()
		types := reflect.TypeOf(tmp)

		for i := 0; i < types.NumField(); i++ {
			fieldType := types.Field(i)
			tag := fieldType.Tag.Get("excelTemplate")
			idx := slices.Index(fieldNames, tag)

			if idx < 0 {
				continue
			}
			pos, err := excelize.CoordinatesToCellName(idx+1, rowIndex+1)
			if err != nil {
				log.Fatal(err)
			}
			cellValue, err := reader.GetCellValue(sheetName, pos)
			if err != nil {
				log.Fatal(err)
			}
			val := value.FieldByName(fieldType.Name)
			switch fieldType.Type.Kind() {
			case reflect.String:
				val.SetString(cellValue)
			case reflect.Int64, reflect.Int32:
				parseInt, err := strconv.ParseInt(cellValue, 10, 64)
				if err != nil {
					log.Fatal(err)
				}
				val.SetInt(parseInt)
			case reflect.Float32, reflect.Float64:
				float, err := strconv.ParseFloat(cellValue, 64)
				if err != nil {
					log.Fatal(err)
				}
				val.SetFloat(float)
			}

		}
		rowIndex++
		ts = append(ts, tmp)
	}

	return ts

}
