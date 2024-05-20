package excelTemplate

import (
	"fmt"
	"github.com/xuri/excelize/v2"
	"log"
	"os"
	"reflect"
	"regexp"
	"strings"
)

var handle = map[reflect.Kind]func(sheet string, reader *excelize.File, position string, val reflect.Value){
	reflect.String: func(sheet string, reader *excelize.File, position string, val reflect.Value) {
		err := reader.SetCellStr(sheet, position, val.String())
		if err != nil {
			log.Fatalf("set value error : %v", val)
		}
	},
	reflect.Int: func(sheet string, reader *excelize.File, position string, val reflect.Value) {
		err := reader.SetCellInt(sheet, position, int(val.Int()))
		if err != nil {
			log.Fatalf("set value error : %v", val)
		}
	},
	reflect.Int32: func(sheet string, reader *excelize.File, position string, val reflect.Value) {
		err := reader.SetCellInt(sheet, position, int(val.Int()))
		if err != nil {
			log.Fatalf("set value error : %v", val)
		}
	},
	reflect.Int64: func(sheet string, reader *excelize.File, position string, val reflect.Value) {
		err := reader.SetCellInt(sheet, position, int(val.Int()))
		if err != nil {
			log.Fatalf("set value error : %v", val)
		}
	},

	reflect.Float64: func(sheet string, reader *excelize.File, position string, val reflect.Value) {
		err := reader.SetCellFloat(sheet, position, val.Float(), 2, 64)
		if err != nil {
			log.Fatalf("set value error : %v", val)
		}
	},
	reflect.Float32: func(sheet string, reader *excelize.File, position string, val reflect.Value) {
		err := reader.SetCellFloat(sheet, position, val.Float(), 2, 64)
		if err != nil {
			log.Fatalf("set value error : %v", val)
		}
	},
}

type loopDataInfo struct {
	RowIndex int
	RefName  string
}

func DoWrite(sourceUrl string, targetUrl string, obj any) {
	compile, _ := regexp.Compile(`\$\{(.*?)}`)
	file, err := os.OpenFile(sourceUrl, os.O_RDWR, 666)
	if err != nil {
		log.Fatal(err)
	}
	reader, err := excelize.OpenReader(file, excelize.Options{
		RawCellValue: true,
	})

	if err != nil {
		log.Fatal(err)
	}
	value := reflect.Indirect(reflect.ValueOf(obj))
	typeof := reflect.TypeOf(obj)
	loopData := make(map[string]loopDataInfo, 10)
	for _, s := range reader.GetSheetList() {
		rows, _ := reader.Rows(s)
		rowIndex := 1
		for rows.Next() {
			columns, _ := rows.Columns()
			for colIndex, val := range columns {
				prefix := "no-prefix"
				if strings.Contains(val, `${fe`) {
					split := strings.Split(val, " ")
					nv := strings.Split(split[1], ":")
					s2 := nv[0]
					prefix = s2 + "."
					loopData[s2] = loopDataInfo{rowIndex, nv[1]}
					pos, err := excelize.CoordinatesToCellName(colIndex+1, rowIndex)
					if err != nil {
						log.Fatal("get cell position error")
					}
					err = reader.SetCellStr(s, pos, "${"+split[len(split)-1]+"}")
					if err != nil {
						log.Fatalf("set value error : %v", val)
					}
				}
				if strings.Contains(val, "ef}") {
					pos, err := excelize.CoordinatesToCellName(colIndex+1, rowIndex)
					if err != nil {
						log.Fatal("get cell position error")
					}
					split := strings.Split(val, " ")
					err = reader.SetCellStr(s, pos, "${"+split[0]+"}")
					if err != nil {
						log.Fatalf("set value error : %v", val)
					}
				}
				if strings.HasPrefix(val, prefix) {
					pos, err := excelize.CoordinatesToCellName(colIndex+1, rowIndex)
					if err != nil {
						log.Fatal("get cell position error")
					}
					cellValue, err := reader.GetCellValue(s, pos)
					if err != nil {
						log.Fatal("get cell value error")
					}

					err = reader.SetCellStr(s, pos, "${"+cellValue+"}")
					if err != nil {
						log.Fatalf("set value error : %v", val)
					}
				}

			}
			rowIndex++
		}

		if typeof.Kind() == reflect.Pointer {
			typeof = typeof.Elem()
		}
		refV := make(map[string]reflect.Value, 10)
		for key, tmpV := range loopData {
			_, ex := typeof.FieldByName(tmpV.RefName)

			if !ex {
				continue
			}
			val := value.FieldByName(tmpV.RefName)

			if val.Kind() != reflect.Array && val.Kind() != reflect.Slice {
				log.Fatalf("type must be slice or array: %v", val)
			}
			if val.IsZero() {
				continue
			}
			refV[key] = val
			for i := 1; i <= val.Len()-1; i++ {
				_ = reader.DuplicateRowTo(s, tmpV.RowIndex, tmpV.RowIndex+i)
			}
		}

		rows, _ = reader.Rows(s)
		rowIndex = 0
		rec := make(map[string]int)
		for rows.Next() {
			columns, _ := rows.Columns()
			for colIndex, cv := range columns {
				if strings.Contains(cv, ".") && strings.Contains(cv, "${") {
					split := strings.Split(strings.ReplaceAll(strings.ReplaceAll(cv, "${", ""), "}", ""), ".")
					info, ex := refV[split[0]]
					if !ex {
						continue
					}
					k, b := rec[split[0]]
					if !b {
						rec[split[0]] = 0
					}
					index := info.Index(k)
					str := index.FieldByName(split[1])
					pos, err := excelize.CoordinatesToCellName(colIndex+1, rowIndex+1)
					if err != nil {
						log.Fatal("get cell position error")
					}
					handle[str.Kind()](s, reader, pos, str)
					continue
				}
				if compile.MatchString(cv) {
					submatch := compile.FindStringSubmatch(cv)
					fi := submatch[1]
					_, b := typeof.FieldByName(fi)
					if !b {
						continue
					}
					pos, err := excelize.CoordinatesToCellName(colIndex+1, rowIndex+1)
					if err != nil {
						log.Fatal("get cell position error")
					}
					name := value.FieldByName(fi)
					handle[name.Kind()](s, reader, pos, name)
				}

			}
			for k, i := range rec {
				rec[k] = i + 1
			}

			rowIndex++
		}

	}

	create, _ := os.Create(targetUrl)
	err = reader.Write(create)
	if err != nil {
		fmt.Println(err)
	}

}
