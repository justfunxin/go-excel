package excel

import (
	"fmt"
	"io"
	"log"
	"mime/multipart"
	"reflect"
	"strconv"
	"strings"
	"time"

	"github.com/kangarooxin/go-tools/convert"
	"github.com/kangarooxin/go-tools/tags"
	"github.com/spf13/cast"
	"github.com/xuri/excelize/v2"
)

func GetRowsFromMultipart[T any](file *multipart.FileHeader, records *[]T) error {
	return GetRowsFromMultipartBySheetIndex(file, 0, records)
}

func GetRowsFromMultipartBySheetIndex[T any](file *multipart.FileHeader, index int, records *[]T) error {
	reader, err := file.Open()
	if err != nil {
		return err
	}
	defer func() {
		if err := reader.Close(); err != nil {
			log.Println("close failed:", err)
		}
	}()
	return GetRowsFromReaderBySheetIndex(reader, index, records)
}

func GetRowsFromMultipartBySheetName[T any](file *multipart.FileHeader, sheet string, records *[]T) error {
	reader, err := file.Open()
	if err != nil {
		return err
	}
	defer func() {
		if err := reader.Close(); err != nil {
			log.Println("close failed:", err)
		}
	}()
	return GetRowsFromReaderBySheetName(reader, sheet, records)
}

func GetRowsFromFile[T any](filepath string, records *[]T) error {
	return GetRowsFromFileBySheetIndex(filepath, 0, records)
}

func GetRowsFromFileBySheetIndex[T any](filepath string, index int, records *[]T) error {
	excelFile, err := excelize.OpenFile(filepath)
	if err != nil {
		return err
	}
	defer func() {
		if err := excelFile.Close(); err != nil {
			log.Println("close failed:", err)
		}
	}()
	return GetRowsBySheetIndex(excelFile, index, records)
}

func GetRowsFromFileBySheetName[T any](filepath string, sheet string, records *[]T) error {
	excelFile, err := excelize.OpenFile(filepath)
	if err != nil {
		return err
	}
	defer func() {
		if err := excelFile.Close(); err != nil {
			log.Println("close failed:", err)
		}
	}()
	return GetRowsBySheetName(excelFile, sheet, records)
}

func GetRowsFromReader[T any](filepath string, records *[]T) error {
	return GetRowsFromFileBySheetIndex(filepath, 0, records)
}

func GetRowsFromReaderBySheetIndex[T any](r io.Reader, index int, records *[]T) error {
	excelFile, err := excelize.OpenReader(r)
	if err != nil {
		return err
	}
	defer func() {
		if err := excelFile.Close(); err != nil {
			log.Println("close failed:", err)
		}
	}()
	return GetRowsBySheetIndex(excelFile, index, records)
}

func GetRowsFromReaderBySheetName[T any](r io.Reader, sheet string, records *[]T) error {
	excelFile, err := excelize.OpenReader(r)
	if err != nil {
		return err
	}
	defer func() {
		if err := excelFile.Close(); err != nil {
			log.Println("close failed:", err)
		}
	}()
	return GetRowsBySheetName(excelFile, sheet, records)
}

func GetRows[T any](file *excelize.File, records *[]T) error {
	return GetRowsBySheetIndex(file, 0, records)
}

func GetRowsBySheetIndex[T any](file *excelize.File, index int, records *[]T) error {
	sheet := file.GetSheetName(index)
	return GetRowsBySheetName(file, sheet, records)
}

func GetRowsBySheetName[T any](file *excelize.File, sheet string, records *[]T) error {
	rows, err := file.GetRows(sheet, excelize.Options{RawCellValue: true})
	if err != nil {
		return err
	}
	tagMap := tags.ParseTagFieldMap(new(T), "xlsx")
	var header []string
	for _, col := range rows[0] {
		header = append(header, tagMap[col])
	}
	for _, row := range rows[1:] {
		record, err := convertRow[T](row, header)
		if err != nil {
			return err
		}
		*records = append(*records, *record)
	}
	return nil
}

func NewFile[T any](records *[]T) (*excelize.File, error) {
	return NewFileWithSheetName("Sheet1", records)
}

func NewFileWithSheetName[T any](sheetName string, records *[]T) (*excelize.File, error) {
	f := excelize.NewFile()
	_, err := f.NewSheet(sheetName)
	if err != nil {
		return nil, err
	}
	// 写入标题
	header := tags.ParseToSlice(new(T), "xlsx")
	for i, v := range header {
		err = WriteCellValue(f, sheetName, i+1, 1, v)
		if err != nil {
			return nil, err
		}
	}
	// 写入数据
	for i, record := range *records {
		row := i + 2
		val := reflect.ValueOf(record)
		for j := 0; j < val.NumField(); j++ {
			field := val.Field(j)
			err = WriteCellValue(f, sheetName, j+1, row, getFieldValue(field))
			if err != nil {
				return nil, err
			}
		}
	}
	return f, nil
}

func WriteCellValue(f *excelize.File, sheet string, col, row int, value interface{}) error {
	cell := GetCellName(col, row)
	err := f.SetCellValue(sheet, cell, value)
	if err != nil {
		return err
	}
	return nil
}

func GetCellName(col, row int) string {
	colName, err := excelize.ColumnNumberToName(col)
	if err != nil {
		return ""
	}
	return colName + strconv.Itoa(row)
}

func getFieldValue(field reflect.Value) interface{} {
	value := field.Interface()
	switch field.Kind() {
	case reflect.Array, reflect.Slice:
		var strArray []string
		for i := 0; i < field.Len(); i++ {
			strArray = append(strArray, fmt.Sprintf("%v", field.Index(i)))
		}
		return strings.Join(strArray, ",")
	default:
		return value
	}
}

// convert 将一行数据转换成结构体
func convertRow[T any](row, header []string) (*T, error) {
	data := make(map[string]string)
	for i, col := range row {
		field := header[i]
		if field == "" {
			continue
		}
		data[field] = col
	}
	record := new(T)
	err := castMapToStruct(data, record)
	if err != nil {
		return nil, err
	}
	return record, nil
}

type TimeConvert struct{}

func (d *TimeConvert) SupportType() reflect.Type {
	return reflect.TypeOf(time.Time{})
}

func (d *TimeConvert) Convert(v string) (any, error) {
	return excelize.ExcelDateToTime(cast.ToFloat64(v), false)
}

func castMapToStruct[T any](data map[string]string, record *T) error {
	return convert.CastMapToStruct(data, record, &TimeConvert{})
}
