package convert

import (
	"fmt"
	"reflect"
	"time"

	"github.com/spf13/cast"
)

type Converter interface {
	SupportType() reflect.Type
	Convert(v string) (any, error)
}

var defaultConverts = []Converter{
	&TimeConvert{},
}

func CastMapToStruct[T any](data map[string]string, record *T, customConverters ...Converter) error {
	val := reflect.ValueOf(record).Elem()
	for k, v := range data {
		field := val.FieldByName(k)
		if field.IsValid() {
			t := field.Type()
			if t.Kind() == reflect.Ptr {
				t = t.Elem()
			}
			switch t.Kind() {
			case reflect.Int, reflect.Int8, reflect.Int16, reflect.Int32, reflect.Int64:
				v, err := cast.ToInt64E(v)
				if err != nil {
					return err
				}
				field.SetInt(v)
			case reflect.Uint, reflect.Uint8, reflect.Uint16, reflect.Uint32, reflect.Uint64:
				v, err := cast.ToUint64E(v)
				if err != nil {
					return err
				}
				field.SetUint(v)
			case reflect.Float32, reflect.Float64:
				v, err := cast.ToFloat64E(v)
				if err != nil {
					return err
				}
				field.SetFloat(v)
			case reflect.Bool:
				v, err := cast.ToBoolE(v)
				if err != nil {
					return err
				}
				field.SetBool(v)
			case reflect.Struct:
				converter := getConvert(t, customConverters)
				if converter != nil {
					v, err := converter.Convert(v)
					if err != nil {
						return err
					}
					field.Set(reflect.ValueOf(v))
				} else {
					return fmt.Errorf("not support field type %s", t.Kind())
				}
			default:
				field.SetString(v)
			}
		}
	}
	return nil
}

func getConvert(t reflect.Type, customConverters []Converter) Converter {
	for _, convert := range customConverters {
		if t == convert.SupportType() {
			return convert
		}
	}
	for _, convert := range defaultConverts {
		if t == convert.SupportType() {
			return convert
		}
	}
	return nil
}

type TimeConvert struct{}

func (d *TimeConvert) SupportType() reflect.Type {
	return reflect.TypeOf(time.Time{})
}

func (d *TimeConvert) Convert(v string) (any, error) {
	return cast.ToTimeE(v)
}
