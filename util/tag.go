package util

import "reflect"

func ParseTag(s any, tag string) []string {
	t := reflect.TypeOf(s).Elem()
	var tags []string
	for i := 0; i < t.NumField(); i++ {
		field := t.Field(i)
		sourceTag := field.Tag.Get(tag)
		if sourceTag == "" {
			sourceTag = field.Name
		}
		tags = append(tags, sourceTag)
	}
	return tags
}

func ParseTagFieldMap(s any, tag string) map[string]string {
	tagMap := make(map[string]string)
	t := reflect.TypeOf(s).Elem()
	for i := 0; i < t.NumField(); i++ {
		field := t.Field(i)
		fieldName := field.Name
		sourceTag := field.Tag.Get(tag)
		if sourceTag == "" {
			sourceTag = fieldName
		}
		tagMap[sourceTag] = fieldName
	}
	return tagMap
}
