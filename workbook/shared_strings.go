package workbook

import (
	"fmt"
	"strings"
)

// SharedStrings 表示Excel文档中的共享字符串表
type SharedStrings struct {
	Strings []string
	Map     map[string]int
}

// NewSharedStrings 创建一个新的共享字符串表
func NewSharedStrings() *SharedStrings {
	return &SharedStrings{
		Strings: make([]string, 0),
		Map:     make(map[string]int),
	}
}

// AddString 添加一个字符串到共享字符串表，并返回其索引
func (ss *SharedStrings) AddString(s string) int {
	// 检查字符串是否已存在
	if index, ok := ss.Map[s]; ok {
		return index
	}

	// 添加新字符串
	index := len(ss.Strings)
	ss.Strings = append(ss.Strings, s)
	ss.Map[s] = index
	return index
}

// GetString 根据索引获取字符串
func (ss *SharedStrings) GetString(index int) (string, error) {
	if index < 0 || index >= len(ss.Strings) {
		return "", fmt.Errorf("index out of range: %d", index)
	}
	return ss.Strings[index], nil
}

// Count 返回共享字符串表中的字符串数量
func (ss *SharedStrings) Count() int {
	return len(ss.Strings)
}

// ToXML 将共享字符串表转换为XML
func (ss *SharedStrings) ToXML() string {
	xml := "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
	xml += fmt.Sprintf("<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"%d\" uniqueCount=\"%d\">\n",
		ss.Count(), ss.Count())

	for _, s := range ss.Strings {
		// 转义XML特殊字符
		s = strings.Replace(s, "&", "&amp;", -1)
		s = strings.Replace(s, "<", "&lt;", -1)
		s = strings.Replace(s, ">", "&gt;", -1)
		s = strings.Replace(s, "\"", "&quot;", -1)
		s = strings.Replace(s, "'", "&apos;", -1)

		xml += "  <si><t>" + s + "</t></si>\n"
	}

	xml += "</sst>"
	return xml
}
