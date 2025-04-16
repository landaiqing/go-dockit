package workbook

import (
	"fmt"
)

// ContentTypes 表示Excel文档中的内容类型集合
type ContentTypes struct {
	Defaults  []*Default
	Overrides []*Override
}

// Default 表示默认的内容类型
type Default struct {
	Extension   string
	ContentType string
}

// Override 表示覆盖的内容类型
type Override struct {
	PartName    string
	ContentType string
}

// NewContentTypes 创建一个新的内容类型集合
func NewContentTypes() *ContentTypes {
	ct := &ContentTypes{
		Defaults:  make([]*Default, 0),
		Overrides: make([]*Override, 0),
	}

	// 添加默认的内容类型
	ct.AddDefault("xml", "application/xml")
	ct.AddDefault("rels", "application/vnd.openxmlformats-package.relationships+xml")
	ct.AddDefault("png", "image/png")
	ct.AddDefault("jpeg", "image/jpeg")
	ct.AddDefault("jpg", "image/jpeg")

	// 添加覆盖的内容类型
	ct.AddOverride("/xl/workbook.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml")
	ct.AddOverride("/xl/styles.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml")
	ct.AddOverride("/xl/theme/theme1.xml", "application/vnd.openxmlformats-officedocument.theme+xml")
	ct.AddOverride("/xl/sharedStrings.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml")

	return ct
}

// AddDefault 添加一个默认的内容类型
func (ct *ContentTypes) AddDefault(extension, contentType string) *Default {
	def := &Default{
		Extension:   extension,
		ContentType: contentType,
	}
	ct.Defaults = append(ct.Defaults, def)
	return def
}

// AddOverride 添加一个覆盖的内容类型
func (ct *ContentTypes) AddOverride(partName, contentType string) *Override {
	ovr := &Override{
		PartName:    partName,
		ContentType: contentType,
	}
	ct.Overrides = append(ct.Overrides, ovr)
	return ovr
}

// AddWorksheetOverride 添加工作表的内容类型覆盖
func (ct *ContentTypes) AddWorksheetOverride(index int) *Override {
	partName := fmt.Sprintf("/xl/worksheets/sheet%d.xml", index)
	return ct.AddOverride(partName, "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml")
}

// ToXML 将内容类型转换为XML
func (ct *ContentTypes) ToXML() string {
	xml := "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
	xml += "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">\n"

	// 添加默认的内容类型
	for _, def := range ct.Defaults {
		xml += fmt.Sprintf("  <Default Extension=\"%s\" ContentType=\"%s\"/>\n", def.Extension, def.ContentType)
	}

	// 添加覆盖的内容类型
	for _, ovr := range ct.Overrides {
		xml += fmt.Sprintf("  <Override PartName=\"%s\" ContentType=\"%s\"/>\n", ovr.PartName, ovr.ContentType)
	}

	xml += "</Types>"
	return xml
}
