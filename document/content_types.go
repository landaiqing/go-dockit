package document

import (
	"fmt"
)

// ContentTypes 表示Word文档中的内容类型集合
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
	ct.AddDefault("gif", "image/gif")
	ct.AddDefault("bmp", "image/bmp")
	ct.AddDefault("tiff", "image/tiff")
	ct.AddDefault("tif", "image/tiff")
	ct.AddDefault("wmf", "image/x-wmf")
	ct.AddDefault("emf", "image/x-emf")

	// 添加覆盖的内容类型
	ct.AddOverride("/document/document.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml")
	ct.AddOverride("/document/styles.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml")
	ct.AddOverride("/document/numbering.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml")
	ct.AddOverride("/document/settings.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml")
	ct.AddOverride("/document/theme/theme1.xml", "application/vnd.openxmlformats-officedocument.theme+xml")
	ct.AddOverride("/docProps/core.xml", "application/vnd.openxmlformats-package.core-properties+xml")
	ct.AddOverride("/docProps/app.xml", "application/vnd.openxmlformats-officedocument.extended-properties+xml")

	return ct
}

// AddDefault 添加一个默认的内容类型
func (c *ContentTypes) AddDefault(extension, contentType string) *Default {
	def := &Default{
		Extension:   extension,
		ContentType: contentType,
	}
	c.Defaults = append(c.Defaults, def)
	return def
}

// AddOverride 添加一个覆盖的内容类型
func (c *ContentTypes) AddOverride(partName, contentType string) *Override {
	override := &Override{
		PartName:    partName,
		ContentType: contentType,
	}
	c.Overrides = append(c.Overrides, override)
	return override
}

// AddHeaderOverride 添加一个页眉的内容类型
func (c *ContentTypes) AddHeaderOverride(index int) *Override {
	return c.AddOverride(
		"/document/header"+fmt.Sprintf("%d", index)+".xml",
		"application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml",
	)
}

// AddFooterOverride 添加一个页脚的内容类型
func (c *ContentTypes) AddFooterOverride(index int) *Override {
	return c.AddOverride(
		"/document/footer"+fmt.Sprintf("%d", index)+".xml",
		"application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml",
	)
}

// ToXML 将内容类型集合转换为XML
func (c *ContentTypes) ToXML() string {
	xml := "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
	xml += "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"

	// 添加所有默认的内容类型
	for _, def := range c.Defaults {
		xml += "<Default Extension=\"" + def.Extension + "\""
		xml += " ContentType=\"" + def.ContentType + "\" />"
	}

	// 添加所有覆盖的内容类型
	for _, override := range c.Overrides {
		xml += "<Override PartName=\"" + override.PartName + "\""
		xml += " ContentType=\"" + override.ContentType + "\" />"
	}

	xml += "</Types>"
	return xml
}
