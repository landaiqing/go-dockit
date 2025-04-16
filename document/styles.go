package document

import (
	"fmt"
)

// Styles 表示Word文档中的样式集合
type Styles struct {
	Styles []*Style
}

// Style 表示Word文档中的样式
type Style struct {
	ID                  string
	Type                string // paragraph, character, table, numbering
	Name                string
	BasedOn             string
	Next                string
	Link                string
	Default             bool
	CustomStyle         bool
	ParagraphProperties *ParagraphProperties
	RunProperties       *RunProperties
	TableProperties     *TableProperties
}

// NewStyles 创建一个新的样式集合
func NewStyles() *Styles {
	return &Styles{
		Styles: make([]*Style, 0),
	}
}

// AddStyle 添加一个样式
func (s *Styles) AddStyle(id, name, styleType string) *Style {
	style := &Style{
		ID:                  id,
		Type:                styleType,
		Name:                name,
		CustomStyle:         true,
		ParagraphProperties: &ParagraphProperties{},
		RunProperties:       &RunProperties{},
	}
	s.Styles = append(s.Styles, style)
	return style
}

// GetStyle 获取指定ID的样式
func (s *Styles) GetStyle(id string) *Style {
	for _, style := range s.Styles {
		if style.ID == id {
			return style
		}
	}
	return nil
}

// SetBasedOn 设置样式的基础样式
func (s *Style) SetBasedOn(basedOn string) *Style {
	s.BasedOn = basedOn
	return s
}

// SetNext 设置样式的下一个样式
func (s *Style) SetNext(next string) *Style {
	s.Next = next
	return s
}

// SetLink 设置样式的链接
func (s *Style) SetLink(link string) *Style {
	s.Link = link
	return s
}

// SetDefault 设置样式是否为默认样式
func (s *Style) SetDefault(isDefault bool) *Style {
	s.Default = isDefault
	return s
}

// SetParagraphProperties 设置段落属性
func (s *Style) SetParagraphProperties(props *ParagraphProperties) *Style {
	s.ParagraphProperties = props
	return s
}

// SetRunProperties 设置文本运行属性
func (s *Style) SetRunProperties(props *RunProperties) *Style {
	s.RunProperties = props
	return s
}

// SetTableProperties 设置表格属性
func (s *Style) SetTableProperties(props *TableProperties) *Style {
	s.TableProperties = props
	return s
}

// ToXML 将样式集合转换为XML
func (s *Styles) ToXML() string {
	xml := "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
	xml += "<w:styles xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">"

	// 添加默认样式
	xml += "<w:docDefaults>"
	xml += "<w:rPrDefault>"
	xml += "<w:rPr>"
	xml += "<w:rFonts w:ascii=\"Calibri\" w:eastAsia=\"Calibri\" w:hAnsi=\"Calibri\" w:cs=\"Calibri\" />"
	xml += "<w:sz w:val=\"22\" />"
	xml += "<w:szCs w:val=\"22\" />"
	xml += "<w:lang w:val=\"en-US\" w:eastAsia=\"en-US\" w:bidi=\"ar-SA\" />"
	xml += "</w:rPr>"
	xml += "</w:rPrDefault>"
	xml += "<w:pPrDefault>"
	xml += "<w:pPr>"
	xml += "<w:spacing w:after=\"200\" w:line=\"276\" w:lineRule=\"auto\" />"
	xml += "</w:pPr>"
	xml += "</w:pPrDefault>"
	xml += "</w:docDefaults>"

	// 添加所有样式
	for _, style := range s.Styles {
		xml += "<w:style w:type=\"" + style.Type + "\" w:styleId=\"" + style.ID + "\">"

		// 样式名称
		xml += "<w:name w:val=\"" + style.Name + "\" />"

		// 基础样式
		if style.BasedOn != "" {
			xml += "<w:basedOn w:val=\"" + style.BasedOn + "\" />"
		}

		// 下一个样式
		if style.Next != "" {
			xml += "<w:next w:val=\"" + style.Next + "\" />"
		}

		// 链接
		if style.Link != "" {
			xml += "<w:link w:val=\"" + style.Link + "\" />"
		}

		// 默认样式
		if style.Default {
			xml += "<w:qFormat />"
		}

		// 自定义样式
		if style.CustomStyle {
			xml += "<w:customStyle w:val=\"1\" />"
		}

		// 段落属性
		if style.ParagraphProperties != nil && style.Type == "paragraph" {
			xml += "<w:pPr>"

			// 对齐方式
			if style.ParagraphProperties.Alignment != "" {
				xml += "<w:jc w:val=\"" + style.ParagraphProperties.Alignment + "\" />"
			}

			// 缩进
			if style.ParagraphProperties.IndentLeft > 0 || style.ParagraphProperties.IndentRight > 0 || style.ParagraphProperties.IndentFirstLine > 0 {
				xml += "<w:ind"
				if style.ParagraphProperties.IndentLeft > 0 {
					xml += " w:left=\"" + fmt.Sprintf("%d", style.ParagraphProperties.IndentLeft) + "\""
				}
				if style.ParagraphProperties.IndentRight > 0 {
					xml += " w:right=\"" + fmt.Sprintf("%d", style.ParagraphProperties.IndentRight) + "\""
				}
				if style.ParagraphProperties.IndentFirstLine > 0 {
					xml += " w:firstLine=\"" + fmt.Sprintf("%d", style.ParagraphProperties.IndentFirstLine) + "\""
				}
				xml += " />"
			}

			// 间距
			if style.ParagraphProperties.SpacingBefore > 0 || style.ParagraphProperties.SpacingAfter > 0 || style.ParagraphProperties.SpacingLine > 0 {
				xml += "<w:spacing"
				if style.ParagraphProperties.SpacingBefore > 0 {
					xml += " w:before=\"" + fmt.Sprintf("%d", style.ParagraphProperties.SpacingBefore) + "\""
				}
				if style.ParagraphProperties.SpacingAfter > 0 {
					xml += " w:after=\"" + fmt.Sprintf("%d", style.ParagraphProperties.SpacingAfter) + "\""
				}
				if style.ParagraphProperties.SpacingLine > 0 {
					xml += " w:line=\"" + fmt.Sprintf("%d", style.ParagraphProperties.SpacingLine) + "\""
					xml += " w:lineRule=\"" + style.ParagraphProperties.SpacingLineRule + "\""
				}
				xml += " />"
			}

			// 分页控制
			if style.ParagraphProperties.KeepNext {
				xml += "<w:keepNext />"
			}
			if style.ParagraphProperties.KeepLines {
				xml += "<w:keepLines />"
			}
			if style.ParagraphProperties.PageBreakBefore {
				xml += "<w:pageBreakBefore />"
			}
			if style.ParagraphProperties.WidowControl {
				xml += "<w:widowControl />"
			}

			// 边框
			if style.ParagraphProperties.BorderTop != nil || style.ParagraphProperties.BorderBottom != nil ||
				style.ParagraphProperties.BorderLeft != nil || style.ParagraphProperties.BorderRight != nil {
				xml += "<w:pBdr>"
				if style.ParagraphProperties.BorderTop != nil {
					xml += "<w:top w:val=\"" + style.ParagraphProperties.BorderTop.Style + "\""
					xml += " w:sz=\"" + fmt.Sprintf("%d", style.ParagraphProperties.BorderTop.Size) + "\""
					xml += " w:space=\"" + fmt.Sprintf("%d", style.ParagraphProperties.BorderTop.Space) + "\""
					xml += " w:color=\"" + style.ParagraphProperties.BorderTop.Color + "\" />"
				}
				if style.ParagraphProperties.BorderBottom != nil {
					xml += "<w:bottom w:val=\"" + style.ParagraphProperties.BorderBottom.Style + "\""
					xml += " w:sz=\"" + fmt.Sprintf("%d", style.ParagraphProperties.BorderBottom.Size) + "\""
					xml += " w:space=\"" + fmt.Sprintf("%d", style.ParagraphProperties.BorderBottom.Space) + "\""
					xml += " w:color=\"" + style.ParagraphProperties.BorderBottom.Color + "\" />"
				}
				if style.ParagraphProperties.BorderLeft != nil {
					xml += "<w:left w:val=\"" + style.ParagraphProperties.BorderLeft.Style + "\""
					xml += " w:sz=\"" + fmt.Sprintf("%d", style.ParagraphProperties.BorderLeft.Size) + "\""
					xml += " w:space=\"" + fmt.Sprintf("%d", style.ParagraphProperties.BorderLeft.Space) + "\""
					xml += " w:color=\"" + style.ParagraphProperties.BorderLeft.Color + "\" />"
				}
				if style.ParagraphProperties.BorderRight != nil {
					xml += "<w:right w:val=\"" + style.ParagraphProperties.BorderRight.Style + "\""
					xml += " w:sz=\"" + fmt.Sprintf("%d", style.ParagraphProperties.BorderRight.Size) + "\""
					xml += " w:space=\"" + fmt.Sprintf("%d", style.ParagraphProperties.BorderRight.Space) + "\""
					xml += " w:color=\"" + style.ParagraphProperties.BorderRight.Color + "\" />"
				}
				xml += "</w:pBdr>"
			}

			// 底纹
			if style.ParagraphProperties.Shading != nil {
				xml += "<w:shd w:val=\"" + style.ParagraphProperties.Shading.Pattern + "\""
				xml += " w:fill=\"" + style.ParagraphProperties.Shading.Fill + "\""
				xml += " w:color=\"" + style.ParagraphProperties.Shading.Color + "\" />"
			}

			xml += "</w:pPr>"
		}

		// 文本运行属性
		if style.RunProperties != nil {
			xml += "<w:rPr>"

			// 字体
			if style.RunProperties.FontFamily != "" {
				xml += "<w:rFonts w:ascii=\"" + style.RunProperties.FontFamily + "\""
				xml += " w:eastAsia=\"" + style.RunProperties.FontFamily + "\""
				xml += " w:hAnsi=\"" + style.RunProperties.FontFamily + "\""
				xml += " w:cs=\"" + style.RunProperties.FontFamily + "\" />"
			}

			// 字号
			if style.RunProperties.FontSize > 0 {
				xml += "<w:sz w:val=\"" + fmt.Sprintf("%d", style.RunProperties.FontSize) + "\" />"
				xml += "<w:szCs w:val=\"" + fmt.Sprintf("%d", style.RunProperties.FontSize) + "\" />"
			}

			// 颜色
			if style.RunProperties.Color != "" {
				xml += "<w:color w:val=\"" + style.RunProperties.Color + "\" />"
			}

			// 粗体
			if style.RunProperties.Bold {
				xml += "<w:b />"
				xml += "<w:bCs />"
			}

			// 斜体
			if style.RunProperties.Italic {
				xml += "<w:i />"
				xml += "<w:iCs />"
			}

			// 下划线
			if style.RunProperties.Underline != "" {
				xml += "<w:u w:val=\"" + style.RunProperties.Underline + "\" />"
			}

			// 删除线
			if style.RunProperties.Strike {
				xml += "<w:strike />"
			}

			// 双删除线
			if style.RunProperties.DoubleStrike {
				xml += "<w:dstrike />"
			}

			// 突出显示颜色
			if style.RunProperties.Highlight != "" {
				xml += "<w:highlight w:val=\"" + style.RunProperties.Highlight + "\" />"
			}

			// 全部大写
			if style.RunProperties.Caps {
				xml += "<w:caps />"
			}

			// 小型大写
			if style.RunProperties.SmallCaps {
				xml += "<w:smallCaps />"
			}

			// 字符间距
			if style.RunProperties.CharacterSpacing != 0 {
				xml += "<w:spacing w:val=\"" + fmt.Sprintf("%d", style.RunProperties.CharacterSpacing) + "\" />"
			}

			// 底纹
			if style.RunProperties.Shading != nil {
				xml += "<w:shd w:val=\"" + style.RunProperties.Shading.Pattern + "\""
				xml += " w:fill=\"" + style.RunProperties.Shading.Fill + "\""
				xml += " w:color=\"" + style.RunProperties.Shading.Color + "\" />"
			}

			// 垂直对齐方式
			if style.RunProperties.VertAlign != "" {
				xml += "<w:vertAlign w:val=\"" + style.RunProperties.VertAlign + "\" />"
			}

			xml += "</w:rPr>"
		}

		// 表格属性
		if style.TableProperties != nil && style.Type == "table" {
			xml += "<w:tblPr>"

			// 表格宽度
			if style.TableProperties.Width > 0 {
				xml += "<w:tblW w:w=\"" + fmt.Sprintf("%d", style.TableProperties.Width) + "\""
				xml += " w:type=\"" + style.TableProperties.WidthType + "\" />"
			}

			// 表格对齐方式
			if style.TableProperties.Alignment != "" {
				xml += "<w:jc w:val=\"" + style.TableProperties.Alignment + "\" />"
			}

			// 表格边框
			if style.TableProperties.Borders != nil {
				xml += "<w:tblBorders>"
				if style.TableProperties.Borders.Top != nil {
					xml += "<w:top w:val=\"" + style.TableProperties.Borders.Top.Style + "\""
					xml += " w:sz=\"" + fmt.Sprintf("%d", style.TableProperties.Borders.Top.Size) + "\""
					xml += " w:space=\"" + fmt.Sprintf("%d", style.TableProperties.Borders.Top.Space) + "\""
					xml += " w:color=\"" + style.TableProperties.Borders.Top.Color + "\" />"
				}
				if style.TableProperties.Borders.Bottom != nil {
					xml += "<w:bottom w:val=\"" + style.TableProperties.Borders.Bottom.Style + "\""
					xml += " w:sz=\"" + fmt.Sprintf("%d", style.TableProperties.Borders.Bottom.Size) + "\""
					xml += " w:space=\"" + fmt.Sprintf("%d", style.TableProperties.Borders.Bottom.Space) + "\""
					xml += " w:color=\"" + style.TableProperties.Borders.Bottom.Color + "\" />"
				}
				if style.TableProperties.Borders.Left != nil {
					xml += "<w:left w:val=\"" + style.TableProperties.Borders.Left.Style + "\""
					xml += " w:sz=\"" + fmt.Sprintf("%d", style.TableProperties.Borders.Left.Size) + "\""
					xml += " w:space=\"" + fmt.Sprintf("%d", style.TableProperties.Borders.Left.Space) + "\""
					xml += " w:color=\"" + style.TableProperties.Borders.Left.Color + "\" />"
				}
				if style.TableProperties.Borders.Right != nil {
					xml += "<w:right w:val=\"" + style.TableProperties.Borders.Right.Style + "\""
					xml += " w:sz=\"" + fmt.Sprintf("%d", style.TableProperties.Borders.Right.Size) + "\""
					xml += " w:space=\"" + fmt.Sprintf("%d", style.TableProperties.Borders.Right.Space) + "\""
					xml += " w:color=\"" + style.TableProperties.Borders.Right.Color + "\" />"
				}
				if style.TableProperties.Borders.InsideH != nil {
					xml += "<w:insideH w:val=\"" + style.TableProperties.Borders.InsideH.Style + "\""
					xml += " w:sz=\"" + fmt.Sprintf("%d", style.TableProperties.Borders.InsideH.Size) + "\""
					xml += " w:space=\"" + fmt.Sprintf("%d", style.TableProperties.Borders.InsideH.Space) + "\""
					xml += " w:color=\"" + style.TableProperties.Borders.InsideH.Color + "\" />"
				}
				if style.TableProperties.Borders.InsideV != nil {
					xml += "<w:insideV w:val=\"" + style.TableProperties.Borders.InsideV.Style + "\""
					xml += " w:sz=\"" + fmt.Sprintf("%d", style.TableProperties.Borders.InsideV.Size) + "\""
					xml += " w:space=\"" + fmt.Sprintf("%d", style.TableProperties.Borders.InsideV.Space) + "\""
					xml += " w:color=\"" + style.TableProperties.Borders.InsideV.Color + "\" />"
				}
				xml += "</w:tblBorders>"
			}

			xml += "</w:tblPr>"
		}

		xml += "</w:style>"
	}

	xml += "</w:styles>"
	return xml
}
