package document

import (
	"fmt"
	"strings"
)

// Run 表示Word文档中的文本运行
type Run struct {
	Text       string
	Properties *RunProperties
	BreakType  string
	Drawing    *Drawing
	Field      *Field
}

// Field 表示Word文档中的域
type Field struct {
	Type string // begin, separate, end
	Code string // 域代码
}

// RunProperties 表示文本运行的属性
type RunProperties struct {
	Bold             bool     // 粗体
	Italic           bool     // 斜体
	Underline        string   // 下划线类型：single, double, thick, dotted, dash, etc.
	Strike           bool     // 删除线
	DoubleStrike     bool     // 双删除线
	Superscript      bool     // 上标
	Subscript        bool     // 下标
	FontSize         int      // 字号，单位为半点
	FontFamily       string   // 字体
	Color            string   // 颜色，格式为RRGGBB
	Highlight        string   // 突出显示颜色
	Caps             bool     // 全部大写
	SmallCaps        bool     // 小型大写
	CharacterSpacing int      // 字符间距
	Shading          *Shading // 底纹
	VertAlign        string   // 垂直对齐方式：baseline, superscript, subscript
	RTL              bool     // 从右到左文本方向
	Language         string   // 语言
}

// BreakType 表示分隔符类型
const (
	BreakTypePage    = "page"         // 分页符
	BreakTypeColumn  = "column"       // 分栏符
	BreakTypeSection = "section"      // 分节符
	BreakTypeLine    = "textWrapping" // 换行符
)

// NewRun 创建一个新的文本运行
func NewRun() *Run {
	return &Run{
		Text: "",
		Properties: &RunProperties{
			FontSize:   22, // 默认11磅 (22半点)
			FontFamily: "Calibri",
			Color:      "000000",
		},
	}
}

// AddText 向文本运行添加文本
func (r *Run) AddText(text string) *Run {
	r.Text = text
	return r
}

// AddBreak 向文本运行添加分隔符
func (r *Run) AddBreak(breakType string) *Run {
	r.BreakType = breakType
	return r
}

// AddDrawing 向文本运行添加图形
func (r *Run) AddDrawing(drawing *Drawing) *Run {
	r.Drawing = drawing
	return r
}

// SetBold 设置粗体
func (r *Run) SetBold(bold bool) *Run {
	r.Properties.Bold = bold
	return r
}

// SetItalic 设置斜体
func (r *Run) SetItalic(italic bool) *Run {
	r.Properties.Italic = italic
	return r
}

// SetUnderline 设置下划线
func (r *Run) SetUnderline(underline string) *Run {
	r.Properties.Underline = underline
	return r
}

// SetStrike 设置删除线
func (r *Run) SetStrike(strike bool) *Run {
	r.Properties.Strike = strike
	return r
}

// SetDoubleStrike 设置双删除线
func (r *Run) SetDoubleStrike(doubleStrike bool) *Run {
	r.Properties.DoubleStrike = doubleStrike
	return r
}

// SetSuperscript 设置上标
func (r *Run) SetSuperscript(superscript bool) *Run {
	r.Properties.Superscript = superscript
	return r
}

// SetSubscript 设置下标
func (r *Run) SetSubscript(subscript bool) *Run {
	r.Properties.Subscript = subscript
	return r
}

// SetFontSize 设置字号
func (r *Run) SetFontSize(fontSize int) *Run {
	r.Properties.FontSize = fontSize
	return r
}

// SetFontFamily 设置字体
func (r *Run) SetFontFamily(fontFamily string) *Run {
	r.Properties.FontFamily = fontFamily
	return r
}

// SetFontFamilyForRunes 为特定字符设置字体
// 该方法允许为特定的字符序列设置不同的字体
func (r *Run) SetFontFamilyForRunes(fontFamily string, runes []rune) *Run {
	if runes == nil || len(runes) == 0 {
		r.Properties.FontFamily = fontFamily
		return r
	}

	// 将当前文本中的指定字符设置为指定字体
	runeText := string(runes)
	if strings.Contains(r.Text, runeText) {
		// 如果当前文本包含指定字符，设置字体
		// 注意：这是一个简化实现，只是将整个Run的字体设置为指定字体
		// 实际上，为了真正支持混合字体，需要将Run拆分为多个不同字体的Run
		// 但这需要对段落对象的引用，当前结构不支持这种操作
		r.Properties.FontFamily = fontFamily
	}

	return r
}

// SetColor 设置颜色
func (r *Run) SetColor(color string) *Run {
	r.Properties.Color = color
	return r
}

// SetHighlight 设置突出显示颜色
func (r *Run) SetHighlight(highlight string) *Run {
	r.Properties.Highlight = highlight
	return r
}

// SetCaps 设置全部大写
func (r *Run) SetCaps(caps bool) *Run {
	r.Properties.Caps = caps
	return r
}

// SetSmallCaps 设置小型大写
func (r *Run) SetSmallCaps(smallCaps bool) *Run {
	r.Properties.SmallCaps = smallCaps
	return r
}

// SetCharacterSpacing 设置字符间距
func (r *Run) SetCharacterSpacing(spacing int) *Run {
	r.Properties.CharacterSpacing = spacing
	return r
}

// SetShading 设置底纹
func (r *Run) SetShading(fill, color, pattern string) *Run {
	r.Properties.Shading = &Shading{
		Fill:    fill,
		Color:   color,
		Pattern: pattern,
	}
	return r
}

// SetVertAlign 设置垂直对齐方式
func (r *Run) SetVertAlign(vertAlign string) *Run {
	r.Properties.VertAlign = vertAlign
	return r
}

// SetRTL 设置从右到左文本方向
func (r *Run) SetRTL(rtl bool) *Run {
	r.Properties.RTL = rtl
	return r
}

// SetLanguage 设置语言
func (r *Run) SetLanguage(language string) *Run {
	r.Properties.Language = language
	return r
}

// AddField 添加Word域
func (r *Run) AddField(fieldType string, fieldCode string) *Run {
	r.Text = ""
	r.Field = &Field{
		Type: fieldType,
		Code: fieldCode,
	}
	return r
}

// AddTab 添加制表符到文本运行中
func (r *Run) AddTab() *Run {
	// 在OOXML中，制表符被表示为<w:tab/>元素
	// 这里我们使用一个特殊标记，在ToXML时会被替换为制表符标签
	r.Text += "\t"
	return r
}

// AddPageNumber 添加页码域
// Deprecated: 建议使用Document.AddPageNumberParagraph或Footer.AddPageNumber方法
func (r *Run) AddPageNumber() *Run {
	// 注意：此方法生成的页码字段可能不完整
	// 推荐使用Document.AddPageNumberParagraph或Footer.AddPageNumber方法
	// 创建一个完整的页码字段，包含所有需要的部分

	// 返回一个标准的域开始标记
	return r.AddField("begin", "PAGE")
}

// findParagraph 尝试找到当前Run所在的段落
func (r *Run) findParagraph() *Paragraph {
	// 这个方法可能需要访问Document实例来找到包含当前Run的段落
	// 由于目前的结构限制，我们简单地返回nil，表示找不到段落
	// 在实际完整实现中，应该找到文档中包含此Run的段落
	return nil
}

// ToXML 将Run转换为XML
func (r *Run) ToXML() string {
	xml := "<w:r>"

	// 添加运行属性 - 严格按照Word XML规范顺序
	xml += "<w:rPr>"

	// 运行样式（缺失，预留位置）

	// 1. 字体
	if r.Properties.FontFamily != "" {
		xml += fmt.Sprintf("<w:rFonts w:ascii=\"%s\" w:hAnsi=\"%s\" w:eastAsia=\"%s\" w:cs=\"%s\" />",
			r.Properties.FontFamily,
			r.Properties.FontFamily,
			r.Properties.FontFamily,
			r.Properties.FontFamily)
	}

	// 2. 加粗
	if r.Properties.Bold {
		xml += "<w:b />"
		// 3. 复杂文本加粗
		xml += "<w:bCs />"
	}

	// 4. 斜体
	if r.Properties.Italic {
		xml += "<w:i />"
		// 5. 复杂文本斜体
		xml += "<w:iCs />"
	}

	// 6. 全部大写
	if r.Properties.Caps {
		xml += "<w:caps />"
	}

	// 7. 小型大写
	if r.Properties.SmallCaps {
		xml += "<w:smallCaps />"
	}

	// 8. 删除线
	if r.Properties.Strike {
		xml += "<w:strike />"
	}

	// 9. 双删除线
	if r.Properties.DoubleStrike {
		xml += "<w:dstrike />"
	}

	// 10-18. 其他格式（outline, shadow, emboss等，缺失，预留位置）

	// 19. 颜色
	if r.Properties.Color != "" {
		xml += fmt.Sprintf("<w:color w:val=\"%s\" />", r.Properties.Color)
	}

	// 20. 字符间距
	if r.Properties.CharacterSpacing != 0 {
		xml += fmt.Sprintf("<w:spacing w:val=\"%d\" />", r.Properties.CharacterSpacing)
	}

	// 21-23. 其他属性（缺失，预留位置）

	// 24. 字号
	if r.Properties.FontSize > 0 {
		xml += fmt.Sprintf("<w:sz w:val=\"%d\" />", r.Properties.FontSize)
		// 25. 复杂文本字号
		xml += fmt.Sprintf("<w:szCs w:val=\"%d\" />", r.Properties.FontSize)
	}

	// 26. 突出显示
	if r.Properties.Highlight != "" {
		xml += fmt.Sprintf("<w:highlight w:val=\"%s\" />", r.Properties.Highlight)
	}

	// 27. 下划线
	if r.Properties.Underline != "" {
		xml += fmt.Sprintf("<w:u w:val=\"%s\" />", r.Properties.Underline)
	}

	// 28-31. 其他属性（缺失，预留位置）

	// 32. 上标/下标
	if r.Properties.Superscript {
		xml += "<w:vertAlign w:val=\"superscript\" />"
	} else if r.Properties.Subscript {
		xml += "<w:vertAlign w:val=\"subscript\" />"
	} else if r.Properties.VertAlign != "" {
		xml += fmt.Sprintf("<w:vertAlign w:val=\"%s\" />", r.Properties.VertAlign)
	}

	// 33. 从右到左文本方向
	if r.Properties.RTL {
		xml += "<w:rtl />"
	}

	// 34-35. 其他属性（缺失，预留位置）

	// 36. 语言
	if r.Properties.Language != "" {
		xml += fmt.Sprintf("<w:lang w:val=\"%s\" />", r.Properties.Language)
	}

	// 底纹（shd应该在位置30）
	if r.Properties.Shading != nil {
		xml += fmt.Sprintf("<w:shd w:val=\"%s\" w:fill=\"%s\" w:color=\"%s\" />",
			r.Properties.Shading.Pattern,
			r.Properties.Shading.Fill,
			r.Properties.Shading.Color)
	}

	xml += "</w:rPr>"

	// 处理域
	if r.Field != nil {
		switch r.Field.Type {
		case "begin":
			xml += fmt.Sprintf("<w:fldChar w:fldCharType=\"begin\" /><w:instrText>%s</w:instrText>", r.Field.Code)
		case "separate":
			xml += "<w:fldChar w:fldCharType=\"separate\" />"
		case "end":
			xml += "<w:fldChar w:fldCharType=\"end\" />"
		}
	} else if r.Drawing != nil {
		// 添加图形元素
		xml += r.Drawing.ToXML()
	} else if r.BreakType != "" {
		// 添加分隔符
		switch r.BreakType {
		case BreakTypePage:
			xml += "<w:br w:type=\"page\" />"
		case BreakTypeColumn:
			xml += "<w:br w:type=\"column\" />"
		case BreakTypeSection:
			xml += "<w:br w:type=\"sectionContinuous\" />"
		case BreakTypeLine:
			xml += "<w:br />"
		}
	} else {
		// 处理文本内容
		// 特殊处理：制表符
		if r.Text != "" {
			var textParts []string
			for _, char := range r.Text {
				if char == '\t' {
					// 如果是制表符，关闭当前文本，添加制表符标签，然后重新开始文本
					if len(textParts) > 0 {
						xml += fmt.Sprintf("<w:t xml:space=\"preserve\">%s</w:t>", escapeXML(textParts[len(textParts)-1]))
						textParts = textParts[:len(textParts)-1]
					}
					xml += "<w:tab/>"
					textParts = append(textParts, "")
				} else {
					// 如果是普通字符，添加到当前文本部分
					if len(textParts) == 0 {
						textParts = append(textParts, "")
					}
					textParts[len(textParts)-1] += string(char)
				}
			}

			// 处理最后一个文本部分
			for _, part := range textParts {
				if part != "" {
					xml += fmt.Sprintf("<w:t xml:space=\"preserve\">%s</w:t>", escapeXML(part))
				}
			}
		}
	}

	xml += "</w:r>"
	return xml
}

// escapeXML 转义XML文本中的特殊字符
func escapeXML(s string) string {
	s = strings.ReplaceAll(s, "&", "&amp;")
	s = strings.ReplaceAll(s, "<", "&lt;")
	s = strings.ReplaceAll(s, ">", "&gt;")
	s = strings.ReplaceAll(s, "\"", "&quot;")
	s = strings.ReplaceAll(s, "'", "&apos;")
	return s
}
