package document

import (
	"fmt"
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

// AddPageNumber 添加页码域
func (r *Run) AddPageNumber() *Run {
	return r.AddField("begin", " PAGE ")
}

// ToXML 将文本运行转换为XML
func (r *Run) ToXML() string {
	xml := "<w:r>"

	// 添加文本运行属性
	xml += "<w:rPr>"

	// 字体
	if r.Properties.FontFamily != "" {
		xml += "<w:rFonts w:ascii=\"" + r.Properties.FontFamily + "\""
		xml += " w:eastAsia=\"" + r.Properties.FontFamily + "\""
		xml += " w:hAnsi=\"" + r.Properties.FontFamily + "\""
		xml += " w:cs=\"" + r.Properties.FontFamily + "\" />"
	}

	// 字号
	if r.Properties.FontSize > 0 {
		xml += "<w:sz w:val=\"" + fmt.Sprintf("%d", r.Properties.FontSize) + "\" />"
		xml += "<w:szCs w:val=\"" + fmt.Sprintf("%d", r.Properties.FontSize) + "\" />"
	}

	// 颜色
	if r.Properties.Color != "" {
		xml += "<w:color w:val=\"" + r.Properties.Color + "\" />"
	}

	// 粗体
	if r.Properties.Bold {
		xml += "<w:b />"
		xml += "<w:bCs />"
	}

	// 斜体
	if r.Properties.Italic {
		xml += "<w:i />"
		xml += "<w:iCs />"
	}

	// 下划线
	if r.Properties.Underline != "" {
		xml += "<w:u w:val=\"" + r.Properties.Underline + "\" />"
	}

	// 删除线
	if r.Properties.Strike {
		xml += "<w:strike />"
	}

	// 双删除线
	if r.Properties.DoubleStrike {
		xml += "<w:dstrike />"
	}

	// 突出显示颜色
	if r.Properties.Highlight != "" {
		xml += "<w:highlight w:val=\"" + r.Properties.Highlight + "\" />"
	}

	// 全部大写
	if r.Properties.Caps {
		xml += "<w:caps />"
	}

	// 小型大写
	if r.Properties.SmallCaps {
		xml += "<w:smallCaps />"
	}

	// 字符间距
	if r.Properties.CharacterSpacing != 0 {
		xml += "<w:spacing w:val=\"" + fmt.Sprintf("%d", r.Properties.CharacterSpacing) + "\" />"
	}

	// 底纹
	if r.Properties.Shading != nil {
		xml += "<w:shd w:val=\"" + r.Properties.Shading.Pattern + "\""
		xml += " w:fill=\"" + r.Properties.Shading.Fill + "\""
		xml += " w:color=\"" + r.Properties.Shading.Color + "\" />"
	}

	// 上标/下标
	if r.Properties.Superscript {
		xml += "<w:vertAlign w:val=\"superscript\" />"
	} else if r.Properties.Subscript {
		xml += "<w:vertAlign w:val=\"subscript\" />"
	} else if r.Properties.VertAlign != "" {
		xml += "<w:vertAlign w:val=\"" + r.Properties.VertAlign + "\" />"
	}

	// 从右到左文本方向
	if r.Properties.RTL {
		xml += "<w:rtl />"
	}

	// 语言
	if r.Properties.Language != "" {
		xml += "<w:lang w:val=\"" + r.Properties.Language + "\" />"
	}

	xml += "</w:rPr>"

	// 添加分隔符
	if r.BreakType != "" {
		xml += "<w:br w:type=\"" + r.BreakType + "\" />"
	}

	// 添加文本
	if r.Text != "" {
		xml += "<w:t xml:space=\"preserve\">" + r.Text + "</w:t>"
	}

	// 添加图形
	if r.Drawing != nil {
		xml += r.Drawing.ToXML()
	}

	// 添加域
	if r.Field != nil {
		if r.Field.Type == "begin" {
			xml += "<w:fldChar w:fldCharType=\"begin\" />"
		} else if r.Field.Type == "separate" {
			xml += "<w:fldChar w:fldCharType=\"separate\" />"
		} else if r.Field.Type == "end" {
			xml += "<w:fldChar w:fldCharType=\"end\" />"
		}

		if r.Field.Code != "" && r.Field.Type == "begin" {
			// 添加域代码
			xml += "</w:r><w:r><w:instrText xml:space=\"preserve\">" + r.Field.Code + "</w:instrText></w:r><w:r><w:fldChar w:fldCharType=\"separate\" />"
			// 添加域结束标记
			xml += "</w:r><w:r><w:fldChar w:fldCharType=\"end\" />"
		}
	}

	xml += "</w:r>"
	return xml
}
