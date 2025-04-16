package document

import "fmt"

// Paragraph 表示Word文档中的段落
type Paragraph struct {
	Runs       []*Run
	Properties *ParagraphProperties
}

// ParagraphProperties 表示段落的属性
type ParagraphProperties struct {
	Alignment       string // left, center, right, justified
	IndentLeft      int    // 左缩进，单位为twip (1/20 point)
	IndentRight     int    // 右缩进
	IndentFirstLine int    // 首行缩进
	SpacingBefore   int    // 段前间距
	SpacingAfter    int    // 段后间距
	SpacingLine     int    // 行间距
	SpacingLineRule string // auto, exact, atLeast
	KeepNext        bool   // 与下段同页
	KeepLines       bool   // 段中不分页
	PageBreakBefore bool   // 段前分页
	WidowControl    bool   // 孤行控制
	OutlineLevel    int    // 大纲级别
	StyleID         string // 样式ID
	NumID           int    // 编号ID
	NumLevel        int    // 编号级别
	BorderTop       *Border
	BorderBottom    *Border
	BorderLeft      *Border
	BorderRight     *Border
	Shading         *Shading
}

// Border 表示边框
type Border struct {
	Style string // single, double, dotted, dashed, etc.
	Size  int    // 边框宽度，单位为1/8点
	Color string // 边框颜色，格式为RRGGBB
	Space int    // 边框与文本的距离
}

// Shading 表示底纹
type Shading struct {
	Fill    string // 填充颜色
	Color   string // 文本颜色
	Pattern string // 底纹样式
}

// NewParagraph 创建一个新的段落
func NewParagraph() *Paragraph {
	return &Paragraph{
		Runs: make([]*Run, 0),
		Properties: &ParagraphProperties{
			Alignment:       "left",
			WidowControl:    true,
			SpacingLineRule: "auto",
		},
	}
}

// AddRun 向段落添加一个文本运行并返回它
func (p *Paragraph) AddRun() *Run {
	r := NewRun()
	p.Runs = append(p.Runs, r)
	return r
}

// AddText 向段落添加文本
func (p *Paragraph) AddText(text string) *Run {
	return p.AddRun().AddText(text)
}

// SetAlignment 设置段落对齐方式
func (p *Paragraph) SetAlignment(alignment string) *Paragraph {
	p.Properties.Alignment = alignment
	return p
}

// SetIndentLeft 设置左缩进
func (p *Paragraph) SetIndentLeft(indent int) *Paragraph {
	p.Properties.IndentLeft = indent
	return p
}

// SetIndentRight 设置右缩进
func (p *Paragraph) SetIndentRight(indent int) *Paragraph {
	p.Properties.IndentRight = indent
	return p
}

// SetIndentFirstLine 设置首行缩进
func (p *Paragraph) SetIndentFirstLine(indent int) *Paragraph {
	p.Properties.IndentFirstLine = indent
	return p
}

// SetSpacingBefore 设置段前间距
func (p *Paragraph) SetSpacingBefore(spacing int) *Paragraph {
	p.Properties.SpacingBefore = spacing
	return p
}

// SetSpacingAfter 设置段后间距
func (p *Paragraph) SetSpacingAfter(spacing int) *Paragraph {
	p.Properties.SpacingAfter = spacing
	return p
}

// SetSpacingLine 设置行间距
func (p *Paragraph) SetSpacingLine(spacing int, rule string) *Paragraph {
	p.Properties.SpacingLine = spacing
	p.Properties.SpacingLineRule = rule
	return p
}

// SetKeepNext 设置与下段同页
func (p *Paragraph) SetKeepNext(keepNext bool) *Paragraph {
	p.Properties.KeepNext = keepNext
	return p
}

// SetKeepLines 设置段中不分页
func (p *Paragraph) SetKeepLines(keepLines bool) *Paragraph {
	p.Properties.KeepLines = keepLines
	return p
}

// SetPageBreakBefore 设置段前分页
func (p *Paragraph) SetPageBreakBefore(pageBreakBefore bool) *Paragraph {
	p.Properties.PageBreakBefore = pageBreakBefore
	return p
}

// SetStyleID 设置样式ID
func (p *Paragraph) SetStyleID(styleID string) *Paragraph {
	p.Properties.StyleID = styleID
	return p
}

// SetNumbering 设置编号
func (p *Paragraph) SetNumbering(numID, numLevel int) *Paragraph {
	p.Properties.NumID = numID
	p.Properties.NumLevel = numLevel
	return p
}

// SetBorder 设置边框
func (p *Paragraph) SetBorder(position string, style string, size int, color string, space int) *Paragraph {
	border := &Border{
		Style: style,
		Size:  size,
		Color: color,
		Space: space,
	}

	switch position {
	case "top":
		p.Properties.BorderTop = border
	case "bottom":
		p.Properties.BorderBottom = border
	case "left":
		p.Properties.BorderLeft = border
	case "right":
		p.Properties.BorderRight = border
	}

	return p
}

// SetShading 设置底纹
func (p *Paragraph) SetShading(fill, color, pattern string) *Paragraph {
	p.Properties.Shading = &Shading{
		Fill:    fill,
		Color:   color,
		Pattern: pattern,
	}
	return p
}

// ToXML 将段落转换为XML
func (p *Paragraph) ToXML() string {
	xml := "<w:p>"

	// 添加段落属性
	xml += "<w:pPr>"

	// 样式 (必须在最前面)
	if p.Properties.StyleID != "" {
		xml += fmt.Sprintf("<w:pStyle w:val=\"%s\" />", p.Properties.StyleID)
	}

	// 分页控制 (必须在编号前面)
	if p.Properties.KeepNext {
		xml += "<w:keepNext />"
	}
	if p.Properties.KeepLines {
		xml += "<w:keepLines />"
	}
	if p.Properties.PageBreakBefore {
		xml += "<w:pageBreakBefore />"
	}
	if p.Properties.WidowControl {
		xml += "<w:widowControl />"
	}

	// 编号
	if p.Properties.NumID > 0 {
		xml += "<w:numPr>"
		xml += fmt.Sprintf("<w:ilvl w:val=\"%d\" />", p.Properties.NumLevel)
		xml += fmt.Sprintf("<w:numId w:val=\"%d\" />", p.Properties.NumID)
		xml += "</w:numPr>"
	}

	// 边框
	if p.Properties.BorderTop != nil || p.Properties.BorderBottom != nil ||
		p.Properties.BorderLeft != nil || p.Properties.BorderRight != nil {
		xml += "<w:pBdr>"
		if p.Properties.BorderTop != nil {
			xml += fmt.Sprintf("<w:top w:val=\"%s\" w:sz=\"%d\" w:space=\"%d\" w:color=\"%s\" />",
				p.Properties.BorderTop.Style,
				p.Properties.BorderTop.Size,
				p.Properties.BorderTop.Space,
				p.Properties.BorderTop.Color)
		}
		if p.Properties.BorderBottom != nil {
			xml += fmt.Sprintf("<w:bottom w:val=\"%s\" w:sz=\"%d\" w:space=\"%d\" w:color=\"%s\" />",
				p.Properties.BorderBottom.Style,
				p.Properties.BorderBottom.Size,
				p.Properties.BorderBottom.Space,
				p.Properties.BorderBottom.Color)
		}
		if p.Properties.BorderLeft != nil {
			xml += fmt.Sprintf("<w:left w:val=\"%s\" w:sz=\"%d\" w:space=\"%d\" w:color=\"%s\" />",
				p.Properties.BorderLeft.Style,
				p.Properties.BorderLeft.Size,
				p.Properties.BorderLeft.Space,
				p.Properties.BorderLeft.Color)
		}
		if p.Properties.BorderRight != nil {
			xml += fmt.Sprintf("<w:right w:val=\"%s\" w:sz=\"%d\" w:space=\"%d\" w:color=\"%s\" />",
				p.Properties.BorderRight.Style,
				p.Properties.BorderRight.Size,
				p.Properties.BorderRight.Space,
				p.Properties.BorderRight.Color)
		}
		xml += "</w:pBdr>"
	}

	// 底纹
	if p.Properties.Shading != nil {
		xml += fmt.Sprintf("<w:shd w:val=\"%s\" w:fill=\"%s\" w:color=\"%s\" />",
			p.Properties.Shading.Pattern,
			p.Properties.Shading.Fill,
			p.Properties.Shading.Color)
	}

	// 间距
	if p.Properties.SpacingBefore > 0 || p.Properties.SpacingAfter > 0 || p.Properties.SpacingLine > 0 {
		xml += "<w:spacing"
		if p.Properties.SpacingBefore > 0 {
			xml += fmt.Sprintf(" w:before=\"%d\"", p.Properties.SpacingBefore)
		}
		if p.Properties.SpacingAfter > 0 {
			xml += fmt.Sprintf(" w:after=\"%d\"", p.Properties.SpacingAfter)
		}
		if p.Properties.SpacingLine > 0 {
			xml += fmt.Sprintf(" w:line=\"%d\"", p.Properties.SpacingLine)
			xml += fmt.Sprintf(" w:lineRule=\"%s\"", p.Properties.SpacingLineRule)
		}
		xml += " />"
	}

	// 缩进
	if p.Properties.IndentLeft > 0 || p.Properties.IndentRight > 0 || p.Properties.IndentFirstLine > 0 {
		xml += "<w:ind"
		if p.Properties.IndentLeft > 0 {
			xml += fmt.Sprintf(" w:left=\"%d\"", p.Properties.IndentLeft)
		}
		if p.Properties.IndentRight > 0 {
			xml += fmt.Sprintf(" w:right=\"%d\"", p.Properties.IndentRight)
		}
		if p.Properties.IndentFirstLine > 0 {
			xml += fmt.Sprintf(" w:firstLine=\"%d\"", p.Properties.IndentFirstLine)
		}
		xml += " />"
	}

	// 对齐方式
	if p.Properties.Alignment != "" {
		xml += fmt.Sprintf("<w:jc w:val=\"%s\" />", p.Properties.Alignment)
	}

	xml += "</w:pPr>"

	// 添加所有Run的XML
	for _, run := range p.Runs {
		xml += run.ToXML()
	}

	xml += "</w:p>"
	return xml
}
