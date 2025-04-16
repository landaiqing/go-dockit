package document

import "fmt"

// Body 表示Word文档的主体部分
type Body struct {
	Content           []interface{} // 可以是段落、表格等元素
	SectionProperties *SectionProperties
}

// SectionProperties 表示节属性
type SectionProperties struct {
	PageSize        *PageSize
	PageMargin      *PageMargin
	Columns         *Columns
	DocGrid         *DocGrid
	HeaderReference []*HeaderFooterReference
	FooterReference []*HeaderFooterReference
}

// PageSize 表示页面大小
type PageSize struct {
	Width       int    // 页面宽度，单位为twip
	Height      int    // 页面高度，单位为twip
	Orientation string // 页面方向：portrait, landscape
}

// PageMargin 表示页面边距
type PageMargin struct {
	Top    int // 上边距，单位为twip
	Right  int // 右边距，单位为twip
	Bottom int // 下边距，单位为twip
	Left   int // 左边距，单位为twip
	Header int // 页眉边距，单位为twip
	Footer int // 页脚边距，单位为twip
	Gutter int // 装订线，单位为twip
}

// Columns 表示分栏
type Columns struct {
	Num   int // 栏数
	Space int // 栏间距，单位为twip
}

// DocGrid 表示文档网格
type DocGrid struct {
	LinePitch int // 行距，单位为twip
}

// HeaderFooterReference 表示页眉页脚引用
type HeaderFooterReference struct {
	Type string // 类型：default, first, even
	ID   string // 引用ID
}

// NewBody 创建一个新的文档主体
func NewBody() *Body {
	return &Body{
		Content: make([]interface{}, 0),
		SectionProperties: &SectionProperties{
			PageSize: &PageSize{
				Width:       12240, // 8.5英寸 = 12240 twip
				Height:      15840, // 11英寸 = 15840 twip
				Orientation: "portrait",
			},
			PageMargin: &PageMargin{
				Top:    1440, // 1英寸 = 1440 twip
				Right:  1440,
				Bottom: 1440,
				Left:   1440,
				Header: 720,
				Footer: 720,
				Gutter: 0,
			},
			Columns: &Columns{
				Num:   1,
				Space: 720,
			},
			DocGrid: &DocGrid{
				LinePitch: 360,
			},
			HeaderReference: make([]*HeaderFooterReference, 0),
			FooterReference: make([]*HeaderFooterReference, 0),
		},
	}
}

// AddParagraph 向文档主体添加一个段落并返回它
func (b *Body) AddParagraph() *Paragraph {
	p := NewParagraph()
	b.Content = append(b.Content, p)
	return p
}

// AddTable 向文档主体添加一个表格并返回它
func (b *Body) AddTable(rows, cols int) *Table {
	t := NewTable(rows, cols)
	b.Content = append(b.Content, t)
	return t
}

// AddPageBreak 向文档主体添加一个分页符
func (b *Body) AddPageBreak() *Paragraph {
	p := NewParagraph()
	p.AddRun().AddBreak(BreakTypePage)
	b.Content = append(b.Content, p)
	return p
}

// AddSectionBreak 向文档主体添加一个分节符
func (b *Body) AddSectionBreak() *Paragraph {
	p := NewParagraph()
	p.AddRun().AddBreak(BreakTypeSection)
	b.Content = append(b.Content, p)
	return p
}

// ToXML 将Body转换为XML
func (b *Body) ToXML() string {
	xml := "<w:body>"

	// 添加所有内容元素的XML
	for _, content := range b.Content {
		switch v := content.(type) {
		case *Paragraph:
			xml += v.ToXML()
		case *Table:
			xml += v.ToXML()
		}
	}

	// 添加节属性
	xml += "<w:sectPr xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"

	// 页眉引用
	for _, headerRef := range b.SectionProperties.HeaderReference {
		xml += fmt.Sprintf("<w:headerReference w:type=\"%s\" r:id=\"%s\" />",
			headerRef.Type, headerRef.ID)
	}

	// 页脚引用
	for _, footerRef := range b.SectionProperties.FooterReference {
		xml += fmt.Sprintf("<w:footerReference w:type=\"%s\" r:id=\"%s\" />",
			footerRef.Type, footerRef.ID)
	}

	// 页面大小
	if b.SectionProperties.PageSize != nil {
		xml += fmt.Sprintf("<w:pgSz w:w=\"%d\" w:h=\"%d\" w:orient=\"%s\" />",
			b.SectionProperties.PageSize.Width,
			b.SectionProperties.PageSize.Height,
			b.SectionProperties.PageSize.Orientation)
	}

	// 页面边距
	if b.SectionProperties.PageMargin != nil {
		xml += fmt.Sprintf("<w:pgMar w:top=\"%d\" w:right=\"%d\" w:bottom=\"%d\" w:left=\"%d\" w:header=\"%d\" w:footer=\"%d\" w:gutter=\"%d\" />",
			b.SectionProperties.PageMargin.Top,
			b.SectionProperties.PageMargin.Right,
			b.SectionProperties.PageMargin.Bottom,
			b.SectionProperties.PageMargin.Left,
			b.SectionProperties.PageMargin.Header,
			b.SectionProperties.PageMargin.Footer,
			b.SectionProperties.PageMargin.Gutter)
	}

	// 分栏
	if b.SectionProperties.Columns != nil {
		xml += fmt.Sprintf("<w:cols w:num=\"%d\" w:space=\"%d\" />",
			b.SectionProperties.Columns.Num,
			b.SectionProperties.Columns.Space)
	}

	// 文档网格
	if b.SectionProperties.DocGrid != nil {
		xml += fmt.Sprintf("<w:docGrid w:linePitch=\"%d\" />",
			b.SectionProperties.DocGrid.LinePitch)
	}

	xml += "</w:sectPr>"

	xml += "</w:body>"
	return xml
}
