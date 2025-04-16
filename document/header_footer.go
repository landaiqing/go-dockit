package document

// Header 表示Word文档中的页眉
type Header struct {
	ID      string
	Content []interface{} // 可以是段落、表格等元素
}

// Footer 表示Word文档中的页脚
type Footer struct {
	ID      string
	Content []interface{} // 可以是段落、表格等元素
}

// NewHeader 创建一个新的页眉
func NewHeader() *Header {
	return &Header{
		ID:      generateUniqueID(),
		Content: make([]interface{}, 0),
	}
}

// NewFooter 创建一个新的页脚
func NewFooter() *Footer {
	return &Footer{
		ID:      generateUniqueID(),
		Content: make([]interface{}, 0),
	}
}

// AddParagraph 向页眉添加一个段落并返回它
func (h *Header) AddParagraph() *Paragraph {
	p := NewParagraph()
	h.Content = append(h.Content, p)
	return p
}

// AddTable 向页眉添加一个表格并返回它
func (h *Header) AddTable(rows, cols int) *Table {
	t := NewTable(rows, cols)
	h.Content = append(h.Content, t)
	return t
}

// AddParagraph 向页脚添加一个段落并返回它
func (f *Footer) AddParagraph() *Paragraph {
	p := NewParagraph()
	f.Content = append(f.Content, p)
	return p
}

// AddTable 向页脚添加一个表格并返回它
func (f *Footer) AddTable(rows, cols int) *Table {
	t := NewTable(rows, cols)
	f.Content = append(f.Content, t)
	return t
}

// AddPageNumber 添加一个居中的页码
func (f *Footer) AddPageNumber() *Paragraph {
	// 创建一个新段落
	para := f.AddParagraph()
	para.SetAlignment("center")

	// 添加"第"文本
	para.AddRun().AddText("第 ")

	// 创建页码域的所有部分
	// 1. 域开始
	fieldBegin := para.AddRun()
	fieldBegin.Field = &Field{
		Type: "begin",
		Code: "PAGE",
	}

	// 2. 域分隔符
	fieldSeparate := para.AddRun()
	fieldSeparate.Field = &Field{
		Type: "separate",
	}

	// 3. 页码内容（在Word中会替换为实际页码）
	fieldContent := para.AddRun()
	fieldContent.AddText("1")

	// 4. 域结束
	fieldEnd := para.AddRun()
	fieldEnd.Field = &Field{
		Type: "end",
	}

	// 添加"页"文本
	para.AddRun().AddText(" 页")

	return para
}

// ToXML 将页眉转换为XML
func (h *Header) ToXML() string {
	xml := "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
	xml += "<w:hdr xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"

	// 添加所有内容元素的XML
	for _, content := range h.Content {
		switch v := content.(type) {
		case *Paragraph:
			xml += v.ToXML()
		case *Table:
			xml += v.ToXML()
		}
	}

	xml += "</w:hdr>"
	return xml
}

// ToXML 将页脚转换为XML
func (f *Footer) ToXML() string {
	xml := "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
	xml += "<w:ftr xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"

	// 添加所有内容元素的XML
	for _, content := range f.Content {
		switch v := content.(type) {
		case *Paragraph:
			xml += v.ToXML()
		case *Table:
			xml += v.ToXML()
		}
	}

	xml += "</w:ftr>"
	return xml
}
