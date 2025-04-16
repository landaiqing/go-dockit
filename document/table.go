package document

import (
	"fmt"
)

// Table 表示Word文档中的表格
type Table struct {
	Rows       []*TableRow
	Properties *TableProperties
}

// TableProperties 表示表格的属性
type TableProperties struct {
	Width       int    // 表格宽度，单位为twip
	WidthType   string // 宽度类型：auto, dxa, pct
	Alignment   string // 对齐方式：left, center, right
	Indent      int    // 缩进
	Borders     *TableBorders
	CellMargin  *TableCellMargin
	Layout      string // 布局方式：fixed, autofit
	Look        string // 表格外观
	Style       string // 表格样式ID
	FirstRow    bool   // 首行特殊格式
	LastRow     bool   // 末行特殊格式
	FirstColumn bool   // 首列特殊格式
	LastColumn  bool   // 末列特殊格式
	NoHBand     bool   // 无水平带状格式
	NoVBand     bool   // 无垂直带状格式
}

// TableBorders 表示表格的边框
type TableBorders struct {
	Top     *Border
	Bottom  *Border
	Left    *Border
	Right   *Border
	InsideH *Border
	InsideV *Border
}

// TableCellMargin 表示表格单元格的边距
type TableCellMargin struct {
	Top    int
	Bottom int
	Left   int
	Right  int
}

// TableRow 表示表格的行
type TableRow struct {
	Cells      []*TableCell
	Properties *TableRowProperties
}

// TableRowProperties 表示表格行的属性
type TableRowProperties struct {
	Height     int    // 行高，单位为twip
	HeightRule string // 行高规则：atLeast, exact, auto
	CantSplit  bool   // 不允许跨页分割
	IsHeader   bool   // 是否为表头行
}

// TableCell 表示表格的单元格
type TableCell struct {
	Content    []interface{} // 可以是段落、表格等元素
	Properties *TableCellProperties
}

// TableCellProperties 表示表格单元格的属性
type TableCellProperties struct {
	Width     int    // 单元格宽度，单位为twip
	WidthType string // 宽度类型：auto, dxa, pct
	VertAlign string // 垂直对齐方式：top, center, bottom
	Borders   *TableBorders
	Shading   *Shading
	GridSpan  int    // 跨列数
	VMerge    string // 垂直合并：restart, continue
	NoWrap    bool   // 不换行
	FitText   bool   // 适应文本
}

// NewTable 创建一个新的表格
func NewTable(rows, cols int) *Table {
	t := &Table{
		Rows: make([]*TableRow, 0),
		Properties: &TableProperties{
			Width:     0,
			WidthType: "auto",
			Alignment: "left",
			Layout:    "autofit",
			Borders: &TableBorders{
				Top:     &Border{Style: "single", Size: 4, Color: "000000", Space: 0},
				Bottom:  &Border{Style: "single", Size: 4, Color: "000000", Space: 0},
				Left:    &Border{Style: "single", Size: 4, Color: "000000", Space: 0},
				Right:   &Border{Style: "single", Size: 4, Color: "000000", Space: 0},
				InsideH: &Border{Style: "single", Size: 4, Color: "000000", Space: 0},
				InsideV: &Border{Style: "single", Size: 4, Color: "000000", Space: 0},
			},
			CellMargin: &TableCellMargin{
				Top:    0,
				Bottom: 0,
				Left:   108, // 约0.15厘米
				Right:  108,
			},
		},
	}

	// 创建行和单元格
	for i := 0; i < rows; i++ {
		row := t.AddRow()
		for j := 0; j < cols; j++ {
			row.AddCell()
		}
	}

	return t
}

// AddRow 向表格添加一行并返回它
func (t *Table) AddRow() *TableRow {
	r := &TableRow{
		Cells: make([]*TableCell, 0),
		Properties: &TableRowProperties{
			Height:     0,
			HeightRule: "auto",
			CantSplit:  false,
			IsHeader:   false,
		},
	}
	t.Rows = append(t.Rows, r)
	return r
}

// SetWidth 设置表格宽度
// width可以是整数(表示twip单位)或字符串(如"100%"表示百分比)
func (t *Table) SetWidth(width interface{}, widthType string) *Table {
	switch v := width.(type) {
	case int:
		t.Properties.Width = v
		t.Properties.WidthType = widthType
	case string:
		// 处理百分比格式，如"100%"
		if v == "100%" && widthType == "pct" {
			// Word中百分比是用整数表示的，5000 = 100%
			t.Properties.Width = 5000
			t.Properties.WidthType = "pct"
		} else {
			// 默认为自动宽度
			t.Properties.Width = 0
			t.Properties.WidthType = "auto"
		}
	default:
		// 默认为自动宽度
		t.Properties.Width = 0
		t.Properties.WidthType = "auto"
	}
	return t
}

// SetAlignment 设置表格对齐方式
func (t *Table) SetAlignment(alignment string) *Table {
	t.Properties.Alignment = alignment
	return t
}

// SetIndent 设置表格缩进
func (t *Table) SetIndent(indent int) *Table {
	t.Properties.Indent = indent
	return t
}

// SetLayout 设置表格布局方式
func (t *Table) SetLayout(layout string) *Table {
	t.Properties.Layout = layout
	return t
}

// SetBorders 设置表格边框
func (t *Table) SetBorders(position string, style string, size int, color string) *Table {
	// 如果style为空，设置为"none"
	if style == "" {
		style = "none"
	}

	// 如果color为空，设置为默认颜色黑色
	if color == "" {
		color = "000000"
	}

	border := &Border{
		Style: style,
		Size:  size,
		Color: color,
		Space: 0,
	}

	switch position {
	case "top":
		t.Properties.Borders.Top = border
	case "bottom":
		t.Properties.Borders.Bottom = border
	case "left":
		t.Properties.Borders.Left = border
	case "right":
		t.Properties.Borders.Right = border
	case "insideH":
		t.Properties.Borders.InsideH = border
	case "insideV":
		t.Properties.Borders.InsideV = border
	case "all":
		t.Properties.Borders.Top = border
		t.Properties.Borders.Bottom = border
		t.Properties.Borders.Left = border
		t.Properties.Borders.Right = border
		t.Properties.Borders.InsideH = border
		t.Properties.Borders.InsideV = border
	}

	return t
}

// SetCellMargin 设置表格单元格边距
func (t *Table) SetCellMargin(position string, margin int) *Table {
	switch position {
	case "top":
		t.Properties.CellMargin.Top = margin
	case "bottom":
		t.Properties.CellMargin.Bottom = margin
	case "left":
		t.Properties.CellMargin.Left = margin
	case "right":
		t.Properties.CellMargin.Right = margin
	case "all":
		t.Properties.CellMargin.Top = margin
		t.Properties.CellMargin.Bottom = margin
		t.Properties.CellMargin.Left = margin
		t.Properties.CellMargin.Right = margin
	}

	return t
}

// SetStyle 设置表格样式
func (t *Table) SetStyle(style string) *Table {
	t.Properties.Style = style
	return t
}

// SetLook 设置表格外观
func (t *Table) SetLook(firstRow, lastRow, firstColumn, lastColumn, noHBand, noVBand bool) *Table {
	t.Properties.FirstRow = firstRow
	t.Properties.LastRow = lastRow
	t.Properties.FirstColumn = firstColumn
	t.Properties.LastColumn = lastColumn
	t.Properties.NoHBand = noHBand
	t.Properties.NoVBand = noVBand

	// 计算Look值
	look := 0
	if firstRow {
		look |= 0x0020
	}
	if lastRow {
		look |= 0x0040
	}
	if firstColumn {
		look |= 0x0080
	}
	if lastColumn {
		look |= 0x0100
	}
	if noHBand {
		look |= 0x0200
	}
	if noVBand {
		look |= 0x0400
	}

	t.Properties.Look = fmt.Sprintf("%04X", look)

	return t
}

// AddCell 向表格行添加一个单元格并返回它
func (r *TableRow) AddCell() *TableCell {
	c := &TableCell{
		Content: make([]interface{}, 0),
		Properties: &TableCellProperties{
			Width:     0,
			WidthType: "auto",
			VertAlign: "top",
			GridSpan:  1,
		},
	}
	r.Cells = append(r.Cells, c)
	return c
}

// SetHeight 设置行高
func (r *TableRow) SetHeight(height int, rule string) *TableRow {
	r.Properties.Height = height
	r.Properties.HeightRule = rule
	return r
}

// SetCantSplit 设置不允许跨页分割
func (r *TableRow) SetCantSplit(cantSplit bool) *TableRow {
	r.Properties.CantSplit = cantSplit
	return r
}

// SetIsHeader 设置是否为表头行
func (r *TableRow) SetIsHeader(isHeader bool) *TableRow {
	r.Properties.IsHeader = isHeader
	return r
}

// AddParagraph 向单元格添加一个段落并返回它
func (c *TableCell) AddParagraph() *Paragraph {
	p := NewParagraph()
	c.Content = append(c.Content, p)
	return p
}

// AddTable 向单元格添加一个表格并返回它
func (c *TableCell) AddTable(rows, cols int) *Table {
	t := NewTable(rows, cols)
	c.Content = append(c.Content, t)
	return t
}

// SetWidth 设置单元格宽度
// width可以是整数(表示twip单位)或字符串(如"100%"表示百分比)
func (c *TableCell) SetWidth(width interface{}, widthType string) *TableCell {
	switch v := width.(type) {
	case int:
		c.Properties.Width = v
		c.Properties.WidthType = widthType
	case string:
		// 处理百分比格式，如"100%"
		if v == "100%" && widthType == "pct" {
			// Word中百分比是用整数表示的，5000 = 100%
			c.Properties.Width = 5000
			c.Properties.WidthType = "pct"
		} else {
			// 默认为自动宽度
			c.Properties.Width = 0
			c.Properties.WidthType = "auto"
		}
	default:
		// 默认为自动宽度
		c.Properties.Width = 0
		c.Properties.WidthType = "auto"
	}
	return c
}

// SetVertAlign 设置单元格垂直对齐方式
func (c *TableCell) SetVertAlign(vertAlign string) *TableCell {
	c.Properties.VertAlign = vertAlign
	return c
}

// SetBorders 设置单元格边框
func (c *TableCell) SetBorders(position string, style string, size int, color string) *TableCell {
	if c.Properties.Borders == nil {
		c.Properties.Borders = &TableBorders{}
	}

	// 如果style为空，设置为"none"
	if style == "" {
		style = "none"
	}

	// 如果color为空，设置为默认颜色黑色
	if color == "" {
		color = "000000"
	}

	border := &Border{
		Style: style,
		Size:  size,
		Color: color,
		Space: 0,
	}

	switch position {
	case "top":
		c.Properties.Borders.Top = border
	case "bottom":
		c.Properties.Borders.Bottom = border
	case "left":
		c.Properties.Borders.Left = border
	case "right":
		c.Properties.Borders.Right = border
	case "all":
		c.Properties.Borders.Top = border
		c.Properties.Borders.Bottom = border
		c.Properties.Borders.Left = border
		c.Properties.Borders.Right = border
	}

	return c
}

// SetShading 设置单元格底纹
func (c *TableCell) SetShading(fill, color, pattern string) *TableCell {
	c.Properties.Shading = &Shading{
		Fill:    fill,
		Color:   color,
		Pattern: pattern,
	}
	return c
}

// SetGridSpan 设置单元格跨列数
func (c *TableCell) SetGridSpan(gridSpan int) *TableCell {
	c.Properties.GridSpan = gridSpan
	return c
}

// SetVMerge 设置单元格垂直合并
func (c *TableCell) SetVMerge(vMerge string) *TableCell {
	c.Properties.VMerge = vMerge
	return c
}

// SetNoWrap 设置单元格不换行
func (c *TableCell) SetNoWrap(noWrap bool) *TableCell {
	c.Properties.NoWrap = noWrap
	return c
}

// SetFitText 设置单元格适应文本
func (c *TableCell) SetFitText(fitText bool) *TableCell {
	c.Properties.FitText = fitText
	return c
}

// ToXML 将表格转换为XML
func (t *Table) ToXML() string {
	xml := "<w:tbl>"

	// 添加表格属性
	xml += "<w:tblPr>"

	// 表格样式ID
	if t.Properties.Style != "" {
		xml += fmt.Sprintf("<w:tblStyle w:val=\"%s\" />", t.Properties.Style)
	}

	// 表格宽度
	xml += fmt.Sprintf("<w:tblW w:w=\"%d\" w:type=\"%s\" />", t.Properties.Width, t.Properties.WidthType)

	// 表格对齐方式
	if t.Properties.Alignment != "" {
		xml += fmt.Sprintf("<w:jc w:val=\"%s\" />", t.Properties.Alignment)
	}

	// 表格缩进
	if t.Properties.Indent > 0 {
		xml += fmt.Sprintf("<w:tblInd w:w=\"%d\" w:type=\"dxa\" />", t.Properties.Indent)
	}

	// 表格边框
	if t.Properties.Borders != nil {
		xml += "<w:tblBorders>"
		if t.Properties.Borders.Top != nil {
			xml += fmt.Sprintf("<w:top w:val=\"%s\" w:sz=\"%d\" w:space=\"%d\" w:color=\"%s\" />",
				t.Properties.Borders.Top.Style,
				t.Properties.Borders.Top.Size,
				t.Properties.Borders.Top.Space,
				t.Properties.Borders.Top.Color)
		}
		if t.Properties.Borders.Left != nil {
			xml += fmt.Sprintf("<w:left w:val=\"%s\" w:sz=\"%d\" w:space=\"%d\" w:color=\"%s\" />",
				t.Properties.Borders.Left.Style,
				t.Properties.Borders.Left.Size,
				t.Properties.Borders.Left.Space,
				t.Properties.Borders.Left.Color)
		}
		if t.Properties.Borders.Bottom != nil {
			xml += fmt.Sprintf("<w:bottom w:val=\"%s\" w:sz=\"%d\" w:space=\"%d\" w:color=\"%s\" />",
				t.Properties.Borders.Bottom.Style,
				t.Properties.Borders.Bottom.Size,
				t.Properties.Borders.Bottom.Space,
				t.Properties.Borders.Bottom.Color)
		}
		if t.Properties.Borders.Right != nil {
			xml += fmt.Sprintf("<w:right w:val=\"%s\" w:sz=\"%d\" w:space=\"%d\" w:color=\"%s\" />",
				t.Properties.Borders.Right.Style,
				t.Properties.Borders.Right.Size,
				t.Properties.Borders.Right.Space,
				t.Properties.Borders.Right.Color)
		}
		if t.Properties.Borders.InsideH != nil {
			xml += fmt.Sprintf("<w:insideH w:val=\"%s\" w:sz=\"%d\" w:space=\"%d\" w:color=\"%s\" />",
				t.Properties.Borders.InsideH.Style,
				t.Properties.Borders.InsideH.Size,
				t.Properties.Borders.InsideH.Space,
				t.Properties.Borders.InsideH.Color)
		}
		if t.Properties.Borders.InsideV != nil {
			xml += fmt.Sprintf("<w:insideV w:val=\"%s\" w:sz=\"%d\" w:space=\"%d\" w:color=\"%s\" />",
				t.Properties.Borders.InsideV.Style,
				t.Properties.Borders.InsideV.Size,
				t.Properties.Borders.InsideV.Space,
				t.Properties.Borders.InsideV.Color)
		}
		xml += "</w:tblBorders>"
	}

	// 表格布局
	if t.Properties.Layout != "" {
		xml += fmt.Sprintf("<w:tblLayout w:type=\"%s\" />", t.Properties.Layout)
	}

	// 单元格边距
	if t.Properties.CellMargin != nil {
		xml += "<w:tblCellMar>"
		if t.Properties.CellMargin.Top > 0 {
			xml += fmt.Sprintf("<w:top w:w=\"%d\" w:type=\"dxa\" />", t.Properties.CellMargin.Top)
		}
		if t.Properties.CellMargin.Left > 0 {
			xml += fmt.Sprintf("<w:left w:w=\"%d\" w:type=\"dxa\" />", t.Properties.CellMargin.Left)
		}
		if t.Properties.CellMargin.Bottom > 0 {
			xml += fmt.Sprintf("<w:bottom w:w=\"%d\" w:type=\"dxa\" />", t.Properties.CellMargin.Bottom)
		}
		if t.Properties.CellMargin.Right > 0 {
			xml += fmt.Sprintf("<w:right w:w=\"%d\" w:type=\"dxa\" />", t.Properties.CellMargin.Right)
		}
		xml += "</w:tblCellMar>"
	}

	// 表格外观
	if t.Properties.Look != "" {
		xml += fmt.Sprintf("<w:tblLook w:val=\"%s\" w:firstRow=\"%v\" w:lastRow=\"%v\" w:firstColumn=\"%v\" w:lastColumn=\"%v\" w:noHBand=\"%v\" w:noVBand=\"%v\" />",
			t.Properties.Look,
			formatBoolToWXml(t.Properties.FirstRow),
			formatBoolToWXml(t.Properties.LastRow),
			formatBoolToWXml(t.Properties.FirstColumn),
			formatBoolToWXml(t.Properties.LastColumn),
			formatBoolToWXml(t.Properties.NoHBand),
			formatBoolToWXml(t.Properties.NoVBand))
	}

	xml += "</w:tblPr>"

	// 添加表格网格
	xml += "<w:tblGrid>"
	if len(t.Rows) > 0 && len(t.Rows[0].Cells) > 0 {
		for i := 0; i < len(t.Rows[0].Cells); i++ {
			xml += "<w:gridCol />"
		}
	}
	xml += "</w:tblGrid>"

	// 添加所有行的XML
	for _, row := range t.Rows {
		xml += row.ToXML()
	}

	xml += "</w:tbl>"
	return xml
}

// formatBoolToWXml 将布尔值转换为Word XML中使用的字符串表示
func formatBoolToWXml(value bool) string {
	if value {
		return "1"
	}
	return "0"
}

// ToXML 将表格行转换为XML
func (r *TableRow) ToXML() string {
	xml := "<w:tr>"

	// 添加行属性
	xml += "<w:trPr>"

	// 行高
	if r.Properties.Height > 0 {
		xml += "<w:trHeight w:val=\"" + fmt.Sprintf("%d", r.Properties.Height) + "\" w:hRule=\"" + r.Properties.HeightRule + "\" />"
	}

	// 不允许跨页分割
	if r.Properties.CantSplit {
		xml += "<w:cantSplit />"
	}

	// 表头行
	if r.Properties.IsHeader {
		xml += "<w:tblHeader />"
	}

	xml += "</w:trPr>"

	// 添加所有单元格的XML
	for _, cell := range r.Cells {
		xml += cell.ToXML()
	}

	xml += "</w:tr>"
	return xml
}

// ToXML 将表格单元格转换为XML
func (c *TableCell) ToXML() string {
	xml := "<w:tc>"

	// 添加单元格属性
	xml += "<w:tcPr>"

	// 1. cnfStyle - 暂不实现

	// 2. 单元格宽度 (tcW)
	if c.Properties.Width > 0 {
		xml += "<w:tcW w:w=\"" + fmt.Sprintf("%d", c.Properties.Width) + "\" w:type=\"" + c.Properties.WidthType + "\" />"
	} else {
		xml += "<w:tcW w:w=\"0\" w:type=\"auto\" />"
	}

	// 3. 跨列数 (gridSpan)
	if c.Properties.GridSpan > 1 {
		xml += "<w:gridSpan w:val=\"" + fmt.Sprintf("%d", c.Properties.GridSpan) + "\" />"
	}

	// 4. hMerge - 暂不实现

	// 5. 垂直合并 (vMerge)
	if c.Properties.VMerge != "" {
		xml += "<w:vMerge w:val=\"" + c.Properties.VMerge + "\" />"
	}

	// 6. 单元格边框 (tcBorders)
	if c.Properties.Borders != nil {
		xml += "<w:tcBorders>"
		if c.Properties.Borders.Top != nil {
			xml += "<w:top w:val=\"" + c.Properties.Borders.Top.Style + "\" w:sz=\"" + fmt.Sprintf("%d", c.Properties.Borders.Top.Size) + "\" w:space=\"" + fmt.Sprintf("%d", c.Properties.Borders.Top.Space) + "\" w:color=\"" + c.Properties.Borders.Top.Color + "\" />"
		}
		if c.Properties.Borders.Bottom != nil {
			xml += "<w:bottom w:val=\"" + c.Properties.Borders.Bottom.Style + "\" w:sz=\"" + fmt.Sprintf("%d", c.Properties.Borders.Bottom.Size) + "\" w:space=\"" + fmt.Sprintf("%d", c.Properties.Borders.Bottom.Space) + "\" w:color=\"" + c.Properties.Borders.Bottom.Color + "\" />"
		}
		if c.Properties.Borders.Left != nil {
			xml += "<w:left w:val=\"" + c.Properties.Borders.Left.Style + "\" w:sz=\"" + fmt.Sprintf("%d", c.Properties.Borders.Left.Size) + "\" w:space=\"" + fmt.Sprintf("%d", c.Properties.Borders.Left.Space) + "\" w:color=\"" + c.Properties.Borders.Left.Color + "\" />"
		}
		if c.Properties.Borders.Right != nil {
			xml += "<w:right w:val=\"" + c.Properties.Borders.Right.Style + "\" w:sz=\"" + fmt.Sprintf("%d", c.Properties.Borders.Right.Size) + "\" w:space=\"" + fmt.Sprintf("%d", c.Properties.Borders.Right.Space) + "\" w:color=\"" + c.Properties.Borders.Right.Color + "\" />"
		}
		xml += "</w:tcBorders>"
	}

	// 7. 底纹 (shd)
	if c.Properties.Shading != nil {
		xml += "<w:shd w:val=\"" + c.Properties.Shading.Pattern + "\" w:fill=\"" + c.Properties.Shading.Fill + "\" w:color=\"" + c.Properties.Shading.Color + "\" />"
	}

	// 8. 不换行 (noWrap)
	if c.Properties.NoWrap {
		xml += "<w:noWrap />"
	}

	// 9. tcMar - 暂不实现

	// 10. textDirection - 暂不实现

	// 11. 适应文本 (tcFitText)
	if c.Properties.FitText {
		xml += "<w:tcFitText />"
	}

	// 12. 垂直对齐方式 (vAlign)
	if c.Properties.VertAlign != "" {
		xml += "<w:vAlign w:val=\"" + c.Properties.VertAlign + "\" />"
	}

	// 13. hideMark - 暂不实现

	xml += "</w:tcPr>"

	// 添加所有内容元素的XML
	for _, content := range c.Content {
		switch v := content.(type) {
		case *Paragraph:
			xml += v.ToXML()
		case *Table:
			xml += v.ToXML()
		}
	}

	// 如果单元格没有内容，添加一个空段落
	if len(c.Content) == 0 {
		xml += "<w:p><w:pPr></w:pPr></w:p>"
	}

	xml += "</w:tc>"
	return xml
}
