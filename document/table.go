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
func (t *Table) SetWidth(width int, widthType string) *Table {
	t.Properties.Width = width
	t.Properties.WidthType = widthType
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
func (c *TableCell) SetWidth(width int, widthType string) *TableCell {
	c.Properties.Width = width
	c.Properties.WidthType = widthType
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

	// 表格宽度
	if t.Properties.Width > 0 {
		xml += "<w:tblW w:w=\"" + fmt.Sprintf("%d", t.Properties.Width) + "\""
		xml += " w:type=\"" + t.Properties.WidthType + "\" />"
	} else {
		xml += "<w:tblW w:w=\"0\" w:type=\"auto\" />"
	}

	// 表格对齐方式
	if t.Properties.Alignment != "" {
		xml += "<w:jc w:val=\"" + t.Properties.Alignment + "\" />"
	}

	// 表格缩进
	if t.Properties.Indent > 0 {
		xml += "<w:tblInd w:w=\"" + fmt.Sprintf("%d", t.Properties.Indent) + "\""
		xml += " w:type=\"dxa\" />"
	}

	// 表格边框
	if t.Properties.Borders != nil {
		xml += "<w:tblBorders>"
		if t.Properties.Borders.Top != nil {
			xml += "<w:top w:val=\"" + t.Properties.Borders.Top.Style + "\""
			xml += " w:sz=\"" + fmt.Sprintf("%d", t.Properties.Borders.Top.Size) + "\""
			xml += " w:space=\"" + fmt.Sprintf("%d", t.Properties.Borders.Top.Space) + "\""
			xml += " w:color=\"" + t.Properties.Borders.Top.Color + "\" />"
		}
		if t.Properties.Borders.Bottom != nil {
			xml += "<w:bottom w:val=\"" + t.Properties.Borders.Bottom.Style + "\""
			xml += " w:sz=\"" + fmt.Sprintf("%d", t.Properties.Borders.Bottom.Size) + "\""
			xml += " w:space=\"" + fmt.Sprintf("%d", t.Properties.Borders.Bottom.Space) + "\""
			xml += " w:color=\"" + t.Properties.Borders.Bottom.Color + "\" />"
		}
		if t.Properties.Borders.Left != nil {
			xml += "<w:left w:val=\"" + t.Properties.Borders.Left.Style + "\""
			xml += " w:sz=\"" + fmt.Sprintf("%d", t.Properties.Borders.Left.Size) + "\""
			xml += " w:space=\"" + fmt.Sprintf("%d", t.Properties.Borders.Left.Space) + "\""
			xml += " w:color=\"" + t.Properties.Borders.Left.Color + "\" />"
		}
		if t.Properties.Borders.Right != nil {
			xml += "<w:right w:val=\"" + t.Properties.Borders.Right.Style + "\""
			xml += " w:sz=\"" + fmt.Sprintf("%d", t.Properties.Borders.Right.Size) + "\""
			xml += " w:space=\"" + fmt.Sprintf("%d", t.Properties.Borders.Right.Space) + "\""
			xml += " w:color=\"" + t.Properties.Borders.Right.Color + "\" />"
		}
		if t.Properties.Borders.InsideH != nil {
			xml += "<w:insideH w:val=\"" + t.Properties.Borders.InsideH.Style + "\""
			xml += " w:sz=\"" + fmt.Sprintf("%d", t.Properties.Borders.InsideH.Size) + "\""
			xml += " w:space=\"" + fmt.Sprintf("%d", t.Properties.Borders.InsideH.Space) + "\""
			xml += " w:color=\"" + t.Properties.Borders.InsideH.Color + "\" />"
		}
		if t.Properties.Borders.InsideV != nil {
			xml += "<w:insideV w:val=\"" + t.Properties.Borders.InsideV.Style + "\""
			xml += " w:sz=\"" + fmt.Sprintf("%d", t.Properties.Borders.InsideV.Size) + "\""
			xml += " w:space=\"" + fmt.Sprintf("%d", t.Properties.Borders.InsideV.Space) + "\""
			xml += " w:color=\"" + t.Properties.Borders.InsideV.Color + "\" />"
		}
		xml += "</w:tblBorders>"
	}

	// 表格单元格边距
	if t.Properties.CellMargin != nil {
		xml += "<w:tblCellMar>"
		if t.Properties.CellMargin.Top > 0 {
			xml += "<w:top w:w=\"" + fmt.Sprintf("%d", t.Properties.CellMargin.Top) + "\""
			xml += " w:type=\"dxa\" />"
		}
		if t.Properties.CellMargin.Bottom > 0 {
			xml += "<w:bottom w:w=\"" + fmt.Sprintf("%d", t.Properties.CellMargin.Bottom) + "\""
			xml += " w:type=\"dxa\" />"
		}
		if t.Properties.CellMargin.Left > 0 {
			xml += "<w:left w:w=\"" + fmt.Sprintf("%d", t.Properties.CellMargin.Left) + "\""
			xml += " w:type=\"dxa\" />"
		}
		if t.Properties.CellMargin.Right > 0 {
			xml += "<w:right w:w=\"" + fmt.Sprintf("%d", t.Properties.CellMargin.Right) + "\""
			xml += " w:type=\"dxa\" />"
		}
		xml += "</w:tblCellMar>"
	}

	// 表格布局方式
	if t.Properties.Layout != "" {
		xml += "<w:tblLayout w:type=\"" + t.Properties.Layout + "\" />"
	}

	// 表格样式
	if t.Properties.Style != "" {
		xml += "<w:tblStyle w:val=\"" + t.Properties.Style + "\" />"
	}

	// 表格外观
	if t.Properties.Look != "" {
		xml += "<w:tblLook w:val=\"" + t.Properties.Look + "\""
		xml += " w:firstRow=\"" + fmt.Sprintf("%d", boolToInt(t.Properties.FirstRow)) + "\""
		xml += " w:lastRow=\"" + fmt.Sprintf("%d", boolToInt(t.Properties.LastRow)) + "\""
		xml += " w:firstColumn=\"" + fmt.Sprintf("%d", boolToInt(t.Properties.FirstColumn)) + "\""
		xml += " w:lastColumn=\"" + fmt.Sprintf("%d", boolToInt(t.Properties.LastColumn)) + "\""
		xml += " w:noHBand=\"" + fmt.Sprintf("%d", boolToInt(t.Properties.NoHBand)) + "\""
		xml += " w:noVBand=\"" + fmt.Sprintf("%d", boolToInt(t.Properties.NoVBand)) + "\" />"
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
		xml += "<w:tr>"

		// 添加行属性
		xml += "<w:trPr>"

		// 行高
		if row.Properties.Height > 0 {
			xml += "<w:trHeight w:val=\"" + fmt.Sprintf("%d", row.Properties.Height) + "\""
			xml += " w:hRule=\"" + row.Properties.HeightRule + "\" />"
		}

		// 不允许跨页分割
		if row.Properties.CantSplit {
			xml += "<w:cantSplit />"
		}

		// 表头行
		if row.Properties.IsHeader {
			xml += "<w:tblHeader />"
		}

		xml += "</w:trPr>"

		// 添加所有单元格的XML
		for _, cell := range row.Cells {
			xml += "<w:tc>"

			// 添加单元格属性
			xml += "<w:tcPr>"

			// 单元格宽度
			if cell.Properties.Width > 0 {
				xml += "<w:tcW w:w=\"" + fmt.Sprintf("%d", cell.Properties.Width) + "\""
				xml += " w:type=\"" + cell.Properties.WidthType + "\" />"
			} else {
				xml += "<w:tcW w:w=\"0\" w:type=\"auto\" />"
			}

			// 垂直对齐方式
			if cell.Properties.VertAlign != "" {
				xml += "<w:vAlign w:val=\"" + cell.Properties.VertAlign + "\" />"
			}

			// 单元格边框
			if cell.Properties.Borders != nil {
				xml += "<w:tcBorders>"
				if cell.Properties.Borders.Top != nil {
					xml += "<w:top w:val=\"" + cell.Properties.Borders.Top.Style + "\""
					xml += " w:sz=\"" + fmt.Sprintf("%d", cell.Properties.Borders.Top.Size) + "\""
					xml += " w:space=\"" + fmt.Sprintf("%d", cell.Properties.Borders.Top.Space) + "\""
					xml += " w:color=\"" + cell.Properties.Borders.Top.Color + "\" />"
				}
				if cell.Properties.Borders.Bottom != nil {
					xml += "<w:bottom w:val=\"" + cell.Properties.Borders.Bottom.Style + "\""
					xml += " w:sz=\"" + fmt.Sprintf("%d", cell.Properties.Borders.Bottom.Size) + "\""
					xml += " w:space=\"" + fmt.Sprintf("%d", cell.Properties.Borders.Bottom.Space) + "\""
					xml += " w:color=\"" + cell.Properties.Borders.Bottom.Color + "\" />"
				}
				if cell.Properties.Borders.Left != nil {
					xml += "<w:left w:val=\"" + cell.Properties.Borders.Left.Style + "\""
					xml += " w:sz=\"" + fmt.Sprintf("%d", cell.Properties.Borders.Left.Size) + "\""
					xml += " w:space=\"" + fmt.Sprintf("%d", cell.Properties.Borders.Left.Space) + "\""
					xml += " w:color=\"" + cell.Properties.Borders.Left.Color + "\" />"
				}
				if cell.Properties.Borders.Right != nil {
					xml += "<w:right w:val=\"" + cell.Properties.Borders.Right.Style + "\""
					xml += " w:sz=\"" + fmt.Sprintf("%d", cell.Properties.Borders.Right.Size) + "\""
					xml += " w:space=\"" + fmt.Sprintf("%d", cell.Properties.Borders.Right.Space) + "\""
					xml += " w:color=\"" + cell.Properties.Borders.Right.Color + "\" />"
				}
				xml += "</w:tcBorders>"
			}

			// 底纹
			if cell.Properties.Shading != nil {
				xml += "<w:shd w:val=\"" + cell.Properties.Shading.Pattern + "\""
				xml += " w:fill=\"" + cell.Properties.Shading.Fill + "\""
				xml += " w:color=\"" + cell.Properties.Shading.Color + "\" />"
			}

			// 跨列数
			if cell.Properties.GridSpan > 1 {
				xml += "<w:gridSpan w:val=\"" + fmt.Sprintf("%d", cell.Properties.GridSpan) + "\" />"
			}

			// 垂直合并
			if cell.Properties.VMerge != "" {
				xml += "<w:vMerge w:val=\"" + cell.Properties.VMerge + "\" />"
			}

			// 不换行
			if cell.Properties.NoWrap {
				xml += "<w:noWrap />"
			}

			// 适应文本
			if cell.Properties.FitText {
				xml += "<w:fitText />"
			}

			xml += "</w:tcPr>"

			// 添加所有内容元素的XML
			for _, content := range cell.Content {
				switch v := content.(type) {
				case *Paragraph:
					xml += v.ToXML()
				case *Table:
					xml += v.ToXML()
				}
			}

			// 如果单元格没有内容，添加一个空段落
			if len(cell.Content) == 0 {
				xml += "<w:p><w:pPr></w:pPr></w:p>"
			}

			xml += "</w:tc>"
		}

		xml += "</w:tr>"
	}

	xml += "</w:tbl>"
	return xml
}
