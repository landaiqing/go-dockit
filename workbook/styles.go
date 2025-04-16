package workbook

import (
	"fmt"
	"strings"
)

// Styles 表示Excel文档中的样式集合
type Styles struct {
	Fonts         []*Font
	Fills         []*Fill
	Borders       []*Border
	NumberFormats []*NumberFormat
	CellStyles    []*CellStyleDef
	CellStyleXfs  []*CellStyleXf
	CellXfs       []*CellXf
}

// NewStyles 创建一个新的样式集合
func NewStyles() *Styles {
	s := &Styles{
		Fonts:         make([]*Font, 0),
		Fills:         make([]*Fill, 0),
		Borders:       make([]*Border, 0),
		NumberFormats: make([]*NumberFormat, 0),
		CellStyles:    make([]*CellStyleDef, 0),
		CellStyleXfs:  make([]*CellStyleXf, 0),
		CellXfs:       make([]*CellXf, 0),
	}

	// 添加默认字体
	s.AddFont("Calibri", 11, false, false, false, "")

	// 添加默认填充
	s.AddFill("none", "")
	s.AddFill("gray125", "")

	// 添加默认边框
	s.AddBorder()

	// 添加默认数字格式
	s.AddNumberFormat(0, "General")

	// 添加默认单元格样式XF
	s.AddCellStyleXf(0, 0, 0, 0, nil)

	// 添加默认单元格样式
	// 第二个参数是XF ID，指向已创建的CellStyleXfs中的索引
	s.AddCellStyle("Normal", 0, 0)

	// 添加默认单元格XF
	s.AddCellXf(0, 0, 0, 0, nil)

	return s
}

// Font 表示字体
type Font struct {
	Name      string
	Size      float64
	Bold      bool
	Italic    bool
	Underline bool
	Color     string
}

// AddFont 添加字体
func (s *Styles) AddFont(name string, size float64, bold, italic, underline bool, color string) *Font {
	font := &Font{
		Name:      name,
		Size:      size,
		Bold:      bold,
		Italic:    italic,
		Underline: underline,
		Color:     color,
	}
	s.Fonts = append(s.Fonts, font)
	return font
}

// Fill 表示填充
type Fill struct {
	PatternType string
	FgColor     string
	BgColor     string
}

// AddFill 添加填充
func (s *Styles) AddFill(patternType, fgColor string) *Fill {
	fill := &Fill{
		PatternType: patternType,
		FgColor:     fgColor,
	}
	s.Fills = append(s.Fills, fill)
	return fill
}

// Border 表示边框
type Border struct {
	Left   *BorderStyle
	Right  *BorderStyle
	Top    *BorderStyle
	Bottom *BorderStyle
}

// BorderStyle 表示边框样式
type BorderStyle struct {
	Style string
	Color string
}

// AddBorder 添加边框
func (s *Styles) AddBorder() *Border {
	border := &Border{
		Left:   &BorderStyle{},
		Right:  &BorderStyle{},
		Top:    &BorderStyle{},
		Bottom: &BorderStyle{},
	}
	s.Borders = append(s.Borders, border)
	return border
}

// NumberFormat 表示数字格式
type NumberFormat struct {
	ID   int
	Code string
}

// AddNumberFormat 添加数字格式
func (s *Styles) AddNumberFormat(id int, code string) *NumberFormat {
	nf := &NumberFormat{
		ID:   id,
		Code: code,
	}
	s.NumberFormats = append(s.NumberFormats, nf)
	return nf
}

// CellStyleDef 表示单元格样式定义
type CellStyleDef struct {
	Name          string
	XfId          int
	BuiltinId     int
	CustomBuiltin bool
}

// AddCellStyle 添加单元格样式
func (s *Styles) AddCellStyle(name string, xfId int, builtinId int) *CellStyleDef {
	cs := &CellStyleDef{
		Name:      name,
		XfId:      xfId,
		BuiltinId: builtinId,
	}
	s.CellStyles = append(s.CellStyles, cs)
	return cs
}

// CellStyleXf 表示单元格样式XF
type CellStyleXf struct {
	FontId            int
	FillId            int
	BorderId          int
	NumFmtId          int
	Alignment         *Alignment
	ApplyFont         bool
	ApplyFill         bool
	ApplyBorder       bool
	ApplyNumberFormat bool
	ApplyAlignment    bool
}

// AddCellStyleXf 添加单元格样式XF
func (s *Styles) AddCellStyleXf(fontId, fillId, borderId, numFmtId int, alignment *Alignment) *CellStyleXf {
	csx := &CellStyleXf{
		FontId:            fontId,
		FillId:            fillId,
		BorderId:          borderId,
		NumFmtId:          numFmtId,
		Alignment:         alignment,
		ApplyFont:         fontId > 0,
		ApplyFill:         fillId > 0,
		ApplyBorder:       borderId > 0,
		ApplyNumberFormat: numFmtId > 0,
		ApplyAlignment:    alignment != nil,
	}
	s.CellStyleXfs = append(s.CellStyleXfs, csx)
	return csx
}

// CellXf 表示单元格XF
type CellXf struct {
	FontId            int
	FillId            int
	BorderId          int
	NumFmtId          int
	Alignment         *Alignment
	ApplyFont         bool
	ApplyFill         bool
	ApplyBorder       bool
	ApplyNumberFormat bool
	ApplyAlignment    bool
}

// AddCellXf 添加单元格XF
func (s *Styles) AddCellXf(fontId, fillId, borderId, numFmtId int, alignment *Alignment) int {
	cx := &CellXf{
		FontId:            fontId,
		FillId:            fillId,
		BorderId:          borderId,
		NumFmtId:          numFmtId,
		Alignment:         alignment,
		ApplyFont:         fontId > 0,
		ApplyFill:         fillId > 0,
		ApplyBorder:       borderId > 0,
		ApplyNumberFormat: numFmtId > 0,
		ApplyAlignment:    alignment != nil,
	}
	s.CellXfs = append(s.CellXfs, cx)
	return len(s.CellXfs) - 1
}

// CreateStyle 创建一个完整的单元格样式并返回样式ID
func (s *Styles) CreateStyle(fontName string, fontSize float64, bold, italic, underline bool, fontColor string,
	fillPattern, fillColor string, borderStyle string, borderColor string, numFmtCode string,
	hAlign, vAlign string, wrapText bool) int {

	// 添加字体
	fontId := 0
	if fontName != "" || fontSize > 0 || bold || italic || underline || fontColor != "" {
		s.AddFont(fontName, fontSize, bold, italic, underline, fontColor)
		fontId = len(s.Fonts) - 1
	}

	// 添加填充
	fillId := 0
	if fillPattern != "" {
		s.AddFill(fillPattern, fillColor)
		fillId = len(s.Fills) - 1
	}

	// 添加边框
	borderId := 0
	if borderStyle != "" {
		border := s.AddBorder()
		border.Left.Style = borderStyle
		border.Left.Color = borderColor
		border.Right.Style = borderStyle
		border.Right.Color = borderColor
		border.Top.Style = borderStyle
		border.Top.Color = borderColor
		border.Bottom.Style = borderStyle
		border.Bottom.Color = borderColor
		borderId = len(s.Borders) - 1
	}

	// 添加数字格式 - 改进
	numFmtId := 0
	if numFmtCode != "" {
		// 处理货币格式中的引号问题
		if strings.Contains(numFmtCode, "\"¥\"") {
			// 人民币格式使用编号7
			numFmtId = 7
		} else if numFmtCode == "0.00E+00" || numFmtCode == "##0.0E+0" {
			// 使用Excel内置的科学计数格式ID
			numFmtId = 11
		} else {
			// 检查是否是其他内置格式
			builtinId := getBuiltinNumberFormatId(numFmtCode)
			if builtinId > 0 {
				numFmtId = builtinId
			} else {
				// 如果不是内置格式，创建自定义格式
				// Excel自定义格式从164开始
				numFmtId = 164 + len(s.NumberFormats)
				s.AddNumberFormat(numFmtId, numFmtCode)
			}
		}
	}

	var alignment *Alignment
	if hAlign != "" || vAlign != "" || wrapText {
		alignment = &Alignment{
			Horizontal: hAlign,
			Vertical:   vAlign,
			WrapText:   wrapText,
		}
	}

	// 创建单元格XF并返回其索引
	styleIndex := s.AddCellXf(fontId, fillId, borderId, numFmtId, alignment)
	return styleIndex
}

// 获取内置数字格式ID
func getBuiltinNumberFormatId(format string) int {
	builtinFormats := map[string]int{
		// 常规格式
		"General": 0,

		// 数值格式
		"0":        1, // 整数
		"0.00":     2, // 小数
		"#,##0":    3, // 千位分隔的整数
		"#,##0.00": 4, // 千位分隔的小数

		// 货币格式
		"¥#,##0;¥\\-#,##0":        7,  // 人民币格式
		"¥#,##0;[Red]¥\\-#,##0":   8,  // 人民币格式(负数为红色)
		"\"¥\"#,##0.00":           7,  // 人民币格式带小数 - 这个格式有问题
		"$#,##0.00":               44, // 美元格式
		"$#,##0.00_);($#,##0.00)": 43, // 会计专用美元格式
		"_(\"$\"* #,##0.00_)":     42, // 会计专用美元格式(负数带括号)
		"_-* #,##0.00_-":          4,  // 会计专用格式

		// 百分比格式
		"0%":    9,  // 整数百分比
		"0.00%": 10, // 小数百分比

		// 科学计数格式
		"0.00E+00": 11, // 科学计数 - 确保正确提供此格式
		"##0.0E+0": 11, // 科学计数 - 另一种写法

		// 分数格式
		"# ?/?":   12, // 分数 (例如:1/4)
		"# ??/??": 13, // 分数 (例如:3/16)

		// 日期格式 - 增加更多常见日期格式
		"mm-dd-yy":                       14, // 月-日-年
		"d-mmm-yy":                       15, // 日-月缩写-年
		"d-mmm":                          16, // 日-月缩写
		"mmm-yy":                         17, // 月缩写-年
		"mm/dd/yy":                       30, // 月/日/年
		"mm/dd/yyyy":                     22, // 月/日/完整年份
		"yyyy/mm/dd":                     20, // ISO 8601 格式
		"dd/mm/yyyy":                     21, // 欧洲日期格式
		"yyyy-mm-dd":                     22, // ISO 日期格式
		"yyyy-mm-dd;@":                   22, // ISO 日期格式(带文本)
		"yyyy/mm/dd;@":                   22, // ISO 日期格式(带文本)
		"[$-804]yyyy\"年\"mm\"月\"dd\"日\"": 31, // 中文日期格式

		// 时间格式
		"h:mm AM/PM":    18, // 12小时制时间
		"h:mm:ss AM/PM": 19, // 12小时制时间(带秒)
		"h:mm":          20, // 24小时制时间
		"h:mm:ss":       21, // 24小时制时间(带秒)
		"m/d/yy h:mm":   22, // 日期和时间
		"mm:ss":         45, // 分:秒
		"[h]:mm:ss":     46, // 超过24小时的时间
		"mmss.0":        47, // 分秒.毫秒

		// 其他格式
		"#,##0 ;(#,##0)":           37, // 带括号的负数
		"#,##0 ;[Red](#,##0)":      38, // 红色负数
		"#,##0.00;(#,##0.00)":      39, // 带括号的负小数
		"#,##0.00;[Red](#,##0.00)": 40, // 红色负小数
		"@":                        49, // 文本
	}

	if id, ok := builtinFormats[format]; ok {
		return id
	}
	return 0
}

// ToXML 将样式转换为XML
func (s *Styles) ToXML() string {
	xml := "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
	xml += "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n"

	// 数字格式
	if len(s.NumberFormats) > 0 {
		xml += fmt.Sprintf("  <numFmts count=\"%d\">\n", len(s.NumberFormats))
		for _, nf := range s.NumberFormats {
			// 处理格式代码中的引号，将其转换为XML实体引用
			formatCode := nf.Code
			formatCode = strings.Replace(formatCode, "\"", "&quot;", -1)
			xml += fmt.Sprintf("    <numFmt numFmtId=\"%d\" formatCode=\"%s\" />\n", nf.ID, formatCode)
		}
		xml += "  </numFmts>\n"
	}

	// 字体
	xml += fmt.Sprintf("  <fonts count=\"%d\">\n", len(s.Fonts))
	for _, font := range s.Fonts {
		xml += "    <font>\n"
		if font.Bold {
			xml += "      <b />\n"
		}
		if font.Italic {
			xml += "      <i />\n"
		}
		if font.Underline {
			xml += "      <u />\n"
		}
		xml += fmt.Sprintf("      <sz val=\"%f\" />\n", font.Size)
		if font.Color != "" {
			xml += fmt.Sprintf("      <color rgb=\"%s\" />\n", font.Color)
		}
		xml += fmt.Sprintf("      <name val=\"%s\" />\n", font.Name)
		xml += "    </font>\n"
	}
	xml += "  </fonts>\n"

	// 填充
	xml += fmt.Sprintf("  <fills count=\"%d\">\n", len(s.Fills))
	for _, fill := range s.Fills {
		xml += "    <fill>\n"
		xml += fmt.Sprintf("      <patternFill patternType=\"%s\">\n", fill.PatternType)
		if fill.FgColor != "" {
			xml += fmt.Sprintf("        <fgColor rgb=\"%s\" />\n", fill.FgColor)
		}
		if fill.BgColor != "" {
			xml += fmt.Sprintf("        <bgColor rgb=\"%s\" />\n", fill.BgColor)
		}
		xml += "      </patternFill>\n"
		xml += "    </fill>\n"
	}
	xml += "  </fills>\n"

	// 边框
	xml += fmt.Sprintf("  <borders count=\"%d\">\n", len(s.Borders))
	for _, border := range s.Borders {
		xml += "    <border>\n"

		// 左边框
		xml += "      <left"
		if border.Left.Style != "" {
			xml += fmt.Sprintf(" style=\"%s\"", border.Left.Style)
		}
		xml += ">\n"
		if border.Left.Color != "" {
			xml += fmt.Sprintf("        <color rgb=\"%s\" />\n", border.Left.Color)
		}
		xml += "      </left>\n"

		// 右边框
		xml += "      <right"
		if border.Right.Style != "" {
			xml += fmt.Sprintf(" style=\"%s\"", border.Right.Style)
		}
		xml += ">\n"
		if border.Right.Color != "" {
			xml += fmt.Sprintf("        <color rgb=\"%s\" />\n", border.Right.Color)
		}
		xml += "      </right>\n"

		// 上边框
		xml += "      <top"
		if border.Top.Style != "" {
			xml += fmt.Sprintf(" style=\"%s\"", border.Top.Style)
		}
		xml += ">\n"
		if border.Top.Color != "" {
			xml += fmt.Sprintf("        <color rgb=\"%s\" />\n", border.Top.Color)
		}
		xml += "      </top>\n"

		// 下边框
		xml += "      <bottom"
		if border.Bottom.Style != "" {
			xml += fmt.Sprintf(" style=\"%s\"", border.Bottom.Style)
		}
		xml += ">\n"
		if border.Bottom.Color != "" {
			xml += fmt.Sprintf("        <color rgb=\"%s\" />\n", border.Bottom.Color)
		}
		xml += "      </bottom>\n"

		xml += "    </border>\n"
	}
	xml += "  </borders>\n"

	// 单元格样式XF
	xml += fmt.Sprintf("  <cellStyleXfs count=\"%d\">\n", len(s.CellStyleXfs))
	for _, xf := range s.CellStyleXfs {
		xml += "    <xf"
		if xf.FontId > 0 {
			xml += fmt.Sprintf(" fontId=\"%d\" applyFont=\"1\"", xf.FontId)
		}
		if xf.FillId > 0 {
			xml += fmt.Sprintf(" fillId=\"%d\" applyFill=\"1\"", xf.FillId)
		}
		if xf.BorderId > 0 {
			xml += fmt.Sprintf(" borderId=\"%d\" applyBorder=\"1\"", xf.BorderId)
		}
		if xf.NumFmtId > 0 {
			xml += fmt.Sprintf(" numFmtId=\"%d\" applyNumberFormat=\"1\"", xf.NumFmtId)
		}
		if xf.Alignment != nil {
			xml += " applyAlignment=\"1\""
		}
		xml += ">\n"
		if xf.Alignment != nil {
			xml += "      <alignment"
			if xf.Alignment.Horizontal != "" {
				xml += fmt.Sprintf(" horizontal=\"%s\"", xf.Alignment.Horizontal)
			}
			if xf.Alignment.Vertical != "" {
				xml += fmt.Sprintf(" vertical=\"%s\"", xf.Alignment.Vertical)
			}
			if xf.Alignment.WrapText {
				xml += " wrapText=\"1\""
			}
			xml += " />\n"
		}
		xml += "    </xf>\n"
	}
	xml += "  </cellStyleXfs>\n"

	// 单元格XF
	xml += fmt.Sprintf("  <cellXfs count=\"%d\">\n", len(s.CellXfs))
	for _, xf := range s.CellXfs {
		xml += "    <xf"

		// 引用已经存在的样式ID
		xml += fmt.Sprintf(" xfId=\"0\"")

		// 设置字体
		if xf.FontId > 0 {
			xml += fmt.Sprintf(" fontId=\"%d\" applyFont=\"1\"", xf.FontId)
		} else {
			xml += " fontId=\"0\""
		}

		// 设置填充
		if xf.FillId > 0 {
			xml += fmt.Sprintf(" fillId=\"%d\" applyFill=\"1\"", xf.FillId)
		} else {
			xml += " fillId=\"0\""
		}

		// 设置边框
		if xf.BorderId > 0 {
			xml += fmt.Sprintf(" borderId=\"%d\" applyBorder=\"1\"", xf.BorderId)
		} else {
			xml += " borderId=\"0\""
		}

		// 设置数字格式
		if xf.NumFmtId > 0 {
			xml += fmt.Sprintf(" numFmtId=\"%d\" applyNumberFormat=\"1\"", xf.NumFmtId)
		} else {
			xml += " numFmtId=\"0\""
		}

		// 设置对齐
		if xf.Alignment != nil {
			xml += " applyAlignment=\"1\""
		}

		xml += ">\n"

		// 添加对齐信息
		if xf.Alignment != nil {
			xml += "      <alignment"
			if xf.Alignment.Horizontal != "" {
				xml += fmt.Sprintf(" horizontal=\"%s\"", xf.Alignment.Horizontal)
			}
			if xf.Alignment.Vertical != "" {
				xml += fmt.Sprintf(" vertical=\"%s\"", xf.Alignment.Vertical)
			}
			if xf.Alignment.WrapText {
				xml += " wrapText=\"1\""
			}
			xml += " />\n"
		}
		xml += "    </xf>\n"
	}
	xml += "  </cellXfs>\n"

	// 单元格样式
	xml += fmt.Sprintf("  <cellStyles count=\"%d\">\n", len(s.CellStyles))
	for _, style := range s.CellStyles {
		xml += fmt.Sprintf("    <cellStyle name=\"%s\" xfId=\"%d\" builtinId=\"%d\"", style.Name, style.XfId, style.BuiltinId)
		if style.CustomBuiltin {
			xml += " customBuiltin=\"1\""
		}
		xml += " />\n"
	}
	xml += "  </cellStyles>\n"

	xml += "</styleSheet>"
	return xml
}

// CreateBorderWithStyle 创建一个边框样式并返回边框ID
func (s *Styles) CreateBorderWithStyle(style, color string) int {
	border := s.AddBorder()
	border.Left.Style = style
	border.Left.Color = color
	border.Right.Style = style
	border.Right.Color = color
	border.Top.Style = style
	border.Top.Color = color
	border.Bottom.Style = style
	border.Bottom.Color = color
	return len(s.Borders) - 1
}

// AddDirectStyleID 直接添加一个CellStyle并返回样式ID
func (s *Styles) AddDirectStyleID(style *CellStyle) int {
	return s.AddCellXf(style.FontID, style.FillID, style.BorderID, style.NumberFormatID, style.Alignment)
}
