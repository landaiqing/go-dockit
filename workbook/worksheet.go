package workbook

import (
	"fmt"
	"time"
)

// Worksheet 表示Excel工作簿中的工作表
type Worksheet struct {
	Name        string
	SheetID     int
	Cells       map[string]*Cell
	Columns     []*Column
	Rows        []*Row
	MergedCells []*MergedCell
}

// NewWorksheet 创建一个新的工作表
func NewWorksheet(name string) *Worksheet {
	return &Worksheet{
		Name:        name,
		SheetID:     1,
		Cells:       make(map[string]*Cell),
		Columns:     make([]*Column, 0),
		Rows:        make([]*Row, 0),
		MergedCells: make([]*MergedCell, 0),
	}
}

// Cell 表示工作表中的单元格
type Cell struct {
	Value    interface{}
	Formula  string
	Style    *CellStyle
	DataType string // s: 字符串, n: 数字, b: 布尔值, d: 日期, e: 错误
}

// NewCell 创建一个新的单元格
func NewCell() *Cell {
	return &Cell{
		Style: NewCellStyle(),
	}
}

// CellStyle 表示单元格样式
type CellStyle struct {
	FontID         int
	FillID         int
	BorderID       int
	NumberFormatID int
	Alignment      *Alignment
}

// NewCellStyle 创建一个新的单元格样式
func NewCellStyle() *CellStyle {
	return &CellStyle{
		Alignment: &Alignment{
			Horizontal: "general",
			Vertical:   "bottom",
		},
	}
}

// Alignment 表示对齐方式
type Alignment struct {
	Horizontal string // left, center, right, fill, justify, centerContinuous, distributed
	Vertical   string // top, center, bottom, justify, distributed
	WrapText   bool
}

// Column 表示工作表中的列
type Column struct {
	Min    int
	Max    int
	Width  float64
	Style  *CellStyle
	Hidden bool
}

// Row 表示工作表中的行
type Row struct {
	Index  int
	Height float64
	Cells  []*Cell
	Style  *CellStyle
	Hidden bool
}

// MergedCell 表示合并的单元格
type MergedCell struct {
	TopLeftRef     string // 例如: "A1"
	BottomRightRef string // 例如: "B2"
}

// AddRow 添加一个新的行
func (ws *Worksheet) AddRow() *Row {
	row := &Row{
		Index: len(ws.Rows) + 1,
		Cells: make([]*Cell, 0),
		Style: NewCellStyle(),
	}
	ws.Rows = append(ws.Rows, row)
	return row
}

// AddColumn 添加一个新的列
func (ws *Worksheet) AddColumn(min, max int, width float64) *Column {
	col := &Column{
		Min:   min,
		Max:   max,
		Width: width,
		Style: NewCellStyle(),
	}
	ws.Columns = append(ws.Columns, col)
	return col
}

// AddCell 在指定位置添加一个单元格
func (ws *Worksheet) AddCell(cellRef string, value interface{}) *Cell {
	cell := NewCell()

	// 根据值类型设置数据类型
	switch v := value.(type) {
	case string:
		cell.DataType = "s"
		cell.Value = value
	case int, int8, int16, int32, int64, uint, uint8, uint16, uint32, uint64, float32, float64:
		cell.DataType = "n"
		cell.Value = value
	case bool:
		cell.DataType = "b"
		cell.Value = value
	case time.Time:
		// 日期和时间在Excel中表示为序列号
		// 数字部分代表天数，小数部分代表一天中的时间
		cell.DataType = "n"

		// 使用正确的Excel日期转换函数
		// GetExcelSerialDate已经处理了日期转换的细节，包括1900年2月29日的Excel错误
		serialDate := GetExcelSerialDate(v)

		// 需要保留小数部分以表示时间
		hours := float64(v.Hour()) / 24.0
		minutes := float64(v.Minute()) / (24.0 * 60.0)
		seconds := float64(v.Second()) / (24.0 * 60.0 * 60.0)

		// 组合日期和时间部分
		cell.Value = serialDate + hours + minutes + seconds
	default:
		cell.DataType = "s"
		cell.Value = fmt.Sprintf("%v", value)
	}

	ws.Cells[cellRef] = cell
	return cell
}

// SetCellFormula 设置单元格公式
func (ws *Worksheet) SetCellFormula(cellRef string, formula string) *Cell {
	cell, ok := ws.Cells[cellRef]
	if !ok {
		cell = ws.AddCell(cellRef, "")
	}
	cell.Formula = formula

	// 公式单元格的初始值设为空，让Excel自动计算
	// 确保DataType不是字符串，这会导致Excel无法正确计算公式
	if cell.DataType == "s" {
		cell.DataType = "" // 让Excel自动判断类型
	}

	return cell
}

// SetCellStyle 设置单元格样式
func (ws *Worksheet) SetCellStyle(cellRef string, style *CellStyle) *Cell {
	cell, ok := ws.Cells[cellRef]
	if !ok {
		cell = ws.AddCell(cellRef, "")
	}

	// 直接设置样式，确保样式引用正确
	cell.Style = style
	return cell
}

// ToXML 将工作表转换为XML
func (ws *Worksheet) ToXML(sharedStrings *SharedStrings) string {
	xml := "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
	xml += "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\n"

	// 列定义
	if len(ws.Columns) > 0 {
		xml += "  <cols>\n"
		for _, col := range ws.Columns {
			xml += fmt.Sprintf("    <col min=\"%d\" max=\"%d\" width=\"%f\" customWidth=\"1\"", col.Min, col.Max, col.Width)
			if col.Hidden {
				xml += " hidden=\"1\""
			}
			xml += "/>\n"
		}
		xml += "  </cols>\n"
	}

	// 单元格数据
	xml += "  <sheetData>\n"

	// 按行组织单元格
	rowMap := make(map[int]map[string]*Cell)
	for cellRef, cell := range ws.Cells {
		rowIndex, colIndex, err := ParseCellRef(cellRef)
		if err != nil {
			continue
		}

		// ParseCellRef返回的是从0开始的索引，而Excel行索引从1开始
		// 将行索引加1，使其与row.Index匹配
		rowNum := rowIndex + 1

		if _, ok := rowMap[rowNum]; !ok {
			rowMap[rowNum] = make(map[string]*Cell)
		}

		// 确保cellRef格式正确，重新生成标准格式的单元格引用
		standardCellRef := CellRef(rowIndex, colIndex)
		rowMap[rowNum][standardCellRef] = cell
	}

	// 收集所有行索引
	rowIndices := make([]int, 0)

	// 添加从单元格收集的行
	for rowIdx := range rowMap {
		rowIndices = append(rowIndices, rowIdx)
	}

	// 添加显式定义的行
	for _, row := range ws.Rows {
		found := false
		for _, idx := range rowIndices {
			if idx == row.Index {
				found = true
				break
			}
		}
		if !found {
			rowIndices = append(rowIndices, row.Index)
		}
	}

	// 按行号排序
	for i := 0; i < len(rowIndices)-1; i++ {
		for j := i + 1; j < len(rowIndices); j++ {
			if rowIndices[i] > rowIndices[j] {
				rowIndices[i], rowIndices[j] = rowIndices[j], rowIndices[i]
			}
		}
	}

	// 输出行和单元格
	for _, rowIdx := range rowIndices {
		// 查找是否有显式添加的行
		var rowHeight float64
		var rowHidden bool

		for _, r := range ws.Rows {
			if r.Index == rowIdx {
				rowHeight = r.Height
				rowHidden = r.Hidden
				break
			}
		}

		xml += fmt.Sprintf("    <row r=\"%d\"", rowIdx)

		if rowHeight > 0 {
			xml += fmt.Sprintf(" ht=\"%f\" customHeight=\"1\"", rowHeight)
		}

		if rowHidden {
			xml += " hidden=\"1\""
		}

		xml += ">\n"

		// 输出该行的单元格
		if cells, ok := rowMap[rowIdx]; ok {
			// 对单元格按列排序
			cellRefs := make([]string, 0, len(cells))
			for cellRef := range cells {
				cellRefs = append(cellRefs, cellRef)
			}

			// 简单排序，确保单元格按列顺序输出
			for i := 0; i < len(cellRefs)-1; i++ {
				for j := i + 1; j < len(cellRefs); j++ {
					_, col1, _ := ParseCellRef(cellRefs[i])
					_, col2, _ := ParseCellRef(cellRefs[j])
					if col1 > col2 {
						cellRefs[i], cellRefs[j] = cellRefs[j], cellRefs[i]
					}
				}
			}

			for _, cellRef := range cellRefs {
				cell := cells[cellRef]

				// 使用标准化的单元格引用
				xml += fmt.Sprintf("      <c r=\"%s\"", cellRef)

				// 正确处理样式ID引用
				if cell.Style != nil {
					// 从样式中获取样式ID
					styleID := 0

					// 首先尝试使用NumberFormatID，这是格式化日期/数字等的关键
					if cell.Style.NumberFormatID > 0 {
						styleID = cell.Style.NumberFormatID
					} else {
						// 如果没有NumberFormatID，则按优先级使用其他样式ID
						if cell.Style.FontID > 0 {
							styleID = cell.Style.FontID
						} else if cell.Style.FillID > 0 {
							styleID = cell.Style.FillID
						} else if cell.Style.BorderID > 0 {
							styleID = cell.Style.BorderID
						}
					}

					if styleID > 0 {
						xml += fmt.Sprintf(" s=\"%d\"", styleID)
					}
				}

				// 设置数据类型
				if cell.DataType != "" {
					// 确保数据类型是有效的Excel类型
					switch cell.DataType {
					case "s":
						xml += " t=\"s\"" // 字符串类型
					case "b":
						xml += " t=\"b\"" // 布尔类型
					case "n":
						// 数字类型不需要特殊的t属性
					default:
						xml += " t=\"s\""
					}
				}

				xml += ">"

				// 添加公式
				if cell.Formula != "" {
					xml += fmt.Sprintf("<f>%s</f>", cell.Formula)

					// 公式单元格不应该标记为字符串类型，除非明确需要
					// 让Excel自动根据公式计算结果决定单元格类型
					if cell.Value != nil {
						xml += fmt.Sprintf("<v>%v</v>", cell.Value)
					}
				} else if cell.Value != nil {
					// 添加值（仅当没有公式时）
					switch cell.DataType {
					case "s": // 字符串
						strValue, ok := cell.Value.(string)
						if ok {
							index := sharedStrings.AddString(strValue)
							xml += fmt.Sprintf("<v>%d</v>", index)
						}
					case "n": // 数字 (包括日期,日期仅是有特殊格式的数字)
						xml += fmt.Sprintf("<v>%v</v>", cell.Value)
					case "b": // 布尔值
						boolValue, ok := cell.Value.(bool)
						if ok {
							if boolValue {
								xml += "<v>1</v>"
							} else {
								xml += "<v>0</v>"
							}
						}
					default:
						// 默认作为字符串处理
						strValue := fmt.Sprintf("%v", cell.Value)
						index := sharedStrings.AddString(strValue)
						xml += fmt.Sprintf("<v>%d</v>", index)
					}
				}

				xml += "</c>\n"
			}
		}

		xml += "    </row>\n"
	}

	xml += "  </sheetData>\n"

	// 合并单元格
	if len(ws.MergedCells) > 0 {
		xml += "  <mergeCells count=\"" + fmt.Sprintf("%d", len(ws.MergedCells)) + "\">\n"
		for _, mergedCell := range ws.MergedCells {
			xml += "    <mergeCell ref=\"" + mergedCell.TopLeftRef + ":" + mergedCell.BottomRightRef + "\" />\n"
		}
		xml += "  </mergeCells>\n"
	}

	xml += "</worksheet>"
	return xml
}

// MergeCells 合并单元格
func (ws *Worksheet) MergeCells(topLeftRef, bottomRightRef string) *MergedCell {
	mergedCell := &MergedCell{
		TopLeftRef:     topLeftRef,
		BottomRightRef: bottomRightRef,
	}
	ws.MergedCells = append(ws.MergedCells, mergedCell)
	return mergedCell
}
