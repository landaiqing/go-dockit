package main

import (
	"fmt"
	"github.com/landaiqing/go-dockit/workbook"
	"time"
)

func main() {
	// 创建一个新的Excel工作簿
	wb := workbook.NewWorkbook()

	// 设置工作簿属性
	wb.Properties.Title = "示例Excel文档"
	wb.Properties.Creator = "Go-DocKit"
	wb.Properties.Created = time.Now()

	// 添加一个工作表
	ws := wb.AddWorksheet("数据报表")

	// 设置列宽
	ws.AddColumn(1, 1, 10) // A列
	ws.AddColumn(2, 2, 20) // B列
	ws.AddColumn(3, 3, 15) // C列
	ws.AddColumn(4, 4, 15) // D列
	ws.AddColumn(5, 5, 20) // E列
	ws.AddColumn(6, 6, 15) // F列 - 百分比列
	ws.AddColumn(7, 7, 15) // G列 - 科学计数列

	// 创建标题样式
	headerStyleID := wb.Styles.CreateStyle(
		"Arial", 12, true, false, false, "FF000000", // 字体
		"solid", "FFD3D3D3", // 填充
		"thin", "FF000000", // 边框
		"",                        // 数字格式
		"center", "center", false, // 对齐
	)

	// 创建日期格式样式 - 使用标准的Excel内置格式
	dateStyleID := wb.Styles.CreateStyle(
		"", 0, false, false, false, "", // 字体
		"", "", // 填充
		"thin", "FF000000", // 边框
		"[$-804]yyyy\"年\"mm\"月\"dd\"日\"", // 中文日期格式
		"center", "bottom", false,        // 对齐
	)

	// 创建人民币货币格式样式
	currencyStyleID := wb.Styles.CreateStyle(
		"", 0, false, false, false, "", // 字体
		"", "", // 填充
		"thin", "FF000000", // 边框
		"¥#,##0.00",              // 人民币货币格式，不使用引号
		"right", "bottom", false, // 对齐
	)

	// 创建百分比格式样式
	percentStyleID := wb.Styles.CreateStyle(
		"", 0, false, false, false, "", // 字体
		"", "", // 填充
		"thin", "FF000000", // 边框
		"0.00%",                  // 百分比格式
		"right", "bottom", false, // 对齐
	)

	// 创建科学计数格式样式 - 使用正确的内置格式
	scientificStyleID := wb.Styles.CreateStyle(
		"", 0, false, false, false, "", // 字体
		"", "", // 填充
		"thin", "FF000000", // 边框
		"0.00E+00",               // 科学计数格式
		"right", "bottom", false, // 对齐
	)

	// 添加标题行
	headers := []string{"编号", "产品名称", "单价(¥)", "数量", "日期", "利润率", "密度"}
	for i, header := range headers {
		cellRef := workbook.CellRef(0, i)
		_ = ws.AddCell(cellRef, header)

		// 设置标题单元格样式
		_ = ws.SetCellStyle(cellRef, &workbook.CellStyle{
			FontID:   headerStyleID,
			BorderID: 0,
			Alignment: &workbook.Alignment{
				Horizontal: "center",
				Vertical:   "center",
			},
		})
	}

	// 添加数据行
	data := [][]interface{}{
		{1, "笔记本电脑", 5999.99, 10, time.Now(), 0.15, 2500000},
		{2, "智能手机", 3999.99, 20, time.Now().AddDate(0, 0, -5), 0.25, 1500000},
		{3, "平板电脑", 2999.99, 15, time.Now().AddDate(0, 0, -10), 0.20, 500000},
		{4, "智能手表", 1999.99, 30, time.Now().AddDate(0, 0, -15), 0.30, 80000},
		{5, "无线耳机", 999.99, 50, time.Now().AddDate(0, 0, -20), 0.40, 5000},
	}

	// 添加数据
	for rowIdx, rowData := range data {
		row := ws.AddRow()
		row.Height = 18

		for colIdx, cellData := range rowData {
			cellRef := workbook.CellRef(rowIdx+1, colIdx)
			_ = ws.AddCell(cellRef, cellData)

			// 根据列类型设置不同的样式
			switch colIdx {
			case 0: // 编号列
				_ = ws.SetCellStyle(cellRef, &workbook.CellStyle{
					Alignment: &workbook.Alignment{
						Horizontal: "center",
					},
				})
			case 2: // 单价列 - 使用人民币格式
				_ = ws.SetCellStyle(cellRef, &workbook.CellStyle{
					NumberFormatID: currencyStyleID,
					Alignment: &workbook.Alignment{
						Horizontal: "right",
					},
				})
			case 4: // 日期列
				_ = ws.SetCellStyle(cellRef, &workbook.CellStyle{
					NumberFormatID: dateStyleID,
					Alignment: &workbook.Alignment{
						Horizontal: "center",
					},
				})
			case 5: // 利润率列 - 使用百分比格式
				_ = ws.SetCellStyle(cellRef, &workbook.CellStyle{
					NumberFormatID: percentStyleID,
					Alignment: &workbook.Alignment{
						Horizontal: "right",
					},
				})
			case 6: // 密度列 - 使用科学计数格式
				_ = ws.SetCellStyle(cellRef, &workbook.CellStyle{
					NumberFormatID: scientificStyleID,
					Alignment: &workbook.Alignment{
						Horizontal: "right",
					},
				})
			}
		}
	}

	// 添加合计行
	ws.AddCell("A7", "合计")
	ws.SetCellFormula("C7", "SUM(C2:C6)")
	ws.SetCellFormula("D7", "SUM(D2:D6)")
	ws.SetCellFormula("F7", "AVERAGE(F2:F6)") // 计算平均利润率

	// 设置合计行样式
	ws.SetCellStyle("A7", &workbook.CellStyle{
		FontID: 0,
		Alignment: &workbook.Alignment{
			Horizontal: "right",
		},
	})
	ws.SetCellStyle("C7", &workbook.CellStyle{
		FontID:         0,
		NumberFormatID: currencyStyleID, // 使用人民币格式
		Alignment: &workbook.Alignment{
			Horizontal: "right",
		},
	})
	ws.SetCellStyle("D7", &workbook.CellStyle{
		FontID: 0,
		Alignment: &workbook.Alignment{
			Horizontal: "center",
		},
	})
	ws.SetCellStyle("F7", &workbook.CellStyle{
		FontID:         0,
		NumberFormatID: percentStyleID, // 使用百分比格式
		Alignment: &workbook.Alignment{
			Horizontal: "right",
		},
	})

	// 添加第二个工作表 - 用于额外的测试
	testSheet := wb.AddWorksheet("格式测试")

	// 设置标题样式
	testSheet.AddColumn(1, 7, 15) // 统一设置列宽

	// 测试表标题
	testHeaders := []string{"类型", "值", "描述", "格式ID", "样式", "显示效果", "备注"}
	for i, header := range testHeaders {
		cellRef := workbook.CellRef(0, i)
		testSheet.AddCell(cellRef, header)
		testSheet.SetCellStyle(cellRef, &workbook.CellStyle{
			FontID: headerStyleID,
			Alignment: &workbook.Alignment{
				Horizontal: "center",
			},
		})
	}

	// 1. 测试日期格式
	testSheet.AddCell("A2", "日期")
	testNow := time.Now()
	_ = testSheet.AddCell("B2", testNow)
	testSheet.AddCell("C2", "当前日期时间")
	testSheet.AddCell("D2", fmt.Sprintf("%d", dateStyleID))
	testSheet.AddCell("E2", "[$-804]yyyy\"年\"mm\"月\"dd\"日\"")
	testSheet.AddCell("G2", "应显示为中文日期格式")
	testSheet.SetCellStyle("B2", &workbook.CellStyle{
		NumberFormatID: dateStyleID,
	})

	// 2. 测试科学计数格式 - 大数值
	testSheet.AddCell("A3", "科学计数")
	testSheet.AddCell("B3", 12345678.9)
	testSheet.AddCell("C3", "大数值")
	testSheet.AddCell("D3", fmt.Sprintf("%d", scientificStyleID))
	testSheet.AddCell("E3", "0.00E+00")
	testSheet.AddCell("G3", "应显示为1.23E+07")
	testSheet.SetCellStyle("B3", &workbook.CellStyle{
		NumberFormatID: scientificStyleID,
	})

	// 3. 测试科学计数格式 - 小数值
	testSheet.AddCell("A4", "科学计数")
	testSheet.AddCell("B4", 0.00000123)
	testSheet.AddCell("C4", "小数值")
	testSheet.AddCell("D4", fmt.Sprintf("%d", scientificStyleID))
	testSheet.AddCell("E4", "0.00E+00")
	testSheet.AddCell("G4", "应显示为1.23E-06")
	testSheet.SetCellStyle("B4", &workbook.CellStyle{
		NumberFormatID: scientificStyleID,
	})

	// 4. 测试百分比格式
	testSheet.AddCell("A5", "百分比")
	testSheet.AddCell("B5", 0.1234)
	testSheet.AddCell("C5", "小数值")
	testSheet.AddCell("D5", fmt.Sprintf("%d", percentStyleID))
	testSheet.AddCell("E5", "0.00%")
	testSheet.AddCell("G5", "应显示为12.34%")
	testSheet.SetCellStyle("B5", &workbook.CellStyle{
		NumberFormatID: percentStyleID,
	})

	// 合并单元格示例
	ws.MergeCells("A9", "G9")
	ws.AddCell("A9", "销售数据分析报表")

	// 设置合并单元格的样式
	ws.SetCellStyle("A9", &workbook.CellStyle{
		FontID: 0,
		Alignment: &workbook.Alignment{
			Horizontal: "center",
			Vertical:   "center",
		},
	})

	// 保存Excel文件
	err := wb.Save("sales_report.xlsx")
	if err != nil {
		fmt.Println("保存Excel文件时出错:", err)
		return
	}

	fmt.Println("Excel文件已成功创建: sales_report.xlsx")
}
