package main

import (
	"fmt"
	"log"

	"github.com/landaiqing/go-dockit/document"
)

func main() {
	// 创建一个新的Word文档
	doc := document.NewDocument()

	// 设置文档属性
	doc.SetTitle("数据库三线表示例")
	doc.SetCreator("go-dockit库")
	doc.SetDescription("使用go-dockit库创建学术论文中常用的数据库三线表")

	// 添加标题
	titlePara := doc.AddParagraph()
	titlePara.SetAlignment("center")
	titlePara.SetSpacingAfter(200)
	titleRun := titlePara.AddRun()
	titleRun.AddText("数据库三线表示例")
	titleRun.SetBold(true)
	titleRun.SetFontSize(32) // 16磅
	titleRun.SetFontFamily("宋体")

	// 添加说明文字
	explainPara := doc.AddParagraph()
	explainPara.SetIndentFirstLine(420)
	explainPara.SetSpacingAfter(300)
	explainPara.AddRun().AddText("数据库三线表是学术论文中常用的一种表格格式，特点是只有三根水平线（顶线、表头分隔线和底线），没有垂直线分隔列。这种表格格式符合APA（美国心理学会）和许多学术期刊的规范。")

	// ========== 示例1：基本三线表 ==========
	// 添加表格标题（通常三线表的标题在表格上方）
	tableTitlePara := doc.AddParagraph()
	tableTitlePara.SetAlignment("center")
	tableTitlePara.SetSpacingAfter(0)
	tableTitlePara.SetSpacingBefore(0)
	tableTitlePara.SetLineSpacing(1.5, "auto") // 设置1.5倍行距
	tableTitleRun := tableTitlePara.AddRun()
	tableTitleRun.AddText("表4-1 文章信息表")
	tableTitleRun.SetBold(true)
	tableTitleRun.SetFontSize(21) // 五号字体约为10.5磅(21)
	tableTitleRun.SetFontFamily("宋体")
	// 表序号设置为Times New Roman
	tableTitleRun.SetFontFamilyForRunes("Times New Roman", []rune("表4-1"))

	// 创建一个表格
	table1 := doc.AddTable(6, 5)
	table1.SetWidth("100%", "pct") // 与文字齐宽
	table1.SetAlignment("center")

	// 填充表头
	headers := []string{"字段", "字段名", "类型", "长度", "非空"}
	for i, header := range headers {
		cellPara := table1.Rows[0].Cells[i].AddParagraph()
		cellPara.SetAlignment("center")
		cellPara.SetLineSpacing(1.5, "auto") // 1.5倍行距
		cellRun := cellPara.AddRun()
		cellRun.AddText(header)
		cellRun.SetBold(false)
		cellRun.SetFontSize(21) // 五号字体
		cellRun.SetFontFamily("宋体")
	}

	// 填充数据行
	data := [][]string{
		{"标题", "title", "varchar", "100", "是"},
		{"文章分类", "sort", "varchar", "150", "是"},
		{"作者学号", "author_sn", "varchar", "100", "是"},
		{"作者姓名", "author_name", "varchar", "100", "否"},
		{"文章内容", "description", "longtext", "默认", "否"},
	}

	for i, row := range data {
		for j, cell := range row {
			para := table1.Rows[i+1].Cells[j].AddParagraph()
			para.SetAlignment("center")
			para.SetLineSpacing(1.5, "auto") // 1.5倍行距
			cellRun := para.AddRun()
			cellRun.AddText(cell)
			cellRun.SetFontSize(21) // 五号字体
			cellRun.SetFontFamily("宋体")
			// 英文字体设置为Times New Roman
			if j == 1 || j == 2 { // 字段名和类型列通常包含英文
				cellRun.SetFontFamilyForRunes("Times New Roman", []rune(cell))
			}
		}
	}

	// 设置三线表样式
	// 1. 首先清除所有默认边框
	table1.SetBorders("all", "", 0, "")

	// 2. 顶线（表格顶部边框），1.5磅
	table1.SetBorders("top", "single", 10, "000000")

	// 3. 表头分隔线（第一行底部边框），1磅
	for i := 0; i < 5; i++ {
		table1.Rows[0].Cells[i].SetBorders("bottom", "single", 4, "000000")
	}

	// 4. 底线（表格底部边框），1.5磅
	table1.SetBorders("bottom", "single", 10, "000000")

	// 显式设置内部边框为"none"，而不是空字符串
	table1.SetBorders("insideH", "none", 0, "000000")
	table1.SetBorders("insideV", "none", 0, "000000")

	// 添加页脚（页码）
	footer := doc.AddFooterWithReference("default")
	footer.AddPageNumber()

	// 保存文档
	err := doc.Save("./document/examples/threeline_table/threeline_table_example.docx")
	if err != nil {
		log.Fatalf("保存文档时出错: %v", err)
	}

	fmt.Println("数据库三线表示例文档已成功保存为 threeline_table_example.docx")
}
