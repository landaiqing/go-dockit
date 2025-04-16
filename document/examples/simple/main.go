package main

import (
	"fmt"
	"github.com/landaiqing/go-dockit/document"
	"log"
)

func main() {
	// 创建一个新的Word文档
	doc := document.NewDocument()

	// 设置文档属性
	doc.SetTitle("示例文档")
	doc.SetCreator("FlowDoc库")
	doc.SetDescription("这是一个使用go-dockit库创建的示例文档")

	// 添加一个标题段落
	titlePara := doc.AddParagraph()
	titlePara.SetAlignment("center")
	titleRun := titlePara.AddRun()
	titleRun.AddText("FlowDoc示例文档")
	titleRun.SetBold(true)
	titleRun.SetFontSize(32) // 16磅
	titleRun.SetFontFamily("黑体")

	// 添加一个普通段落
	para1 := doc.AddParagraph()
	para1.SetAlignment("left")
	para1.SetIndentFirstLine(420) // 首行缩进0.3厘米
	run1 := para1.AddRun()
	run1.AddText("这是一个使用go-dockit库创建的示例文档。该库提供了一种简单的方式来生成Word文档，支持段落、表格、列表、图片等元素。")

	// 添加一个带样式的段落
	para2 := doc.AddParagraph()
	para2.SetAlignment("left")
	para2.SetIndentFirstLine(420)
	para2.SetSpacingAfter(200) // 段后间距
	run2 := para2.AddRun()
	run2.AddText("这个段落演示了不同的文本样式：")

	// 添加不同样式的文本
	para2.AddRun().AddText("粗体").SetBold(true)
	para2.AddRun().AddText("、")
	para2.AddRun().AddText("斜体").SetItalic(true)
	para2.AddRun().AddText("、")
	para2.AddRun().AddText("下划线").SetUnderline("single")
	para2.AddRun().AddText("、")
	para2.AddRun().AddText("红色文本").SetColor("FF0000")
	para2.AddRun().AddText("、")
	para2.AddRun().AddText("黄色高亮").SetHighlight("yellow")

	// 添加一个标题
	headingPara := doc.AddParagraph()
	headingPara.SetSpacingBefore(400)
	headingPara.SetSpacingAfter(200)
	headingRun := headingPara.AddRun()
	headingRun.AddText("表格示例")
	headingRun.SetBold(true)
	headingRun.SetFontSize(28) // 14磅

	// 添加一个表格
	table := doc.AddTable(3, 3)
	table.SetWidth(8000, "dxa") // 约14厘米宽
	table.SetAlignment("center")

	// 设置表头
	headerRow := table.Rows[0]
	headerRow.SetIsHeader(true)

	// 填充表头单元格
	headerRow.Cells[0].AddParagraph().AddRun().AddText("产品名称").SetBold(true)
	headerRow.Cells[1].AddParagraph().AddRun().AddText("数量").SetBold(true)
	headerRow.Cells[2].AddParagraph().AddRun().AddText("单价").SetBold(true)

	// 填充表格数据
	table.Rows[1].Cells[0].AddParagraph().AddRun().AddText("产品A")
	table.Rows[1].Cells[1].AddParagraph().AddRun().AddText("10")
	table.Rows[1].Cells[2].AddParagraph().AddRun().AddText("¥100.00")

	table.Rows[2].Cells[0].AddParagraph().AddRun().AddText("产品B")
	table.Rows[2].Cells[1].AddParagraph().AddRun().AddText("5")
	table.Rows[2].Cells[2].AddParagraph().AddRun().AddText("¥200.00")

	// 添加一个分页符
	doc.Body.AddPageBreak()

	// 添加一个标题
	listHeadingPara := doc.AddParagraph()
	listHeadingPara.SetSpacingBefore(400)
	listHeadingPara.SetSpacingAfter(200)
	listHeadingRun := listHeadingPara.AddRun()
	listHeadingRun.AddText("列表示例")
	listHeadingRun.SetBold(true)
	listHeadingRun.SetFontSize(28) // 14磅

	// 创建一个项目符号列表
	bulletListId := doc.Numbering.CreateBulletList()

	// 添加列表项
	listItem1 := doc.AddParagraph()
	listItem1.SetNumbering(bulletListId, 0)
	listItem1.AddRun().AddText("这是第一个列表项")

	listItem2 := doc.AddParagraph()
	listItem2.SetNumbering(bulletListId, 0)
	listItem2.AddRun().AddText("这是第二个列表项")

	listItem3 := doc.AddParagraph()
	listItem3.SetNumbering(bulletListId, 0)
	listItem3.AddRun().AddText("这是第三个列表项")

	// 创建一个数字列表
	numberListId := doc.Numbering.CreateNumberList()

	// 添加列表项
	numListItem1 := doc.AddParagraph()
	numListItem1.SetNumbering(numberListId, 0)
	numListItem1.AddRun().AddText("这是第一个数字列表项")

	numListItem2 := doc.AddParagraph()
	numListItem2.SetNumbering(numberListId, 0)
	numListItem2.AddRun().AddText("这是第二个数字列表项")

	numListItem3 := doc.AddParagraph()
	numListItem3.SetNumbering(numberListId, 0)
	numListItem3.AddRun().AddText("这是第三个数字列表项")

	// 添加页眉并同时添加页眉引用
	header := doc.AddHeaderWithReference("default")
	headerPara := header.AddParagraph()
	headerPara.SetAlignment("right")
	headerPara.AddRun().AddText("FlowDoc示例文档 - 页眉")

	// 添加页脚并同时添加页脚引用
	footer := doc.AddFooterWithReference("default")

	// 使用新的方法添加完整的页码段落
	footer.AddPageNumber()

	// 注释掉旧的方式，它可能会导致页码显示为"第PAGE页"
	// footerPara := footer.AddParagraph()
	// footerPara.SetAlignment("center")
	// footerPara.AddRun().AddText("第 ")
	// footerPara.AddRun().AddPageNumber()
	// footerPara.AddRun().AddText(" 页")

	// 保存文档
	err := doc.Save("./document/examples/simple/example.docx")
	if err != nil {
		log.Fatalf("保存文档时出错: %v", err)
	}

	fmt.Println("文档已成功保存为 example.docx")
}
