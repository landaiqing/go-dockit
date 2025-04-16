package main

import (
	"fmt"
	"github.com/landaiqing/go-dockit/document"
	"log"
	"time"
)

func main() {
	// 创建一个新的Word文档
	doc := document.NewDocument()

	// 设置文档属性
	doc.SetTitle("高级功能演示文档")
	doc.SetCreator("go-dockit库")
	doc.SetDescription("这是一个使用go-dockit库创建的高级功能演示文档")
	doc.SetCreated(time.Now())
	doc.SetModified(time.Now())

	// 设置页面大小为A4纸，竖向
	doc.SetPageSizeA4(false)

	// 设置页面边距
	doc.SetPageMargin(1440, 1440, 1440, 1440, 720, 720, 0)

	// 添加页眉和页脚
	header := doc.AddHeaderWithReference("default")
	headerPara := header.AddParagraph()
	headerPara.SetAlignment("right")
	headerRun := headerPara.AddRun()
	headerRun.AddText("高级功能演示文档")
	headerRun.SetFontSize(20) // 10磅
	headerRun.SetFontFamily("宋体")

	// 添加页脚
	footer := doc.AddFooterWithReference("default")
	// 使用新的方法添加完整的页码段落
	footer.AddPageNumber()

	// =================== 封面页 ===================
	// 添加一个标题段落作为封面
	titlePara := doc.AddParagraph()
	titlePara.SetAlignment("center")
	titlePara.SetSpacingBefore(2400) // 封面标题前空白
	titleRun := titlePara.AddRun()
	titleRun.AddText("高级功能演示文档")
	titleRun.SetBold(true)
	titleRun.SetFontSize(48) // 24磅
	titleRun.SetFontFamily("黑体")

	// 添加副标题
	subtitlePara := doc.AddParagraph()
	subtitlePara.SetAlignment("center")
	subtitlePara.SetSpacingBefore(800)
	subtitleRun := subtitlePara.AddRun()
	subtitleRun.AddText("go-dockit功能展示")
	subtitleRun.SetFontSize(32) // 16磅
	subtitleRun.SetFontFamily("黑体")

	// 添加作者信息
	authorPara := doc.AddParagraph()
	authorPara.SetAlignment("center")
	authorPara.SetSpacingBefore(2400)
	authorRun := authorPara.AddRun()
	authorRun.AddText("作者：go-dockit团队")
	authorRun.SetFontSize(24) // 12磅
	authorRun.SetFontFamily("宋体")

	// 添加日期信息
	datePara := doc.AddParagraph()
	datePara.SetAlignment("center")
	datePara.SetSpacingBefore(200)
	dateRun := datePara.AddRun()
	dateRun.AddText(time.Now().Format("2006年01月02日"))
	dateRun.SetFontSize(24) // 12磅
	dateRun.SetFontFamily("宋体")

	// 添加分页符
	doc.AddPageBreak()

	// =================== 目录页 ===================
	// 添加目录标题
	tocTitlePara := doc.AddParagraph()
	tocTitlePara.SetAlignment("center")
	tocTitlePara.SetSpacingBefore(400)
	tocTitlePara.SetSpacingAfter(600)
	tocTitleRun := tocTitlePara.AddRun()
	tocTitleRun.AddText("目录")
	tocTitleRun.SetBold(true)
	tocTitleRun.SetFontSize(36) // 18磅
	tocTitleRun.SetFontFamily("黑体")

	// 添加目录条目（在实际文档中，这部分会由Word自动更新）
	addTocEntry(doc, "1. 文本样式展示", 0)
	addTocEntry(doc, "2. 段落排版展示", 0)
	addTocEntry(doc, "   2.1 段落对齐方式", 1)
	addTocEntry(doc, "   2.2 段落间距和缩进", 1)
	addTocEntry(doc, "3. 表格展示", 0)
	addTocEntry(doc, "   3.1 基本表格", 1)
	addTocEntry(doc, "   3.2 表格样式", 1)
	addTocEntry(doc, "4. 列表展示", 0)
	addTocEntry(doc, "   4.1 项目符号列表", 1)
	addTocEntry(doc, "   4.2 多级编号列表", 1)

	// 添加分页符
	doc.AddPageBreak()

	// =================== 正文开始 ===================
	// 1. 文本样式展示
	addHeading(doc, "1. 文本样式展示", 1)

	para := doc.AddParagraph()
	para.SetIndentFirstLine(420)
	para.AddRun().AddText("go-dockit库支持多种文本样式，下面展示各种常见的文本格式：")

	// 文本样式示例段落
	stylePara := doc.AddParagraph()
	stylePara.SetIndentFirstLine(420)
	stylePara.SetSpacingBefore(200)
	stylePara.SetSpacingAfter(200)

	// 不同样式的文本展示
	stylePara.AddRun().AddText("正常文本 ")

	boldRun := stylePara.AddRun()
	boldRun.AddText("粗体文本 ")
	boldRun.SetBold(true)

	italicRun := stylePara.AddRun()
	italicRun.AddText("斜体文本 ")
	italicRun.SetItalic(true)

	boldItalicRun := stylePara.AddRun()
	boldItalicRun.AddText("粗斜体文本 ")
	boldItalicRun.SetBold(true)
	boldItalicRun.SetItalic(true)

	underlineRun := stylePara.AddRun()
	underlineRun.AddText("下划线文本 ")
	underlineRun.SetUnderline("single")

	strikeRun := stylePara.AddRun()
	strikeRun.AddText("删除线文本 ")
	strikeRun.SetStrike(true)

	superRun := stylePara.AddRun()
	superRun.AddText("上标")
	superRun.SetVertAlign("superscript")

	stylePara.AddRun().AddText(" 正常文本 ")

	subRun := stylePara.AddRun()
	subRun.AddText("下标")
	subRun.SetVertAlign("subscript")

	// 字体和颜色
	fontPara := doc.AddParagraph()
	fontPara.SetIndentFirstLine(420)

	colorRun := fontPara.AddRun()
	colorRun.AddText("彩色文本 ")
	colorRun.SetColor("FF0000") // 红色

	fontSizeRun := fontPara.AddRun()
	fontSizeRun.AddText("大号文本 ")
	fontSizeRun.SetFontSize(28) // 14磅

	fontFamilyRun := fontPara.AddRun()
	fontFamilyRun.AddText("不同字体 ")
	fontFamilyRun.SetFontFamily("黑体")

	highlightRun := fontPara.AddRun()
	highlightRun.AddText("背景高亮")
	highlightRun.SetHighlight("yellow")

	// 2. 段落排版展示
	addHeading(doc, "2. 段落排版展示", 1)
	addHeading(doc, "2.1 段落对齐方式", 2)

	// 左对齐（默认）
	leftPara := doc.AddParagraph()
	leftPara.SetAlignment("left")
	leftPara.SetSpacingBefore(200)
	leftPara.SetSpacingAfter(200)
	leftPara.AddRun().AddText("这是左对齐的段落。左对齐是最常用的对齐方式，段落的文本从左侧开始排列，右侧自然结束，形成参差不齐的边缘。")

	// 居中对齐
	centerPara := doc.AddParagraph()
	centerPara.SetAlignment("center")
	centerPara.SetSpacingBefore(200)
	centerPara.SetSpacingAfter(200)
	centerPara.AddRun().AddText("这是居中对齐的段落。居中对齐常用于标题、诗歌等需要特殊强调的文本。")

	// 右对齐
	rightPara := doc.AddParagraph()
	rightPara.SetAlignment("right")
	rightPara.SetSpacingBefore(200)
	rightPara.SetSpacingAfter(200)
	rightPara.AddRun().AddText("这是右对齐的段落。右对齐使文本的右边缘整齐，左侧参差不齐。")

	// 两端对齐
	justifyPara := doc.AddParagraph()
	justifyPara.SetAlignment("both")
	justifyPara.SetSpacingBefore(200)
	justifyPara.SetSpacingAfter(200)
	justifyPara.AddRun().AddText("这是两端对齐的段落。两端对齐会使段落的文本在左右两侧都形成整齐的边缘，适合正式文档和出版物。这是两端对齐的段落。两端对齐会使段落的文本在左右两侧都形成整齐的边缘，适合正式文档和出版物。")

	addHeading(doc, "2.2 段落间距和缩进", 2)

	// 解释段落
	spacingExplainPara := doc.AddParagraph()
	spacingExplainPara.SetIndentFirstLine(420)
	spacingExplainPara.AddRun().AddText("段落间距和缩进对于文档的可读性非常重要。以下展示不同的段落间距和缩进效果：")

	// 段落间距示例
	firstSpacingPara := doc.AddParagraph()
	firstSpacingPara.SetIndentFirstLine(420)
	firstSpacingPara.SetSpacingBefore(400) // 段前20磅
	firstSpacingPara.SetSpacingAfter(200)  // 段后10磅
	firstSpacingPara.AddRun().AddText("这个段落有较大的段前间距（20磅）和较小的段后间距（10磅）。通过调整段落间距，可以使文档结构更加清晰。")

	// 行间距示例
	lineSpacingPara := doc.AddParagraph()
	lineSpacingPara.SetIndentFirstLine(420)
	lineSpacingPara.SetSpacingLine(480, "exact") // 24磅的固定行距
	lineSpacingPara.AddRun().AddText("这个段落使用了固定的行间距（24磅）。行间距影响段落内部各行之间的距离，合适的行间距可以提高文本的可读性。这个段落使用了固定的行间距（24磅）。行间距影响段落内部各行之间的距离，合适的行间距可以提高文本的可读性。")

	// 缩进示例
	indentPara := doc.AddParagraph()
	indentPara.SetIndentLeft(720)  // 左侧缩进36磅
	indentPara.SetIndentRight(720) // 右侧缩进36磅
	indentPara.SetSpacingBefore(200)
	indentPara.AddRun().AddText("这个段落左右两侧都有缩进（36磅）。缩进可以用来强调某段文本，或者用于引用格式。不同类型的缩进可以用来表达不同的文档结构和层次。")

	// 悬挂缩进示例
	hangingPara := doc.AddParagraph()
	hangingPara.SetIndentLeft(720)       // 左侧缩进36磅
	hangingPara.SetIndentFirstLine(-420) // 首行缩进-21磅（悬挂缩进）
	hangingPara.SetSpacingBefore(200)
	hangingPara.AddRun().AddText("悬挂缩进：").SetBold(true)
	hangingPara.AddRun().AddText("这个段落使用了悬挂缩进，第一行文本会向左突出，形成一种特殊的格式效果。悬挂缩进常用于定义列表、参考文献等场景。")

	// 3. 表格展示
	addHeading(doc, "3. 表格展示", 1)
	addHeading(doc, "3.1 基本表格", 2)

	// 表格说明段落
	tableExplainPara := doc.AddParagraph()
	tableExplainPara.SetIndentFirstLine(420)
	tableExplainPara.SetSpacingAfter(200)
	tableExplainPara.AddRun().AddText("表格是文档中展示结构化数据的重要方式。以下展示基本表格功能：")

	// 创建一个5行3列的表格
	table := doc.AddTable(5, 3)
	table.SetWidth(8000, "dxa")  // 设置表格宽度
	table.SetAlignment("center") // 表格居中

	// 设置表头
	headerRow := table.Rows[0]
	headerRow.SetIsHeader(true)

	// 填充表头
	headerRow.Cells[0].AddParagraph().AddRun().AddText("产品名称").SetBold(true)
	headerRow.Cells[1].AddParagraph().AddRun().AddText("单价").SetBold(true)
	headerRow.Cells[2].AddParagraph().AddRun().AddText("库存数量").SetBold(true)

	// 填充表格数据
	products := []struct {
		name  string
		price string
		stock string
	}{
		{"商品A", "¥100.00", "150"},
		{"商品B", "¥200.00", "85"},
		{"商品C", "¥150.00", "200"},
		{"商品D", "¥300.00", "35"},
	}

	for i, product := range products {
		row := table.Rows[i+1]

		// 居中显示所有单元格内容
		nameCell := row.Cells[0].AddParagraph()
		nameCell.SetAlignment("center")
		nameCell.AddRun().AddText(product.name)

		priceCell := row.Cells[1].AddParagraph()
		priceCell.SetAlignment("center")
		priceCell.AddRun().AddText(product.price)

		stockCell := row.Cells[2].AddParagraph()
		stockCell.SetAlignment("center")
		stockCell.AddRun().AddText(product.stock)
	}

	// 设置表格边框
	table.SetBorders("all", "single", 4, "000000")

	addHeading(doc, "3.2 表格样式", 2)

	// 表格样式说明段落
	tableStyleExplainPara := doc.AddParagraph()
	tableStyleExplainPara.SetIndentFirstLine(420)
	tableStyleExplainPara.SetSpacingAfter(200)
	tableStyleExplainPara.AddRun().AddText("表格可以应用不同的样式效果，如单元格合并、背景色、对齐方式等：")

	// 创建一个样式化的表格
	styleTable := doc.AddTable(4, 4)
	styleTable.SetWidth(8000, "dxa")  // 设置表格宽度
	styleTable.SetAlignment("center") // 表格居中

	// 设置表头
	styleHeaderRow := styleTable.Rows[0]
	styleHeaderRow.SetIsHeader(true)

	// 设置表头背景色
	for i := 0; i < 4; i++ {
		styleHeaderRow.Cells[i].SetShading("DDDDDD", "000000", "clear")
	}

	// 填充表头
	styleHeaderRow.Cells[0].AddParagraph().AddRun().AddText("季度").SetBold(true)
	styleHeaderRow.Cells[1].AddParagraph().AddRun().AddText("北区").SetBold(true)
	styleHeaderRow.Cells[2].AddParagraph().AddRun().AddText("南区").SetBold(true)
	styleHeaderRow.Cells[3].AddParagraph().AddRun().AddText("总计").SetBold(true)

	// 填充数据
	quarters := []string{"第一季度", "第二季度", "第三季度"}
	northData := []string{"¥10,000", "¥12,000", "¥15,000"}
	southData := []string{"¥8,000", "¥9,000", "¥11,000"}
	totalData := []string{"¥18,000", "¥21,000", "¥26,000"}

	for i := 0; i < 3; i++ {
		row := styleTable.Rows[i+1]

		quarterCell := row.Cells[0].AddParagraph()
		quarterCell.SetAlignment("center")
		quarterCell.AddRun().AddText(quarters[i])

		northCell := row.Cells[1].AddParagraph()
		northCell.SetAlignment("center")
		northCell.AddRun().AddText(northData[i])

		southCell := row.Cells[2].AddParagraph()
		southCell.SetAlignment("center")
		southCell.AddRun().AddText(southData[i])

		totalCell := row.Cells[3].AddParagraph()
		totalCell.SetAlignment("center")
		totalCell.AddRun().AddText(totalData[i])
	}

	// 设置边框
	styleTable.SetBorders("all", "single", 4, "000000")

	// 给某些单元格设置背景色
	styleTable.Rows[1].Cells[1].SetShading("E6F2FF", "000000", "clear") // 浅蓝色
	styleTable.Rows[2].Cells[2].SetShading("E6F2FF", "000000", "clear") // 浅蓝色
	styleTable.Rows[3].Cells[3].SetShading("E6F2FF", "000000", "clear") // 浅蓝色

	// 4. 列表展示
	doc.AddPageBreak() // 分页，避免内容过多
	addHeading(doc, "4. 列表展示", 1)
	addHeading(doc, "4.1 项目符号列表", 2)

	// 列表说明段落
	listExplainPara := doc.AddParagraph()
	listExplainPara.SetIndentFirstLine(420)
	listExplainPara.SetSpacingAfter(200)
	listExplainPara.AddRun().AddText("列表可以清晰地组织和展示相关信息。以下展示项目符号列表：")

	// 创建一个项目符号列表
	bulletListId := doc.Numbering.CreateBulletList()

	// 添加列表项
	bulletItems := []string{
		"项目符号列表项一：项目符号列表用于展示无特定顺序的条目",
		"项目符号列表项二：可以使用不同级别的缩进表示层次结构",
		"项目符号列表项三：适合展示特点、优势等并列信息",
		"项目符号列表项四：列表使文档的结构更加清晰",
	}

	for _, item := range bulletItems {
		listItem := doc.AddParagraph()
		listItem.SetNumbering(bulletListId, 0)
		listItem.AddRun().AddText(item)
	}

	// 创建一个嵌套的项目符号列表
	nestedListPara := doc.AddParagraph()
	nestedListPara.SetSpacingBefore(200)
	nestedListPara.SetIndentFirstLine(420)
	nestedListPara.AddRun().AddText("以下是嵌套的项目符号列表：")

	// 主列表项
	mainItem1 := doc.AddParagraph()
	mainItem1.SetNumbering(bulletListId, 0)
	mainItem1.AddRun().AddText("主要类别A")

	// 嵌套列表项
	subItems1 := []string{"子类别A-1", "子类别A-2", "子类别A-3"}
	for _, item := range subItems1 {
		subItem := doc.AddParagraph()
		subItem.SetNumbering(bulletListId, 1) // 使用级别1表示嵌套
		subItem.AddRun().AddText(item)
	}

	// 另一个主列表项
	mainItem2 := doc.AddParagraph()
	mainItem2.SetNumbering(bulletListId, 0)
	mainItem2.AddRun().AddText("主要类别B")

	// 嵌套列表项
	subItems2 := []string{"子类别B-1", "子类别B-2"}
	for _, item := range subItems2 {
		subItem := doc.AddParagraph()
		subItem.SetNumbering(bulletListId, 1) // 使用级别1表示嵌套
		subItem.AddRun().AddText(item)
	}

	addHeading(doc, "4.2 多级编号列表", 2)

	// 编号列表说明段落
	numberListExplainPara := doc.AddParagraph()
	numberListExplainPara.SetIndentFirstLine(420)
	numberListExplainPara.SetSpacingAfter(200)
	numberListExplainPara.AddRun().AddText("编号列表适合表示有序的步骤或层次分明的结构：")

	// 创建一个数字列表
	numberListId := doc.Numbering.CreateNumberList()

	// 添加列表项
	numberItems := []string{
		"第一步：准备所需材料",
		"第二步：按照说明书组装底座",
		"第三步：连接各个组件",
		"第四步：检查并测试功能",
	}

	for _, item := range numberItems {
		listItem := doc.AddParagraph()
		listItem.SetNumbering(numberListId, 0)
		listItem.AddRun().AddText(item)
	}

	// 创建一个多级编号列表示例
	multiLevelExplainPara := doc.AddParagraph()
	multiLevelExplainPara.SetSpacingBefore(200)
	multiLevelExplainPara.SetIndentFirstLine(420)
	multiLevelExplainPara.AddRun().AddText("以下是多级编号列表示例：")

	// 创建一个多级编号列表
	multiLevelListId := doc.Numbering.CreateNumberList()

	// 一级条目
	level1Item1 := doc.AddParagraph()
	level1Item1.SetNumbering(multiLevelListId, 0)
	level1Item1.AddRun().AddText("第一章：绪论")

	// 二级条目
	level2Item1 := doc.AddParagraph()
	level2Item1.SetNumbering(multiLevelListId, 1)
	level2Item1.AddRun().AddText("研究背景")

	level2Item2 := doc.AddParagraph()
	level2Item2.SetNumbering(multiLevelListId, 1)
	level2Item2.AddRun().AddText("研究意义")

	// 三级条目
	level3Item1 := doc.AddParagraph()
	level3Item1.SetNumbering(multiLevelListId, 2)
	level3Item1.AddRun().AddText("理论意义")

	level3Item2 := doc.AddParagraph()
	level3Item2.SetNumbering(multiLevelListId, 2)
	level3Item2.AddRun().AddText("实践意义")

	// 二级条目
	level2Item3 := doc.AddParagraph()
	level2Item3.SetNumbering(multiLevelListId, 1)
	level2Item3.AddRun().AddText("研究方法")

	// 一级条目
	level1Item2 := doc.AddParagraph()
	level1Item2.SetNumbering(multiLevelListId, 0)
	level1Item2.AddRun().AddText("第二章：文献综述")

	// 结论段落
	doc.AddPageBreak()
	conclusionHeading := doc.AddParagraph()
	conclusionHeading.SetAlignment("center")
	conclusionHeading.SetSpacingBefore(400)
	conclusionHeading.SetSpacingAfter(200)
	conclusionRun := conclusionHeading.AddRun()
	conclusionRun.AddText("总结")
	conclusionRun.SetBold(true)
	conclusionRun.SetFontSize(32) // 16磅

	conclusionPara := doc.AddParagraph()
	conclusionPara.SetIndentFirstLine(420)
	conclusionPara.AddRun().AddText("本文档展示了go-dockit库的各种功能，包括文本样式、段落排版、表格和列表等。通过这些功能，用户可以创建格式丰富、结构清晰的Word文档。欢迎根据需要扩展和完善这些功能。")

	// 保存文档
	err := doc.Save("./document/examples/advanced/advanced_example.docx")
	if err != nil {
		log.Fatalf("保存文档时出错: %v", err)
	}

	fmt.Println("高级功能演示文档已成功保存为 advanced_example.docx")
}

// 辅助函数：添加目录条目
func addTocEntry(doc *document.Document, text string, level int) {
	tocEntryPara := doc.AddParagraph()

	// 根据层级设置不同的缩进
	if level > 0 {
		tocEntryPara.SetIndentLeft(420 * level)
	}

	tocEntryPara.SetSpacingBefore(60)
	tocEntryPara.SetSpacingAfter(60)

	// 添加文本
	run := tocEntryPara.AddRun()
	run.AddText(text)

	// 添加制表符和页码（实际文档中会有页码）
	run.AddTab()
	run.AddText("...")
	run.AddTab()
	run.AddText("1") // 示例页码
}

// 辅助函数：添加标题
func addHeading(doc *document.Document, text string, level int) {
	headingPara := doc.AddParagraph()

	// 设置段落间距
	headingPara.SetSpacingBefore(400)
	headingPara.SetSpacingAfter(200)

	// 添加标题文本
	headingRun := headingPara.AddRun()
	headingRun.AddText(text)
	headingRun.SetBold(true)

	// 根据级别设置不同的字体大小
	switch level {
	case 1:
		headingRun.SetFontSize(32) // 16磅
	case 2:
		headingRun.SetFontSize(28) // 14磅
	case 3:
		headingRun.SetFontSize(26) // 13磅
	default:
		headingRun.SetFontSize(24) // 12磅
	}
}
