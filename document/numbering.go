package document

import (
	"fmt"
)

// Numbering 表示Word文档中的编号集合
type Numbering struct {
	AbstractNums []*AbstractNum
	Nums         []*Num
}

// AbstractNum 表示抽象编号
type AbstractNum struct {
	ID     int
	Levels []*NumberingLevel
}

// Num 表示具体编号
type Num struct {
	ID             int
	AbstractNumID  int
	LevelOverrides []*LevelOverride
}

// LevelOverride 表示级别覆盖
type LevelOverride struct {
	Level          int
	StartAt        int
	NumberingLevel *NumberingLevel
}

// NumberingLevel 表示编号级别
type NumberingLevel struct {
	Level           int
	Start           int
	NumberingFormat string // decimal, upperRoman, lowerRoman, upperLetter, lowerLetter, bullet, etc.
	Text            string // 编号文本，如 "%1."
	Justification   string // left, center, right
	ParagraphStyle  string // 段落样式ID
	Font            string // 字体
	Indent          int    // 缩进
	HangingIndent   int    // 悬挂缩进
	TabStop         int    // 制表位
	Suffix          string // tab, space, nothing
}

// NewNumbering 创建一个新的编号集合
func NewNumbering() *Numbering {
	return &Numbering{
		AbstractNums: make([]*AbstractNum, 0),
		Nums:         make([]*Num, 0),
	}
}

// AddAbstractNum 添加一个抽象编号
func (n *Numbering) AddAbstractNum() *AbstractNum {
	abstractNum := &AbstractNum{
		ID:     len(n.AbstractNums) + 1,
		Levels: make([]*NumberingLevel, 0),
	}
	n.AbstractNums = append(n.AbstractNums, abstractNum)
	return abstractNum
}

// AddNum 添加一个具体编号
func (n *Numbering) AddNum(abstractNumID int) *Num {
	num := &Num{
		ID:             len(n.Nums) + 1,
		AbstractNumID:  abstractNumID,
		LevelOverrides: make([]*LevelOverride, 0),
	}
	n.Nums = append(n.Nums, num)
	return num
}

// AddLevel 向抽象编号添加一个级别
func (a *AbstractNum) AddLevel(level int) *NumberingLevel {
	numberingLevel := &NumberingLevel{
		Level:           level,
		Start:           1,
		NumberingFormat: "decimal",
		Text:            "%" + fmt.Sprintf("%d", level+1) + ".",
		Justification:   "left",
		Indent:          720 * (level + 1), // 720 twip = 0.5 inch
		HangingIndent:   360,               // 360 twip = 0.25 inch
		TabStop:         720 * (level + 1),
		Suffix:          "tab",
	}
	a.Levels = append(a.Levels, numberingLevel)
	return numberingLevel
}

// AddLevelOverride 向具体编号添加一个级别覆盖
func (n *Num) AddLevelOverride(level int) *LevelOverride {
	levelOverride := &LevelOverride{
		Level:   level,
		StartAt: 1,
	}
	n.LevelOverrides = append(n.LevelOverrides, levelOverride)
	return levelOverride
}

// SetStart 设置编号级别的起始值
func (l *NumberingLevel) SetStart(start int) *NumberingLevel {
	l.Start = start
	return l
}

// SetNumberingFormat 设置编号级别的格式
func (l *NumberingLevel) SetNumberingFormat(format string) *NumberingLevel {
	l.NumberingFormat = format
	return l
}

// SetText 设置编号级别的文本
func (l *NumberingLevel) SetText(text string) *NumberingLevel {
	l.Text = text
	return l
}

// SetJustification 设置编号级别的对齐方式
func (l *NumberingLevel) SetJustification(justification string) *NumberingLevel {
	l.Justification = justification
	return l
}

// SetParagraphStyle 设置编号级别的段落样式
func (l *NumberingLevel) SetParagraphStyle(style string) *NumberingLevel {
	l.ParagraphStyle = style
	return l
}

// SetFont 设置编号级别的字体
func (l *NumberingLevel) SetFont(font string) *NumberingLevel {
	l.Font = font
	return l
}

// SetIndent 设置编号级别的缩进
func (l *NumberingLevel) SetIndent(indent int) *NumberingLevel {
	l.Indent = indent
	return l
}

// SetHangingIndent 设置编号级别的悬挂缩进
func (l *NumberingLevel) SetHangingIndent(hangingIndent int) *NumberingLevel {
	l.HangingIndent = hangingIndent
	return l
}

// SetTabStop 设置编号级别的制表位
func (l *NumberingLevel) SetTabStop(tabStop int) *NumberingLevel {
	l.TabStop = tabStop
	return l
}

// SetSuffix 设置编号级别的后缀
func (l *NumberingLevel) SetSuffix(suffix string) *NumberingLevel {
	l.Suffix = suffix
	return l
}

// SetStartAt 设置级别覆盖的起始值
func (l *LevelOverride) SetStartAt(startAt int) *LevelOverride {
	l.StartAt = startAt
	return l
}

// SetNumberingLevel 设置级别覆盖的编号级别
func (l *LevelOverride) SetNumberingLevel(level *NumberingLevel) *LevelOverride {
	l.NumberingLevel = level
	return l
}

// CreateBulletList 创建一个项目符号列表
func (n *Numbering) CreateBulletList() int {
	// 创建抽象编号
	abstractNum := n.AddAbstractNum()

	// 添加9个级别
	for i := 0; i < 9; i++ {
		level := abstractNum.AddLevel(i)
		level.SetNumberingFormat("bullet")

		// 根据级别设置不同的项目符号
		switch i % 3 {
		case 0:
			level.SetText("•")
			level.SetFont("Symbol")
		case 1:
			level.SetText("○")
			level.SetFont("Courier New")
		case 2:
			level.SetText("▪")
			level.SetFont("Wingdings")
		}
	}

	// 创建具体编号
	num := n.AddNum(abstractNum.ID)

	return num.ID
}

// CreateNumberList 创建一个数字列表
func (n *Numbering) CreateNumberList() int {
	// 创建抽象编号
	abstractNum := n.AddAbstractNum()

	// 添加9个级别
	for i := 0; i < 9; i++ {
		level := abstractNum.AddLevel(i)

		// 根据级别设置不同的编号格式
		switch i % 3 {
		case 0:
			level.SetNumberingFormat("decimal")
			level.SetText("%" + fmt.Sprintf("%d", i+1) + ".")
		case 1:
			level.SetNumberingFormat("lowerLetter")
			level.SetText("%" + fmt.Sprintf("%d", i+1) + ").")
		case 2:
			level.SetNumberingFormat("lowerRoman")
			level.SetText("%" + fmt.Sprintf("%d", i+1) + ").")
		}
	}

	// 创建具体编号
	num := n.AddNum(abstractNum.ID)

	return num.ID
}

// ToXML 将编号集合转换为XML
func (n *Numbering) ToXML() string {
	xml := "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
	xml += "<w:numbering xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">"

	// 添加所有抽象编号
	for _, abstractNum := range n.AbstractNums {
		xml += "<w:abstractNum w:abstractNumId=\"" + fmt.Sprintf("%d", abstractNum.ID) + "\">"

		// 添加所有级别
		for _, level := range abstractNum.Levels {
			xml += "<w:lvl w:ilvl=\"" + fmt.Sprintf("%d", level.Level) + "\">"

			// 起始值
			xml += "<w:start w:val=\"" + fmt.Sprintf("%d", level.Start) + "\" />"

			// 编号格式
			xml += "<w:numFmt w:val=\"" + level.NumberingFormat + "\" />"

			// 编号文本
			xml += "<w:lvlText w:val=\"" + level.Text + "\" />"

			// 对齐方式
			xml += "<w:lvlJc w:val=\"" + level.Justification + "\" />"

			// 段落属性
			xml += "<w:pPr>"

			// 缩进
			xml += "<w:ind w:left=\"" + fmt.Sprintf("%d", level.Indent) + "\""
			xml += " w:hanging=\"" + fmt.Sprintf("%d", level.HangingIndent) + "\" />"

			// 制表位
			if level.TabStop > 0 {
				xml += "<w:tabs>"
				xml += "<w:tab w:val=\"num\" w:pos=\"" + fmt.Sprintf("%d", level.TabStop) + "\" />"
				xml += "</w:tabs>"
			}

			xml += "</w:pPr>"

			// 文本运行属性
			xml += "<w:rPr>"

			// 字体
			if level.Font != "" {
				xml += "<w:rFonts w:ascii=\"" + level.Font + "\""
				xml += " w:hAnsi=\"" + level.Font + "\""
				xml += " w:hint=\"default\" />"
			}

			xml += "</w:rPr>"

			// 后缀
			if level.Suffix != "" {
				xml += "<w:suff w:val=\"" + level.Suffix + "\" />"
			}

			xml += "</w:lvl>"
		}

		xml += "</w:abstractNum>"
	}

	// 添加所有具体编号
	for _, num := range n.Nums {
		xml += "<w:num w:numId=\"" + fmt.Sprintf("%d", num.ID) + "\">"

		// 抽象编号ID
		xml += "<w:abstractNumId w:val=\"" + fmt.Sprintf("%d", num.AbstractNumID) + "\" />"

		// 添加所有级别覆盖
		for _, levelOverride := range num.LevelOverrides {
			xml += "<w:lvlOverride w:ilvl=\"" + fmt.Sprintf("%d", levelOverride.Level) + "\">"

			// 起始值
			if levelOverride.StartAt > 0 {
				xml += "<w:startOverride w:val=\"" + fmt.Sprintf("%d", levelOverride.StartAt) + "\" />"
			}

			// 编号级别
			if levelOverride.NumberingLevel != nil {
				xml += "<w:lvl w:ilvl=\"" + fmt.Sprintf("%d", levelOverride.NumberingLevel.Level) + "\">"

				// 起始值
				xml += "<w:start w:val=\"" + fmt.Sprintf("%d", levelOverride.NumberingLevel.Start) + "\" />"

				// 编号格式
				xml += "<w:numFmt w:val=\"" + levelOverride.NumberingLevel.NumberingFormat + "\" />"

				// 编号文本
				xml += "<w:lvlText w:val=\"" + levelOverride.NumberingLevel.Text + "\" />"

				// 对齐方式
				xml += "<w:lvlJc w:val=\"" + levelOverride.NumberingLevel.Justification + "\" />"

				// 段落属性
				xml += "<w:pPr>"

				// 缩进
				xml += "<w:ind w:left=\"" + fmt.Sprintf("%d", levelOverride.NumberingLevel.Indent) + "\""
				xml += " w:hanging=\"" + fmt.Sprintf("%d", levelOverride.NumberingLevel.HangingIndent) + "\" />"

				// 制表位
				if levelOverride.NumberingLevel.TabStop > 0 {
					xml += "<w:tabs>"
					xml += "<w:tab w:val=\"num\" w:pos=\"" + fmt.Sprintf("%d", levelOverride.NumberingLevel.TabStop) + "\" />"
					xml += "</w:tabs>"
				}

				xml += "</w:pPr>"

				// 文本运行属性
				xml += "<w:rPr>"

				// 字体
				if levelOverride.NumberingLevel.Font != "" {
					xml += "<w:rFonts w:ascii=\"" + levelOverride.NumberingLevel.Font + "\""
					xml += " w:hAnsi=\"" + levelOverride.NumberingLevel.Font + "\""
					xml += " w:hint=\"default\" />"
				}

				xml += "</w:rPr>"

				// 后缀
				if levelOverride.NumberingLevel.Suffix != "" {
					xml += "<w:suff w:val=\"" + levelOverride.NumberingLevel.Suffix + "\" />"
				}

				xml += "</w:lvl>"
			}

			xml += "</w:lvlOverride>"
		}

		xml += "</w:num>"
	}

	xml += "</w:numbering>"
	return xml
}
