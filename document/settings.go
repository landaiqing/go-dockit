package document

import (
	"fmt"
)

// Settings 表示Word文档中的设置
type Settings struct {
	UpdateFields            bool   // 更新域
	Zoom                    int    // 缩放比例
	DefaultTabStop          int    // 默认制表位
	CharacterSpacingControl string // 字符间距控制
	Compatibility           *Compatibility
}

// Compatibility 表示兼容性设置
type Compatibility struct {
	CompatibilityMode                string // 兼容模式
	DoNotExpandShiftReturn           bool   // 不展开Shift+Enter
	DoNotBreakWrappedTables          bool   // 不断开环绕表格
	DoNotSnapToGridInCell            bool   // 单元格中不对齐网格
	DoNotWrapTextWithPunct           bool   // 不使用标点符号换行
	DoNotUseEastAsianBreakRules      bool   // 不使用东亚换行规则
	DoNotUseIndentAsNumberingTabStop bool   // 不使用缩进作为编号制表位
	UseAnsiKerningPairs              bool   // 使用ANSI字距调整对
	DoNotAutofitConstrainedTables    bool   // 不自动调整受限表格
	SplitPgBreakAndParaMark          bool   // 分割分页符和段落标记
	DoNotVertAlignCellWithSp         bool   // 不垂直对齐带有形状的单元格
	DoNotBreakConstrainedForcedTable bool   // 不断开受限强制表格
	DoNotVertAlignInTxbx             bool   // 不在文本框中垂直对齐
	UseAnsiSpaceForEnglishInEastAsia bool   // 在东亚语言中为英文使用ANSI空格
	AllowSpaceOfSameStyleInTable     bool   // 允许表格中相同样式的空格
	DoNotSuppressIndentation         bool   // 不抑制缩进
	DoNotAutospaceEastAsianText      bool   // 不自动调整东亚文本间距
	DoNotUseHTMLParagraphAutoSpacing bool   // 不使用HTML段落自动间距
}

// NewSettings 创建一个新的设置
func NewSettings() *Settings {
	return &Settings{
		UpdateFields:            true,
		Zoom:                    100,
		DefaultTabStop:          720, // 720 twip = 0.5 inch
		CharacterSpacingControl: "doNotCompress",
		Compatibility:           NewCompatibility(),
	}
}

// NewCompatibility 创建一个新的兼容性设置
func NewCompatibility() *Compatibility {
	return &Compatibility{
		CompatibilityMode: "15", // Word 2013
	}
}

// SetUpdateFields 设置是否更新域
func (s *Settings) SetUpdateFields(updateFields bool) *Settings {
	s.UpdateFields = updateFields
	return s
}

// SetZoom 设置缩放比例
func (s *Settings) SetZoom(zoom int) *Settings {
	s.Zoom = zoom
	return s
}

// SetDefaultTabStop 设置默认制表位
func (s *Settings) SetDefaultTabStop(defaultTabStop int) *Settings {
	s.DefaultTabStop = defaultTabStop
	return s
}

// SetCharacterSpacingControl 设置字符间距控制
func (s *Settings) SetCharacterSpacingControl(characterSpacingControl string) *Settings {
	s.CharacterSpacingControl = characterSpacingControl
	return s
}

// SetCompatibilityMode 设置兼容模式
func (s *Settings) SetCompatibilityMode(compatibilityMode string) *Settings {
	s.Compatibility.CompatibilityMode = compatibilityMode
	return s
}

// ToXML 将设置转换为XML
func (s *Settings) ToXML() string {
	xml := "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
	xml += "<w:settings xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">"

	// 更新域
	if s.UpdateFields {
		xml += "<w:updateFields w:val=\"true\" />"
	}

	// 缩放比例
	xml += "<w:zoom w:percent=\"" + fmt.Sprintf("%d", s.Zoom) + "\" />"

	// 默认制表位
	xml += "<w:defaultTabStop w:val=\"" + fmt.Sprintf("%d", s.DefaultTabStop) + "\" />"

	// 字符间距控制
	xml += "<w:characterSpacingControl w:val=\"" + s.CharacterSpacingControl + "\" />"

	// 兼容性设置
	xml += "<w:compat>"

	// 兼容模式
	if s.Compatibility.CompatibilityMode != "" {
		xml += "<w:compatSetting w:name=\"compatibilityMode\" w:uri=\"http://schemas.microsoft.com/office/document\" w:val=\"" + s.Compatibility.CompatibilityMode + "\" />"
	}

	// 不展开Shift+Enter
	if s.Compatibility.DoNotExpandShiftReturn {
		xml += "<w:doNotExpandShiftReturn />"
	}

	// 不断开环绕表格
	if s.Compatibility.DoNotBreakWrappedTables {
		xml += "<w:doNotBreakWrappedTables />"
	}

	// 单元格中不对齐网格
	if s.Compatibility.DoNotSnapToGridInCell {
		xml += "<w:doNotSnapToGridInCell />"
	}

	// 不使用标点符号换行
	if s.Compatibility.DoNotWrapTextWithPunct {
		xml += "<w:doNotWrapTextWithPunct />"
	}

	// 不使用东亚换行规则
	if s.Compatibility.DoNotUseEastAsianBreakRules {
		xml += "<w:doNotUseEastAsianBreakRules />"
	}

	// 不使用缩进作为编号制表位
	if s.Compatibility.DoNotUseIndentAsNumberingTabStop {
		xml += "<w:doNotUseIndentAsNumberingTabStop />"
	}

	// 使用ANSI字距调整对
	if s.Compatibility.UseAnsiKerningPairs {
		xml += "<w:useAnsiKerningPairs />"
	}

	// 不自动调整受限表格
	if s.Compatibility.DoNotAutofitConstrainedTables {
		xml += "<w:doNotAutofitConstrainedTables />"
	}

	// 分割分页符和段落标记
	if s.Compatibility.SplitPgBreakAndParaMark {
		xml += "<w:splitPgBreakAndParaMark />"
	}

	// 不垂直对齐带有形状的单元格
	if s.Compatibility.DoNotVertAlignCellWithSp {
		xml += "<w:doNotVertAlignCellWithSp />"
	}

	// 不断开受限强制表格
	if s.Compatibility.DoNotBreakConstrainedForcedTable {
		xml += "<w:doNotBreakConstrainedForcedTable />"
	}

	// 不在文本框中垂直对齐
	if s.Compatibility.DoNotVertAlignInTxbx {
		xml += "<w:doNotVertAlignInTxbx />"
	}

	// 在东亚语言中为英文使用ANSI空格
	if s.Compatibility.UseAnsiSpaceForEnglishInEastAsia {
		xml += "<w:useAnsiSpaceForEnglishInEastAsia />"
	}

	// 允许表格中相同样式的空格
	if s.Compatibility.AllowSpaceOfSameStyleInTable {
		xml += "<w:allowSpaceOfSameStyleInTable />"
	}

	// 不抑制缩进
	if s.Compatibility.DoNotSuppressIndentation {
		xml += "<w:doNotSuppressIndentation />"
	}

	// 不自动调整东亚文本间距
	if s.Compatibility.DoNotAutospaceEastAsianText {
		xml += "<w:doNotAutospaceEastAsianText />"
	}

	// 不使用HTML段落自动间距
	if s.Compatibility.DoNotUseHTMLParagraphAutoSpacing {
		xml += "<w:doNotUseHTMLParagraphAutoSpacing />"
	}

	xml += "</w:compat>"

	xml += "</w:settings>"
	return xml
}
