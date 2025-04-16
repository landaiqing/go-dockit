package workbook

import (
	"archive/zip"
	"fmt"
	"os"
	"time"
)

// Workbook 表示一个Excel工作簿
type Workbook struct {
	Worksheets    []*Worksheet
	Properties    *WorkbookProperties
	Relationships *Relationships
	Styles        *Styles
	Theme         *Theme
	ContentTypes  *ContentTypes
	Rels          *WorkbookRels
	SharedStrings *SharedStrings
}

// WorkbookProperties 包含工作簿的元数据
type WorkbookProperties struct {
	Title          string
	Subject        string
	Creator        string
	Keywords       string
	Description    string
	LastModifiedBy string
	Revision       int
	Created        time.Time
	Modified       time.Time
}

// NewWorkbook 创建一个新的Excel工作簿
func NewWorkbook() *Workbook {
	return &Workbook{
		Worksheets: make([]*Worksheet, 0),
		Properties: &WorkbookProperties{
			Created:  time.Now(),
			Modified: time.Now(),
			Revision: 1,
		},
		Relationships: NewRelationships(),
		Styles:        NewStyles(),
		Theme:         NewTheme(),
		ContentTypes:  NewContentTypes(),
		Rels:          NewWorkbookRels(),
		SharedStrings: NewSharedStrings(),
	}
}

// AddWorksheet 添加一个新的工作表
func (wb *Workbook) AddWorksheet(name string) *Worksheet {
	ws := NewWorksheet(name)
	ws.SheetID = len(wb.Worksheets) + 1
	wb.Worksheets = append(wb.Worksheets, ws)
	return ws
}

// Save 保存Excel工作簿到文件
func (wb *Workbook) Save(filename string) error {
	// 创建一个新的zip文件
	file, err := os.Create(filename)
	if err != nil {
		return err
	}
	defer file.Close()

	// 创建一个新的zip writer
	zipWriter := zip.NewWriter(file)
	defer zipWriter.Close()

	// 添加[Content_Types].xml
	contentTypesWriter, err := zipWriter.Create("[Content_Types].xml")
	if err != nil {
		return err
	}
	_, err = contentTypesWriter.Write([]byte(wb.ContentTypes.ToXML()))
	if err != nil {
		return err
	}

	// 添加_rels/.rels
	relsDir, err := zipWriter.Create("_rels/.rels")
	if err != nil {
		return err
	}

	// 创建根关系
	rootRels := NewRelationships()
	rootRels.AddRelationship("rId1", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument", "xl/workbook.xml")
	_, err = relsDir.Write([]byte(rootRels.ToXML()))
	if err != nil {
		return err
	}

	// 添加xl/workbook.xml
	workbookWriter, err := zipWriter.Create("xl/workbook.xml")
	if err != nil {
		return err
	}
	_, err = workbookWriter.Write([]byte(wb.ToXML()))
	if err != nil {
		return err
	}

	// 添加xl/_rels/workbook.xml.rels
	workbookRelsWriter, err := zipWriter.Create("xl/_rels/workbook.xml.rels")
	if err != nil {
		return err
	}

	// 创建工作簿关系
	wbRels := NewRelationships()

	// 添加样式关系
	wbRels.AddRelationship("rId1", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles", "styles.xml")

	// 添加主题关系
	wbRels.AddRelationship("rId2", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme", "theme/theme1.xml")

	// 添加共享字符串表关系
	wbRels.AddRelationship("rId3", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings", "sharedStrings.xml")

	// 添加工作表关系
	for i := range wb.Worksheets {
		relID := fmt.Sprintf("rId%d", i+4) // 从rId4开始
		target := fmt.Sprintf("worksheets/sheet%d.xml", i+1)
		wbRels.AddRelationship(relID, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet", target)
	}

	_, err = workbookRelsWriter.Write([]byte(wbRels.ToXML()))
	if err != nil {
		return err
	}

	// 添加xl/worksheets/sheet1.xml, sheet2.xml, ...
	for i, ws := range wb.Worksheets {
		sheetPath := fmt.Sprintf("xl/worksheets/sheet%d.xml", i+1)
		sheetWriter, err := zipWriter.Create(sheetPath)
		if err != nil {
			return err
		}
		_, err = sheetWriter.Write([]byte(ws.ToXML(wb.SharedStrings)))
		if err != nil {
			return err
		}
	}

	// 添加xl/styles.xml
	stylesWriter, err := zipWriter.Create("xl/styles.xml")
	if err != nil {
		return err
	}
	_, err = stylesWriter.Write([]byte(wb.Styles.ToXML()))
	if err != nil {
		return err
	}

	// 添加xl/theme/theme1.xml
	themeWriter, err := zipWriter.Create("xl/theme/theme1.xml")
	if err != nil {
		return err
	}
	_, err = themeWriter.Write([]byte(wb.Theme.ToXML()))
	if err != nil {
		return err
	}

	// 添加xl/sharedStrings.xml
	sharedStringsWriter, err := zipWriter.Create("xl/sharedStrings.xml")
	if err != nil {
		return err
	}
	_, err = sharedStringsWriter.Write([]byte(wb.SharedStrings.ToXML()))
	if err != nil {
		return err
	}

	return nil
}

// WorkbookRels 表示工作簿的关系
type WorkbookRels struct {
	Relationships *Relationships
}

// NewWorkbookRels 创建一个新的工作簿关系
func NewWorkbookRels() *WorkbookRels {
	return &WorkbookRels{
		Relationships: NewRelationships(),
	}
}

// ToXML 将工作簿转换为XML
func (wb *Workbook) ToXML() string {
	xml := "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
	xml += "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\n"

	// 工作簿属性
	xml += "  <workbookPr defaultThemeVersion=\"124226\"/>\n"

	// 工作表
	xml += "  <sheets>\n"
	for i, ws := range wb.Worksheets {
		relID := fmt.Sprintf("rId%d", i+4) // 从rId4开始，与上面的关系ID对应
		xml += fmt.Sprintf("    <sheet name=\"%s\" sheetId=\"%d\" r:id=\"%s\"/>\n", ws.Name, ws.SheetID, relID)
	}
	xml += "  </sheets>\n"

	xml += "</workbook>"
	return xml
}
