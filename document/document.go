package document

import (
	"archive/zip"
	"fmt"
	"os"
	"path/filepath"
	"time"
)

// Document 表示一个Word文档
type Document struct {
	Body          *Body
	Properties    *DocumentProperties
	Relationships *Relationships
	Styles        *Styles
	Numbering     *Numbering
	Footers       []*Footer
	Headers       []*Header
	Theme         *Theme
	Settings      *Settings
	ContentTypes  *ContentTypes
	Rels          *DocumentRels
}

// DocumentProperties 包含文档的元数据
type DocumentProperties struct {
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

// NewDocument 创建一个新的Word文档
func NewDocument() *Document {
	return &Document{
		Body: NewBody(),
		Properties: &DocumentProperties{
			Created:  time.Now(),
			Modified: time.Now(),
			Revision: 1,
		},
		Relationships: NewRelationships(),
		Styles:        NewStyles(),
		Numbering:     NewNumbering(),
		Footers:       make([]*Footer, 0),
		Headers:       make([]*Header, 0),
		Theme:         NewTheme(),
		Settings:      NewSettings(),
		ContentTypes:  NewContentTypes(),
		Rels:          NewDocumentRels(),
	}
}

// Save 将文档保存到指定路径
func (d *Document) Save(path string) error {
	// 创建一个新的zip文件
	zipFile, err := os.Create(path)
	if err != nil {
		return err
	}
	defer zipFile.Close()

	// 创建一个zip writer
	zipWriter := zip.NewWriter(zipFile)
	defer zipWriter.Close()

	// 添加[Content_Types].xml
	if err := d.addContentTypes(zipWriter); err != nil {
		return err
	}

	// 添加_rels/.rels
	if err := d.addRels(zipWriter); err != nil {
		return err
	}

	// 添加docProps/app.xml和docProps/core.xml
	if err := d.addDocProps(zipWriter); err != nil {
		return err
	}

	// 添加word/document.xml
	if err := d.addDocument(zipWriter); err != nil {
		return err
	}

	// 添加word/styles.xml
	if err := d.addStyles(zipWriter); err != nil {
		return err
	}

	// 添加word/numbering.xml
	if err := d.addNumbering(zipWriter); err != nil {
		return err
	}

	// 添加word/_rels/document.xml.rels
	if err := d.addDocumentRels(zipWriter); err != nil {
		return err
	}

	// 添加word/theme/theme1.xml
	if err := d.addTheme(zipWriter); err != nil {
		return err
	}

	// 添加word/settings.xml
	if err := d.addSettings(zipWriter); err != nil {
		return err
	}

	// 添加页眉
	for i, header := range d.Headers {
		if err := d.addHeader(zipWriter, header, i+1); err != nil {
			return err
		}
	}

	// 添加页脚
	for i, footer := range d.Footers {
		if err := d.addFooter(zipWriter, footer, i+1); err != nil {
			return err
		}
	}

	// 添加图片
	for _, rel := range d.Rels.Relationships.GetRelationshipsByType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/image") {
		if err := d.addImage(zipWriter, rel); err != nil {
			return err
		}
	}

	return nil
}

// 以下是内部方法，用于将各个部分添加到zip文件中
func (d *Document) addContentTypes(zipWriter *zip.Writer) error {
	// 添加[Content_Types].xml
	w, err := zipWriter.Create("[Content_Types].xml")
	if err != nil {
		return err
	}

	_, err = w.Write([]byte(d.ContentTypes.ToXML()))
	return err
}

func (d *Document) addRels(zipWriter *zip.Writer) error {
	// 添加_rels/.rels
	w, err := zipWriter.Create("_rels/.rels")
	if err != nil {
		return err
	}

	// 创建基本关系
	rels := NewRelationships()
	rels.AddRelationship("rId1", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument", "document/document.xml")
	rels.AddRelationship("rId2", "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties", "docProps/core.xml")
	rels.AddRelationship("rId3", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties", "docProps/app.xml")

	_, err = w.Write([]byte(rels.ToXML()))
	return err
}

func (d *Document) addDocProps(zipWriter *zip.Writer) error {
	// 添加docProps/app.xml
	w1, err := zipWriter.Create("docProps/app.xml")
	if err != nil {
		return err
	}

	// 创建app.xml内容
	appXML := "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
	appXML += "<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\" xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\">\n"
	appXML += "<Application>go-flowdoc</Application>\n"
	appXML += "<AppVersion>1.0.0</AppVersion>\n"
	appXML += "</Properties>\n"

	_, err = w1.Write([]byte(appXML))
	if err != nil {
		return err
	}

	// 添加docProps/core.xml
	w2, err := zipWriter.Create("docProps/core.xml")
	if err != nil {
		return err
	}

	// 创建core.xml内容
	coreXML := "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
	coreXML += "<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" "
	coreXML += "xmlns:dc=\"http://purl.org/dc/elements/1.1/\" "
	coreXML += "xmlns:dcterms=\"http://purl.org/dc/terms/\" "
	coreXML += "xmlns:dcmitype=\"http://purl.org/dc/dcmitype/\" "
	coreXML += "xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">\n"

	if d.Properties.Title != "" {
		coreXML += "<dc:title>" + d.Properties.Title + "</dc:title>\n"
	}

	if d.Properties.Subject != "" {
		coreXML += "<dc:subject>" + d.Properties.Subject + "</dc:subject>\n"
	}

	if d.Properties.Creator != "" {
		coreXML += "<dc:creator>" + d.Properties.Creator + "</dc:creator>\n"
	}

	if d.Properties.Keywords != "" {
		coreXML += "<cp:keywords>" + d.Properties.Keywords + "</cp:keywords>\n"
	}

	if d.Properties.Description != "" {
		coreXML += "<dc:description>" + d.Properties.Description + "</dc:description>\n"
	}

	if d.Properties.LastModifiedBy != "" {
		coreXML += "<cp:lastModifiedBy>" + d.Properties.LastModifiedBy + "</cp:lastModifiedBy>\n"
	}

	if d.Properties.Revision > 0 {
		coreXML += "<cp:revision>" + fmt.Sprintf("%d", d.Properties.Revision) + "</cp:revision>\n"
	}

	// 格式化时间
	createdTime := d.Properties.Created.Format("2006-01-02T15:04:05Z")
	modifiedTime := d.Properties.Modified.Format("2006-01-02T15:04:05Z")

	coreXML += "<dcterms:created xsi:type=\"dcterms:W3CDTF\">" + createdTime + "</dcterms:created>\n"
	coreXML += "<dcterms:modified xsi:type=\"dcterms:W3CDTF\">" + modifiedTime + "</dcterms:modified>\n"

	coreXML += "</cp:coreProperties>\n"

	_, err = w2.Write([]byte(coreXML))
	return err
}

func (d *Document) addDocument(zipWriter *zip.Writer) error {
	// 添加word/document.xml
	w, err := zipWriter.Create("document/document.xml")
	if err != nil {
		return err
	}

	// 创建document.xml内容
	docXML := "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
	docXML += "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" "
	docXML += "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" "
	docXML += "xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\">\n"

	// 添加文档主体
	docXML += d.Body.ToXML()

	docXML += "</w:document>"

	_, err = w.Write([]byte(docXML))
	return err
}

func (d *Document) addStyles(zipWriter *zip.Writer) error {
	// 添加word/styles.xml
	w, err := zipWriter.Create("document/styles.xml")
	if err != nil {
		return err
	}

	_, err = w.Write([]byte(d.Styles.ToXML()))
	return err
}

func (d *Document) addNumbering(zipWriter *zip.Writer) error {
	// 添加word/numbering.xml
	w, err := zipWriter.Create("document/numbering.xml")
	if err != nil {
		return err
	}

	_, err = w.Write([]byte(d.Numbering.ToXML()))
	return err
}

func (d *Document) addDocumentRels(zipWriter *zip.Writer) error {
	// 添加word/_rels/document.xml.rels
	w, err := zipWriter.Create("document/_rels/document.xml.rels")
	if err != nil {
		return err
	}

	_, err = w.Write([]byte(d.Rels.ToXML()))
	return err
}

func (d *Document) addTheme(zipWriter *zip.Writer) error {
	// 添加word/theme/theme1.xml
	w, err := zipWriter.Create("document/theme/theme1.xml")
	if err != nil {
		return err
	}

	_, err = w.Write([]byte(d.Theme.ToXML()))
	return err
}

func (d *Document) addSettings(zipWriter *zip.Writer) error {
	// 添加word/settings.xml
	w, err := zipWriter.Create("document/settings.xml")
	if err != nil {
		return err
	}

	_, err = w.Write([]byte(d.Settings.ToXML()))
	return err
}

func (d *Document) addHeader(zipWriter *zip.Writer, header *Header, index int) error {
	// 添加word/header{index}.xml
	headerPath := fmt.Sprintf("document/header%d.xml", index)
	w, err := zipWriter.Create(headerPath)
	if err != nil {
		return err
	}

	_, err = w.Write([]byte(header.ToXML()))
	return err
}

func (d *Document) addFooter(zipWriter *zip.Writer, footer *Footer, index int) error {
	// 添加word/footer{index}.xml
	footerPath := fmt.Sprintf("document/footer%d.xml", index)
	w, err := zipWriter.Create(footerPath)
	if err != nil {
		return err
	}

	_, err = w.Write([]byte(footer.ToXML()))
	return err
}

func (d *Document) addImage(zipWriter *zip.Writer, rel *Relationship) error {
	// 从关系中提取图片ID和路径
	imageID := rel.ID
	imagePath := rel.Target

	// 查找对应的Drawing对象
	var imageData []byte

	// 在文档主体中查找
	for _, para := range d.Body.Content {
		if p, ok := para.(*Paragraph); ok {
			for _, run := range p.Runs {
				if run.Drawing != nil && run.Drawing.ID == imageID {
					imageData = run.Drawing.ImageData
					break
				}
			}
		}
	}

	// 在页眉中查找
	if len(imageData) == 0 {
		for _, header := range d.Headers {
			for _, content := range header.Content {
				if p, ok := content.(*Paragraph); ok {
					for _, run := range p.Runs {
						if run.Drawing != nil && run.Drawing.ID == imageID {
							imageData = run.Drawing.ImageData
							break
						}
					}
				}
			}
		}
	}

	// 在页脚中查找
	if len(imageData) == 0 {
		for _, footer := range d.Footers {
			for _, content := range footer.Content {
				if p, ok := content.(*Paragraph); ok {
					for _, run := range p.Runs {
						if run.Drawing != nil && run.Drawing.ID == imageID {
							imageData = run.Drawing.ImageData
							break
						}
					}
				}
			}
		}
	}

	if len(imageData) == 0 {
		return fmt.Errorf("未找到图片数据: %s", imageID)
	}

	// 添加图片文件
	w, err := zipWriter.Create("document/" + imagePath)
	if err != nil {
		return err
	}

	_, err = w.Write(imageData)
	return err
}

// AddParagraph 向文档添加一个段落
func (d *Document) AddParagraph() *Paragraph {
	return d.Body.AddParagraph()
}

// AddTable 向文档添加一个表格
func (d *Document) AddTable(rows, cols int) *Table {
	return d.Body.AddTable(rows, cols)
}

// AddPageBreak 向文档添加一个分页符
func (d *Document) AddPageBreak() *Paragraph {
	return d.Body.AddPageBreak()
}

// AddSectionBreak 向文档添加一个分节符
func (d *Document) AddSectionBreak() *Paragraph {
	return d.Body.AddSectionBreak()
}

// AddHeader 向文档添加一个页眉并返回它
func (d *Document) AddHeader() *Header {
	header := NewHeader()
	d.Headers = append(d.Headers, header)

	// 添加页眉关系
	headerID := fmt.Sprintf("rId%d", len(d.Rels.Relationships.Relationships)+1)
	headerPath := fmt.Sprintf("header%d.xml", len(d.Headers))
	d.Rels.AddHeader(headerID, headerPath)

	// 添加页眉内容类型
	d.ContentTypes.AddHeaderOverride(len(d.Headers))

	return header
}

// AddFooter 向文档添加一个页脚并返回它
func (d *Document) AddFooter() *Footer {
	footer := NewFooter()
	d.Footers = append(d.Footers, footer)

	// 添加页脚关系
	footerID := fmt.Sprintf("rId%d", len(d.Rels.Relationships.Relationships)+1)
	footerPath := fmt.Sprintf("footer%d.xml", len(d.Footers))
	d.Rels.AddFooter(footerID, footerPath)

	// 添加页脚内容类型
	d.ContentTypes.AddFooterOverride(len(d.Footers))

	return footer
}

// AddImage 向文档添加一个图片
func (d *Document) AddImage(path string, width, height int) (*Run, error) {
	// 创建一个新段落和运行
	para := d.AddParagraph()
	run := para.AddRun()

	// 创建图片
	drawing := NewDrawing()
	err := drawing.SetImagePath(path)
	if err != nil {
		return nil, err
	}

	// 设置图片大小
	drawing.SetSize(width, height)

	// 添加图片关系
	imageID := fmt.Sprintf("rId%d", len(d.Rels.Relationships.Relationships)+1)
	imageName := filepath.Base(path)
	imagePath := fmt.Sprintf("media/%s", imageName)
	d.Rels.AddImage(imageID, imagePath)

	// 设置图片ID
	drawing.ID = imageID

	// 添加图片到运行
	run.AddDrawing(drawing)

	return run, nil
}

// AddImageBytes 通过字节数据添加图片
func (d *Document) AddImageBytes(data []byte, format, name string, width, height int) (*Run, error) {
	// 创建一个新段落和运行
	para := d.AddParagraph()
	run := para.AddRun()

	// 创建图片
	drawing := NewDrawing()
	drawing.SetImageData(data)
	drawing.SetName(name)

	// 设置图片大小
	drawing.SetSize(width, height)

	// 添加图片关系
	imageID := fmt.Sprintf("rId%d", len(d.Rels.Relationships.Relationships)+1)
	imagePath := fmt.Sprintf("media/%s.%s", name, format)
	d.Rels.AddImage(imageID, imagePath)

	// 设置图片ID
	drawing.ID = imageID

	// 添加图片到运行
	run.AddDrawing(drawing)

	return run, nil
}

// SetTitle 设置文档标题
func (d *Document) SetTitle(title string) *Document {
	d.Properties.Title = title
	return d
}

// SetSubject 设置文档主题
func (d *Document) SetSubject(subject string) *Document {
	d.Properties.Subject = subject
	return d
}

// SetCreator 设置文档创建者
func (d *Document) SetCreator(creator string) *Document {
	d.Properties.Creator = creator
	d.Properties.LastModifiedBy = creator
	return d
}

// SetKeywords 设置文档关键词
func (d *Document) SetKeywords(keywords string) *Document {
	d.Properties.Keywords = keywords
	return d
}

// SetDescription 设置文档描述
func (d *Document) SetDescription(description string) *Document {
	d.Properties.Description = description
	return d
}

// SetLastModifiedBy 设置文档最后修改者
func (d *Document) SetLastModifiedBy(lastModifiedBy string) *Document {
	d.Properties.LastModifiedBy = lastModifiedBy
	return d
}

// SetRevision 设置文档修订版本
func (d *Document) SetRevision(revision int) *Document {
	d.Properties.Revision = revision
	return d
}

// SetCreated 设置文档创建时间
func (d *Document) SetCreated(created time.Time) *Document {
	d.Properties.Created = created
	return d
}

// SetModified 设置文档修改时间
func (d *Document) SetModified(modified time.Time) *Document {
	d.Properties.Modified = modified
	return d
}

// SetPageSize 设置页面大小
func (d *Document) SetPageSize(width, height int, orientation string) *Document {
	d.Body.SectionProperties.PageSize.Width = width
	d.Body.SectionProperties.PageSize.Height = height
	d.Body.SectionProperties.PageSize.Orientation = orientation
	return d
}

// SetPageSizeA4 设置页面大小为A4
func (d *Document) SetPageSizeA4(landscape bool) *Document {
	if landscape {
		return d.SetPageSize(16838, 11906, "landscape")
	}
	return d.SetPageSize(11906, 16838, "portrait")
}

// SetPageSizeA5 设置页面大小为A5
func (d *Document) SetPageSizeA5(landscape bool) *Document {
	if landscape {
		return d.SetPageSize(11906, 8419, "landscape")
	}
	return d.SetPageSize(8419, 11906, "portrait")
}

// SetPageSizeLetter 设置页面大小为Letter
func (d *Document) SetPageSizeLetter(landscape bool) *Document {
	if landscape {
		return d.SetPageSize(15840, 12240, "landscape")
	}
	return d.SetPageSize(12240, 15840, "portrait")
}

// SetPageMargin 设置页面边距
func (d *Document) SetPageMargin(top, right, bottom, left, header, footer, gutter int) *Document {
	d.Body.SectionProperties.PageMargin.Top = top
	d.Body.SectionProperties.PageMargin.Right = right
	d.Body.SectionProperties.PageMargin.Bottom = bottom
	d.Body.SectionProperties.PageMargin.Left = left
	d.Body.SectionProperties.PageMargin.Header = header
	d.Body.SectionProperties.PageMargin.Footer = footer
	d.Body.SectionProperties.PageMargin.Gutter = gutter
	return d
}

// SetColumns 设置分栏
func (d *Document) SetColumns(num, space int) *Document {
	d.Body.SectionProperties.Columns.Num = num
	d.Body.SectionProperties.Columns.Space = space
	return d
}

// AddHeaderReference 添加页眉引用
func (d *Document) AddHeaderReference(headerType, id string) *Document {
	headerRef := &HeaderFooterReference{
		Type: headerType,
		ID:   id,
	}
	d.Body.SectionProperties.HeaderReference = append(d.Body.SectionProperties.HeaderReference, headerRef)
	return d
}

// AddFooterReference 添加页脚引用
func (d *Document) AddFooterReference(footerType, id string) *Document {
	footerRef := &HeaderFooterReference{
		Type: footerType,
		ID:   id,
	}
	d.Body.SectionProperties.FooterReference = append(d.Body.SectionProperties.FooterReference, footerRef)
	return d
}

// AddHeaderWithReference 添加页眉并同时添加页眉引用
func (d *Document) AddHeaderWithReference(headerType string) *Header {
	header := NewHeader()
	d.Headers = append(d.Headers, header)

	// 添加页眉关系
	headerID := fmt.Sprintf("rId%d", len(d.Rels.Relationships.Relationships)+1)
	headerPath := fmt.Sprintf("header%d.xml", len(d.Headers))
	d.Rels.AddHeader(headerID, headerPath)

	// 添加页眉内容类型
	d.ContentTypes.AddHeaderOverride(len(d.Headers))

	// 添加页眉引用
	d.AddHeaderReference(headerType, headerID)

	return header
}

// AddFooterWithReference 添加页脚并同时添加页脚引用
func (d *Document) AddFooterWithReference(footerType string) *Footer {
	footer := NewFooter()
	d.Footers = append(d.Footers, footer)

	// 添加页脚关系
	footerID := fmt.Sprintf("rId%d", len(d.Rels.Relationships.Relationships)+1)
	footerPath := fmt.Sprintf("footer%d.xml", len(d.Footers))
	d.Rels.AddFooter(footerID, footerPath)

	// 添加页脚内容类型
	d.ContentTypes.AddFooterOverride(len(d.Footers))

	// 添加页脚引用
	d.AddFooterReference(footerType, footerID)

	return footer
}

// AddPageNumberParagraph 添加一个居中的页码段落
func (d *Document) AddPageNumberParagraph() *Paragraph {
	// 创建一个新段落
	para := d.AddParagraph()
	para.SetAlignment("center")

	// 添加"第"文本
	para.AddRun().AddText("第 ")

	// 创建页码域的所有部分
	// 1. 域开始
	fieldBegin := para.AddRun()
	fieldBegin.Field = &Field{
		Type: "begin",
		Code: "PAGE",
	}

	// 2. 域分隔符
	fieldSeparate := para.AddRun()
	fieldSeparate.Field = &Field{
		Type: "separate",
	}

	// 3. 页码内容（在Word中会替换为实际页码）
	fieldContent := para.AddRun()
	fieldContent.AddText("1")

	// 4. 域结束
	fieldEnd := para.AddRun()
	fieldEnd.Field = &Field{
		Type: "end",
	}

	// 添加"页"文本
	para.AddRun().AddText(" 页")

	return para
}
