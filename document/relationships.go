package document

// Relationships 表示Word文档中的关系集合
type Relationships struct {
	Relationships []*Relationship
}

// Relationship 表示Word文档中的关系
type Relationship struct {
	ID         string
	Type       string
	Target     string
	TargetMode string // 目标模式：Internal, External
}

// NewRelationships 创建一个新的关系集合
func NewRelationships() *Relationships {
	return &Relationships{
		Relationships: make([]*Relationship, 0),
	}
}

// AddRelationship 添加一个关系
func (r *Relationships) AddRelationship(id, relType, target string) *Relationship {
	rel := &Relationship{
		ID:     id,
		Type:   relType,
		Target: target,
	}
	r.Relationships = append(r.Relationships, rel)
	return rel
}

// AddExternalRelationship 添加一个外部关系
func (r *Relationships) AddExternalRelationship(id, relType, target string) *Relationship {
	rel := &Relationship{
		ID:         id,
		Type:       relType,
		Target:     target,
		TargetMode: "External",
	}
	r.Relationships = append(r.Relationships, rel)
	return rel
}

// GetRelationshipByID 根据ID获取关系
func (r *Relationships) GetRelationshipByID(id string) *Relationship {
	for _, rel := range r.Relationships {
		if rel.ID == id {
			return rel
		}
	}
	return nil
}

// GetRelationshipsByType 根据类型获取关系
func (r *Relationships) GetRelationshipsByType(relType string) []*Relationship {
	result := make([]*Relationship, 0)
	for _, rel := range r.Relationships {
		if rel.Type == relType {
			result = append(result, rel)
		}
	}
	return result
}

// ToXML 将关系集合转换为XML
func (r *Relationships) ToXML() string {
	xml := "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
	xml += "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"

	for _, rel := range r.Relationships {
		xml += "<Relationship Id=\"" + rel.ID + "\""
		xml += " Type=\"" + rel.Type + "\""
		xml += " Target=\"" + rel.Target + "\""
		if rel.TargetMode != "" {
			xml += " TargetMode=\"" + rel.TargetMode + "\""
		}
		xml += " />"
	}

	xml += "</Relationships>"
	return xml
}

// DocumentRels 表示Word文档中的文档关系
type DocumentRels struct {
	Relationships *Relationships
}

// NewDocumentRels 创建一个新的文档关系
func NewDocumentRels() *DocumentRels {
	return &DocumentRels{
		Relationships: NewRelationships(),
	}
}

// AddImage 添加一个图片关系
func (d *DocumentRels) AddImage(id, target string) *Relationship {
	return d.Relationships.AddRelationship(id, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image", target)
}

// AddHyperlink 添加一个超链接关系
func (d *DocumentRels) AddHyperlink(id, target string) *Relationship {
	return d.Relationships.AddExternalRelationship(id, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink", target)
}

// AddHeader 添加一个页眉关系
func (d *DocumentRels) AddHeader(id, target string) *Relationship {
	return d.Relationships.AddRelationship(id, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header", target)
}

// AddFooter 添加一个页脚关系
func (d *DocumentRels) AddFooter(id, target string) *Relationship {
	return d.Relationships.AddRelationship(id, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer", target)
}

// ToXML 将文档关系转换为XML
func (d *DocumentRels) ToXML() string {
	return d.Relationships.ToXML()
}
