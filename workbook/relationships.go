package workbook

// Relationships 表示Excel文档中的关系集合
type Relationships struct {
	Relationships []*Relationship
}

// Relationship 表示Excel文档中的关系
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

// ToXML 将关系转换为XML
func (r *Relationships) ToXML() string {
	xml := "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
	xml += "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n"

	for _, rel := range r.Relationships {
		if rel.TargetMode == "External" {
			xml += "  <Relationship Id=\"" + rel.ID + "\" Type=\"" + rel.Type + "\" Target=\"" + rel.Target + "\" TargetMode=\"External\"/>\n"
		} else {
			xml += "  <Relationship Id=\"" + rel.ID + "\" Type=\"" + rel.Type + "\" Target=\"" + rel.Target + "\"/>\n"
		}
	}

	xml += "</Relationships>"
	return xml
}
