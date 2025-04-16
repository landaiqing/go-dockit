package document

import (
	"encoding/base64"
	"fmt"
	"os"
	"path/filepath"
	"strings"
)

// Drawing 表示Word文档中的图形
type Drawing struct {
	ID          string
	Name        string
	Description string
	ImagePath   string
	ImageData   []byte
	Width       int    // 单位为EMU (English Metric Unit)
	Height      int    // 单位为EMU (1厘米 = 360000 EMU)
	WrapType    string // 文字环绕方式：inline, square, tight, through, topAndBottom, behind, inFront
	PositionH   *DrawingPosition
	PositionV   *DrawingPosition
}

// DrawingPosition 表示图形的位置
type DrawingPosition struct {
	RelativeFrom string // 相对位置：page, margin, column, paragraph, line, character
	Align        string // 对齐方式：left, center, right, inside, outside
	Offset       int    // 偏移量，单位为EMU
}

// NewDrawing 创建一个新的图形
func NewDrawing() *Drawing {
	return &Drawing{
		ID:       generateUniqueID(),
		WrapType: "inline",
	}
}

// SetImagePath 设置图片路径
func (d *Drawing) SetImagePath(path string) *Drawing {
	d.ImagePath = path

	// 设置图片名称
	if d.Name == "" {
		d.Name = filepath.Base(path)
	}

	// 读取图片数据
	data, err := os.ReadFile(path)
	if err == nil {
		d.ImageData = data
	}

	return d
}

// SetImageData 设置图片数据
func (d *Drawing) SetImageData(data []byte) *Drawing {
	d.ImageData = data
	return d
}

// SetSize 设置图片大小
func (d *Drawing) SetSize(width, height int) *Drawing {
	d.Width = width
	d.Height = height
	return d
}

// SetName 设置图片名称
func (d *Drawing) SetName(name string) *Drawing {
	d.Name = name
	return d
}

// SetDescription 设置图片描述
func (d *Drawing) SetDescription(description string) *Drawing {
	d.Description = description
	return d
}

// SetWrapType 设置文字环绕方式
func (d *Drawing) SetWrapType(wrapType string) *Drawing {
	d.WrapType = wrapType
	return d
}

// SetPositionH 设置水平位置
func (d *Drawing) SetPositionH(relativeFrom, align string, offset int) *Drawing {
	d.PositionH = &DrawingPosition{
		RelativeFrom: relativeFrom,
		Align:        align,
		Offset:       offset,
	}
	return d
}

// SetPositionV 设置垂直位置
func (d *Drawing) SetPositionV(relativeFrom, align string, offset int) *Drawing {
	d.PositionV = &DrawingPosition{
		RelativeFrom: relativeFrom,
		Align:        align,
		Offset:       offset,
	}
	return d
}

// ToXML 将图形转换为XML
func (d *Drawing) ToXML() string {
	xml := "<w:drawing>"

	// 内联图片
	if d.WrapType == "inline" {
		xml += "<wp:inline distT=\"0\" distB=\"0\" distL=\"0\" distR=\"0\">"

		// 图片大小
		xml += "<wp:extent cx=\"" + fmt.Sprintf("%d", d.Width) + "\" cy=\"" + fmt.Sprintf("%d", d.Height) + "\" />"

		// 图片效果
		xml += "<wp:effectExtent l=\"0\" t=\"0\" r=\"0\" b=\"0\" />"

		// 文档中的图片
		xml += "<wp:docPr id=\"" + d.ID + "\" name=\"" + d.Name + "\" descr=\"" + d.Description + "\" />"

		// 图片属性
		xml += "<wp:cNvGraphicFramePr>"
		xml += "<a:graphicFrameLocks xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" noChangeAspect=\"1\" />"
		xml += "</wp:cNvGraphicFramePr>"

		// 图片
		xml += "<a:graphic xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">"
		xml += "<a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">"
		xml += "<pic:pic xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">"

		// 图片信息
		xml += "<pic:nvPicPr>"
		xml += "<pic:cNvPr id=\"0\" name=\"" + d.Name + "\" descr=\"" + d.Description + "\" />"
		xml += "<pic:cNvPicPr>"
		xml += "<a:picLocks noChangeAspect=\"1\" noChangeArrowheads=\"1\" />"
		xml += "</pic:cNvPicPr>"
		xml += "</pic:nvPicPr>"

		// 图片填充
		xml += "<pic:blipFill>"
		xml += "<a:blip r:embed=\"rId" + d.ID + "\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" />"
		xml += "<a:stretch>"
		xml += "<a:fillRect />"
		xml += "</a:stretch>"
		xml += "</pic:blipFill>"

		// 图片形状
		xml += "<pic:spPr>"
		xml += "<a:xfrm>"
		xml += "<a:off x=\"0\" y=\"0\" />"
		xml += "<a:ext cx=\"" + fmt.Sprintf("%d", d.Width) + "\" cy=\"" + fmt.Sprintf("%d", d.Height) + "\" />"
		xml += "</a:xfrm>"
		xml += "<a:prstGeom prst=\"rect\">"
		xml += "<a:avLst />"
		xml += "</a:prstGeom>"
		xml += "</pic:spPr>"

		xml += "</pic:pic>"
		xml += "</a:graphicData>"
		xml += "</a:graphic>"

		xml += "</wp:inline>"
	} else {
		// 浮动图片
		xml += "<wp:anchor distT=\"0\" distB=\"0\" distL=\"0\" distR=\"0\" simplePos=\"0\" relativeHeight=\"0\" behindDoc=\"" + boolToString(d.WrapType == "behind") + "\" locked=\"0\" layoutInCell=\"1\" allowOverlap=\"1\">"

		// 简单位置
		xml += "<wp:simplePos x=\"0\" y=\"0\" />"

		// 水平位置
		if d.PositionH != nil {
			xml += "<wp:positionH relativeFrom=\"" + d.PositionH.RelativeFrom + "\">"
			if d.PositionH.Align != "" {
				xml += "<wp:align>" + d.PositionH.Align + "</wp:align>"
			} else {
				xml += "<wp:posOffset>" + fmt.Sprintf("%d", d.PositionH.Offset) + "</wp:posOffset>"
			}
			xml += "</wp:positionH>"
		} else {
			xml += "<wp:positionH relativeFrom=\"column\">"
			xml += "<wp:align>left</wp:align>"
			xml += "</wp:positionH>"
		}

		// 垂直位置
		if d.PositionV != nil {
			xml += "<wp:positionV relativeFrom=\"" + d.PositionV.RelativeFrom + "\">"
			if d.PositionV.Align != "" {
				xml += "<wp:align>" + d.PositionV.Align + "</wp:align>"
			} else {
				xml += "<wp:posOffset>" + fmt.Sprintf("%d", d.PositionV.Offset) + "</wp:posOffset>"
			}
			xml += "</wp:positionV>"
		} else {
			xml += "<wp:positionV relativeFrom=\"paragraph\">"
			xml += "<wp:align>top</wp:align>"
			xml += "</wp:positionV>"
		}

		// 图片大小
		xml += "<wp:extent cx=\"" + fmt.Sprintf("%d", d.Width) + "\" cy=\"" + fmt.Sprintf("%d", d.Height) + "\" />"

		// 图片效果
		xml += "<wp:effectExtent l=\"0\" t=\"0\" r=\"0\" b=\"0\" />"

		// 文字环绕方式
		switch d.WrapType {
		case "square":
			xml += "<wp:wrapSquare wrapText=\"bothSides\" />"
		case "tight":
			xml += "<wp:wrapTight wrapText=\"bothSides\" />"
		case "through":
			xml += "<wp:wrapThrough wrapText=\"bothSides\" />"
		case "topAndBottom":
			xml += "<wp:wrapTopAndBottom />"
		case "behind":
			xml += "<wp:wrapNone />"
		case "inFront":
			xml += "<wp:wrapNone />"
		default:
			xml += "<wp:wrapSquare wrapText=\"bothSides\" />"
		}

		// 文档中的图片
		xml += "<wp:docPr id=\"" + d.ID + "\" name=\"" + d.Name + "\" descr=\"" + d.Description + "\" />"

		// 图片属性
		xml += "<wp:cNvGraphicFramePr>"
		xml += "<a:graphicFrameLocks xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" noChangeAspect=\"1\" />"
		xml += "</wp:cNvGraphicFramePr>"

		// 图片
		xml += "<a:graphic xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">"
		xml += "<a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">"
		xml += "<pic:pic xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">"

		// 图片信息
		xml += "<pic:nvPicPr>"
		xml += "<pic:cNvPr id=\"0\" name=\"" + d.Name + "\" descr=\"" + d.Description + "\" />"
		xml += "<pic:cNvPicPr>"
		xml += "<a:picLocks noChangeAspect=\"1\" noChangeArrowheads=\"1\" />"
		xml += "</pic:cNvPicPr>"
		xml += "</pic:nvPicPr>"

		// 图片填充
		xml += "<pic:blipFill>"
		xml += "<a:blip r:embed=\"rId" + d.ID + "\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" />"
		xml += "<a:stretch>"
		xml += "<a:fillRect />"
		xml += "</a:stretch>"
		xml += "</pic:blipFill>"

		// 图片形状
		xml += "<pic:spPr>"
		xml += "<a:xfrm>"
		xml += "<a:off x=\"0\" y=\"0\" />"
		xml += "<a:ext cx=\"" + fmt.Sprintf("%d", d.Width) + "\" cy=\"" + fmt.Sprintf("%d", d.Height) + "\" />"
		xml += "</a:xfrm>"
		xml += "<a:prstGeom prst=\"rect\">"
		xml += "<a:avLst />"
		xml += "</a:prstGeom>"
		xml += "</pic:spPr>"

		xml += "</pic:pic>"
		xml += "</a:graphicData>"
		xml += "</a:graphic>"

		xml += "</wp:anchor>"
	}

	xml += "</w:drawing>"
	return xml
}

// GetImageData 获取图片数据的Base64编码
func (d *Drawing) GetImageData() string {
	return base64.StdEncoding.EncodeToString(d.ImageData)
}

// GetImageType 获取图片类型
func (d *Drawing) GetImageType() string {
	if d.ImagePath != "" {
		ext := strings.ToLower(filepath.Ext(d.ImagePath))
		switch ext {
		case ".jpg", ".jpeg":
			return "jpeg"
		case ".png":
			return "png"
		case ".gif":
			return "gif"
		case ".bmp":
			return "bmp"
		case ".tif", ".tiff":
			return "tiff"
		case ".wmf":
			return "x-wmf"
		case ".emf":
			return "x-emf"
		default:
			return "jpeg"
		}
	}
	return "jpeg"
}

// GetContentType 获取图片的Content-Type
func (d *Drawing) GetContentType() string {
	switch d.GetImageType() {
	case "jpeg":
		return "image/jpeg"
	case "png":
		return "image/png"
	case "gif":
		return "image/gif"
	case "bmp":
		return "image/bmp"
	case "tiff":
		return "image/tiff"
	case "x-wmf":
		return "image/x-wmf"
	case "x-emf":
		return "image/x-emf"
	default:
		return "image/jpeg"
	}
}

// Error 实现error接口
func (d *Drawing) Error() string {
	return fmt.Sprintf("Drawing error: %s", d.Description)
}
