package workbook

// Theme 表示Excel文档中的主题
type Theme struct {
}

// NewTheme 创建一个新的主题
func NewTheme() *Theme {
	return &Theme{}
}

// ToXML 将主题转换为XML
func (t *Theme) ToXML() string {
	// 使用Office默认主题
	xml := "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
	xml += "<a:theme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" name=\"Office Theme\">\n"

	// 颜色方案
	xml += "  <a:themeElements>\n"
	xml += "    <a:clrScheme name=\"Office\">\n"
	xml += "      <a:dk1><a:sysClr val=\"windowText\" lastClr=\"000000\"/></a:dk1>\n"
	xml += "      <a:lt1><a:sysClr val=\"window\" lastClr=\"FFFFFF\"/></a:lt1>\n"
	xml += "      <a:dk2><a:srgbClr val=\"1F497D\"/></a:dk2>\n"
	xml += "      <a:lt2><a:srgbClr val=\"EEECE1\"/></a:lt2>\n"
	xml += "      <a:accent1><a:srgbClr val=\"4F81BD\"/></a:accent1>\n"
	xml += "      <a:accent2><a:srgbClr val=\"C0504D\"/></a:accent2>\n"
	xml += "      <a:accent3><a:srgbClr val=\"9BBB59\"/></a:accent3>\n"
	xml += "      <a:accent4><a:srgbClr val=\"8064A2\"/></a:accent4>\n"
	xml += "      <a:accent5><a:srgbClr val=\"4BACC6\"/></a:accent5>\n"
	xml += "      <a:accent6><a:srgbClr val=\"F79646\"/></a:accent6>\n"
	xml += "      <a:hlink><a:srgbClr val=\"0000FF\"/></a:hlink>\n"
	xml += "      <a:folHlink><a:srgbClr val=\"800080\"/></a:folHlink>\n"
	xml += "    </a:clrScheme>\n"

	// 字体方案
	xml += "    <a:fontScheme name=\"Office\">\n"
	xml += "      <a:majorFont>\n"
	xml += "        <a:latin typeface=\"Calibri\"/>\n"
	xml += "        <a:ea typeface=\"\"/>\n"
	xml += "        <a:cs typeface=\"\"/>\n"
	xml += "      </a:majorFont>\n"
	xml += "      <a:minorFont>\n"
	xml += "        <a:latin typeface=\"Calibri\"/>\n"
	xml += "        <a:ea typeface=\"\"/>\n"
	xml += "        <a:cs typeface=\"\"/>\n"
	xml += "      </a:minorFont>\n"
	xml += "    </a:fontScheme>\n"

	// 格式方案
	xml += "    <a:fmtScheme name=\"Office\">\n"
	xml += "      <a:fillStyleLst>\n"
	xml += "        <a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill>\n"
	xml += "        <a:gradFill rotWithShape=\"1\">\n"
	xml += "          <a:gsLst>\n"
	xml += "            <a:gs pos=\"0\"><a:schemeClr val=\"phClr\"><a:tint val=\"50000\"/><a:satMod val=\"300000\"/></a:schemeClr></a:gs>\n"
	xml += "            <a:gs pos=\"35000\"><a:schemeClr val=\"phClr\"><a:tint val=\"37000\"/><a:satMod val=\"300000\"/></a:schemeClr></a:gs>\n"
	xml += "            <a:gs pos=\"100000\"><a:schemeClr val=\"phClr\"><a:tint val=\"15000\"/><a:satMod val=\"350000\"/></a:schemeClr></a:gs>\n"
	xml += "          </a:gsLst>\n"
	xml += "          <a:lin ang=\"16200000\" scaled=\"1\"/>\n"
	xml += "        </a:gradFill>\n"
	xml += "        <a:gradFill rotWithShape=\"1\">\n"
	xml += "          <a:gsLst>\n"
	xml += "            <a:gs pos=\"0\"><a:schemeClr val=\"phClr\"><a:shade val=\"51000\"/><a:satMod val=\"130000\"/></a:schemeClr></a:gs>\n"
	xml += "            <a:gs pos=\"80000\"><a:schemeClr val=\"phClr\"><a:shade val=\"93000\"/><a:satMod val=\"130000\"/></a:schemeClr></a:gs>\n"
	xml += "            <a:gs pos=\"100000\"><a:schemeClr val=\"phClr\"><a:shade val=\"94000\"/><a:satMod val=\"135000\"/></a:schemeClr></a:gs>\n"
	xml += "          </a:gsLst>\n"
	xml += "          <a:lin ang=\"16200000\" scaled=\"0\"/>\n"
	xml += "        </a:gradFill>\n"
	xml += "      </a:fillStyleLst>\n"
	xml += "      <a:lnStyleLst>\n"
	xml += "        <a:ln w=\"9525\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"phClr\"><a:shade val=\"95000\"/><a:satMod val=\"105000\"/></a:schemeClr></a:solidFill><a:prstDash val=\"solid\"/></a:ln>\n"
	xml += "        <a:ln w=\"25400\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:prstDash val=\"solid\"/></a:ln>\n"
	xml += "        <a:ln w=\"38100\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:prstDash val=\"solid\"/></a:ln>\n"
	xml += "      </a:lnStyleLst>\n"
	xml += "      <a:effectStyleLst>\n"
	xml += "        <a:effectStyle><a:effectLst/></a:effectStyle>\n"
	xml += "        <a:effectStyle><a:effectLst/></a:effectStyle>\n"
	xml += "        <a:effectStyle><a:effectLst/></a:effectStyle>\n"
	xml += "      </a:effectStyleLst>\n"
	xml += "      <a:bgFillStyleLst>\n"
	xml += "        <a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill>\n"
	xml += "        <a:gradFill rotWithShape=\"1\">\n"
	xml += "          <a:gsLst>\n"
	xml += "            <a:gs pos=\"0\"><a:schemeClr val=\"phClr\"><a:tint val=\"40000\"/><a:satMod val=\"350000\"/></a:schemeClr></a:gs>\n"
	xml += "            <a:gs pos=\"40000\"><a:schemeClr val=\"phClr\"><a:tint val=\"45000\"/><a:shade val=\"99000\"/><a:satMod val=\"350000\"/></a:schemeClr></a:gs>\n"
	xml += "            <a:gs pos=\"100000\"><a:schemeClr val=\"phClr\"><a:shade val=\"20000\"/><a:satMod val=\"255000\"/></a:schemeClr></a:gs>\n"
	xml += "          </a:gsLst>\n"
	xml += "          <a:path path=\"circle\"><a:fillToRect l=\"50000\" t=\"50000\" r=\"50000\" b=\"50000\"/></a:path>\n"
	xml += "        </a:gradFill>\n"
	xml += "        <a:gradFill rotWithShape=\"1\">\n"
	xml += "          <a:gsLst>\n"
	xml += "            <a:gs pos=\"0\"><a:schemeClr val=\"phClr\"><a:tint val=\"80000\"/><a:satMod val=\"300000\"/></a:schemeClr></a:gs>\n"
	xml += "            <a:gs pos=\"100000\"><a:schemeClr val=\"phClr\"><a:shade val=\"30000\"/><a:satMod val=\"200000\"/></a:schemeClr></a:gs>\n"
	xml += "          </a:gsLst>\n"
	xml += "          <a:path path=\"circle\"><a:fillToRect l=\"50000\" t=\"50000\" r=\"50000\" b=\"50000\"/></a:path>\n"
	xml += "        </a:gradFill>\n"
	xml += "      </a:bgFillStyleLst>\n"
	xml += "    </a:fmtScheme>\n"
	xml += "  </a:themeElements>\n"
	xml += "  <a:objectDefaults/>\n"
	xml += "  <a:extraClrSchemeLst/>\n"
	xml += "</a:theme>\n"

	return xml
}
