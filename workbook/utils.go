package workbook

import (
	"fmt"
	"strconv"
	"strings"
	"time"
)

// CellRef 根据行列索引生成单元格引用
// 例如: CellRef(0, 0) 返回 "A1"
func CellRef(row, col int) string {
	return fmt.Sprintf("%s%d", ColIndexToName(col), row+1)
}

// ColIndexToName 将列索引转换为列名
// 例如: ColIndexToName(0) 返回 "A", ColIndexToName(25) 返回 "Z", ColIndexToName(26) 返回 "AA"
func ColIndexToName(colIndex int) string {
	if colIndex < 0 {
		return ""
	}

	result := ""
	for colIndex >= 0 {
		remainder := colIndex % 26
		result = string(rune('A'+remainder)) + result
		colIndex = colIndex/26 - 1
		if colIndex < 0 {
			break
		}
	}
	return result
}

// ColNameToIndex 将列名转换为列索引
// 例如: ColNameToIndex("A") 返回 0, ColNameToIndex("Z") 返回 25, ColNameToIndex("AA") 返回 26
func ColNameToIndex(colName string) int {
	colName = strings.ToUpper(colName)
	result := 0
	for i := 0; i < len(colName); i++ {
		result = result*26 + int(colName[i]-'A'+1)
	}
	return result - 1
}

// ParseCellRef 解析单元格引用为行列索引
// 例如: ParseCellRef("A1") 返回 (0, 0)
func ParseCellRef(cellRef string) (row, col int, err error) {
	// 找到字母和数字的分界点
	index := 0
	for index < len(cellRef) && (cellRef[index] < '0' || cellRef[index] > '9') {
		index++
	}

	if index == 0 || index == len(cellRef) {
		return 0, 0, fmt.Errorf("invalid cell reference: %s", cellRef)
	}

	colName := cellRef[:index]
	rowStr := cellRef[index:]

	// 解析行号
	rowNum, err := strconv.Atoi(rowStr)
	if err != nil {
		return 0, 0, fmt.Errorf("invalid row number in cell reference: %s", cellRef)
	}

	// 解析列名
	colIndex := ColNameToIndex(colName)

	return rowNum - 1, colIndex, nil
}

// FormatDate 将时间格式化为Excel日期格式
func FormatDate(t time.Time, format string) string {
	// Excel日期格式映射到Go时间格式
	format = strings.Replace(format, "yyyy", "2006", -1)
	format = strings.Replace(format, "yy", "06", -1)
	format = strings.Replace(format, "mm", "01", -1)
	format = strings.Replace(format, "dd", "02", -1)
	format = strings.Replace(format, "hh", "15", -1)
	format = strings.Replace(format, "ss", "05", -1)

	// 确保分钟格式正确处理
	if strings.Contains(format, "15:01") {
		format = strings.Replace(format, "01", "04", -1)
	}

	return t.Format(format)
}

// GetExcelSerialDate 将时间转换为Excel序列日期值
// Excel日期系统: 1900年1月1日为1，每天加1
func GetExcelSerialDate(t time.Time) float64 {
	// 转换到当地时区，避免时区问题
	t = t.Local()

	// Excel基准日期: 1900年1月0日（但Excel实际以1900年1月1日为1）
	baseDate := time.Date(1899, 12, 30, 0, 0, 0, 0, time.Local)

	// 计算相差的天数（包括小数部分来表示时间）
	days := t.Sub(baseDate).Hours() / 24.0

	// Excel有一个关于1900年2月29日的错误，实际上1900年不是闰年
	// 如果日期在1900年3月1日之后，需要加1来匹配Excel的错误
	if t.After(time.Date(1900, 3, 1, 0, 0, 0, 0, time.Local)) {
		days += 1
	}

	// Excel中1900年1月1日是1而不是0
	return days + 1
}

// GetTimeFromExcelSerialDate 将Excel序列日期值转换为时间
func GetTimeFromExcelSerialDate(serialDate float64) time.Time {
	// Excel基准日期: 1900年1月0日
	baseDate := time.Date(1899, 12, 30, 0, 0, 0, 0, time.Local)

	// 减1是因为Excel中1900年1月1日是1而不是0
	daysPassed := serialDate - 1

	// 处理Excel的1900年2月29日错误
	if serialDate >= 60 {
		daysPassed -= 1
	}

	// 计算时间
	duration := time.Duration(daysPassed * 24 * float64(time.Hour))
	return baseDate.Add(duration)
}
