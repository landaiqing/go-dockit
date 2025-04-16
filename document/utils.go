package document

import (
	"fmt"
	"math/rand"
	"time"
)

// 初始化随机数生成器
func init() {
	rand.New(rand.NewSource(time.Now().UnixNano()))
}

// generateUniqueID 生成一个唯一的ID
func generateUniqueID() string {
	return fmt.Sprintf("%d", rand.Intn(1000000))
}

// boolToInt 将布尔值转换为整数
func boolToInt(b bool) int {
	if b {
		return 1
	}
	return 0
}

// boolToString 将布尔值转换为字符串
func boolToString(b bool) string {
	if b {
		return "1"
	}
	return "0"
}

// twipToCm 将twip转换为厘米
func twipToCm(twip int) float64 {
	return float64(twip) / 1440.0
}

// cmToTwip 将厘米转换为twip
func cmToTwip(cm float64) int {
	return int(cm * 1440.0)
}

// pointToTwip 将磅转换为twip
func pointToTwip(point float64) int {
	return int(point * 20.0)
}

// twipToPoint 将twip转换为磅
func twipToPoint(twip int) float64 {
	return float64(twip) / 20.0
}

// emuToPx 将EMU(English Metric Unit)转换为像素
func emuToPx(emu int) int {
	return emu / 9525
}

// pxToEmu 将像素转换为EMU
func pxToEmu(px int) int {
	return px * 9525
}
