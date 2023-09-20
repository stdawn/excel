/*
@Time : 2021/12/14 14:09
@Author : LiuKun
@File : style
@Software: GoLand
@Description:
*/

package excel

import (
	"fmt"
	"github.com/xuri/excelize/v2"
	"strconv"
	"strings"
)

const (
	Style400 = "Style400"
	Style410 = "Style410"
	Style420 = "Style420"
	Style430 = "Style430"
	Style440 = "Style440"
	Style340 = "Style340"
	Style240 = "Style240"
	Style140 = "Style140"
	Style040 = "Style040"
	Style041 = "Style041"
	Style042 = "Style042"
	Style043 = "Style043"
	Style044 = "Style044"
	Style034 = "Style034"
	Style024 = "Style024"
	Style014 = "Style014"
	Style004 = "Style004"
	Style104 = "Style104"
	Style204 = "Style204"
	Style304 = "Style304"
	Style404 = "Style404"
	Style403 = "Style403"
	Style402 = "Style402"
	Style401 = "Style401"
	Style444 = "Style444"
	Style000 = "Style000"

	Style141 = "Style141"
	Style241 = "Style241"
	Style431 = "Style431"
	Style421 = "Style421"
	Style411 = "Style411"
)

const (
	DefaultBorderType = 1 //默认边框类型
)

type StyleModel struct {
	Type      string
	BgColor   uint64
	TextColor uint64
}

// NewStyleModel 新建样式模型
func NewStyleModel(t string) *StyleModel {
	sm := new(StyleModel)

	if !strings.Contains(t, "Style") {
		return sm
	}

	color := strings.ReplaceAll(t, "Style", "")
	if len(color) < 3 {
		return sm
	}

	r, e := strconv.Atoi(color[:1])
	g, e1 := strconv.Atoi(color[1:2])
	b, e2 := strconv.Atoi(color[2:])

	if e != nil || e1 != nil || e2 != nil {
		return sm
	}

	if r < 0 || r > 4 || g < 0 || g > 4 || b < 0 || b > 4 {
		return sm
	}

	strMap := map[int]string{
		0: "00", 1: "3F", 2: "7F",
		3: "BF", 4: "FF",
	}

	bgStr := fmt.Sprintf("%s%s%s", strMap[r], strMap[g], strMap[b])
	txStr := fmt.Sprintf("%s%s%s", strMap[4-r], strMap[4-g], strMap[4-b])

	sm.BgColor = RgbStrToHex(bgStr)
	sm.TextColor = RgbStrToHex(txStr)

	return sm
}

// GetStyleId 创建可设置的样式
// horizontal: 0-left, 1-center, 2-right
// shading: (叠加白色)0-横向，1-纵向，2-对角线向上，3-对角线向下，4-从对角线向内，5-从中心向外, 其他纯色填充
//
//boards:0-5 上左下右,斜上，斜下
func (sm *StyleModel) GetStyleId(f *File, borders []int, horizontal, vertical int, shading int) int {
	return f.NewDefaultStyle(sm.BgColor, sm.TextColor, borders, horizontal, vertical, shading)
}

/* 初始化单元格样式 */
func (f *File) initCellStyleMap() {

	if len(f.CellStyleMap) > 0 {
		return
	}
	f.CellStyleMap = make(map[string]*StyleModel)

	styleTypes := []string{
		Style400,
		Style410, Style420, Style430,
		Style440,
		Style340, Style240, Style140,
		Style040,
		Style041, Style042, Style043,
		Style044,
		Style034, Style024, Style014,
		Style004,
		Style104, Style204, Style304,
		Style404,
		Style403, Style402, Style401,
		Style444,
		Style000,
		Style141, Style241,
		Style431, Style421, Style411,
	}

	for _, t := range styleTypes {
		sm := NewStyleModel(t)
		f.CellStyleMap[t] = sm
	}

}

// SetCellsBySameStyle 批量设置相同样式
func (f *File) SetCellsBySameStyle(sheetIndex, startRow, startCol, endRow, endCol int, style int) {

	for r := startRow; r <= endRow; r++ {
		for c := startCol; c <= endCol; c++ {
			f.SetStyleToCell(sheetIndex, r, c, style)
		}
	}
}

// SetCellsByStyles 批量设置样式， 行列数组
func (f *File) SetCellsByStyles(sheetIndex, startRow, startCol int, styles [][]int) {

	for row, colStyles := range styles {
		for col, style := range colStyles {
			f.SetStyleToCell(sheetIndex, startRow+row, startCol+col, style)
		}
	}
}

// SetCellsByStylesWithCol 批量设置样式， 列行数组
func (f *File) SetCellsByStylesWithCol(sheetIndex, startRow, startCol int, styles [][]int) {

	for col, rowStyles := range styles {
		for row, style := range rowStyles {
			f.SetStyleToCell(sheetIndex, startRow+row, startCol+col, style)
		}
	}
}

// SetStyleToCell  设置样式
func (f *File) SetStyleToCell(sheet, row, col, styleId int) {

	errStr := fmt.Sprintf("设置值到[%s-表%d-行%d-列%d]失败: ", f.Path, sheet, row, col)

	if styleId < 0 {
		fmt.Printf(errStr + "无效的样式ID" + "\n")
		return
	}

	sheetName, axis, err := f.GetSheetNameAndAxis(sheet, row, col)
	if err != nil {
		fmt.Printf(errStr + err.Error() + "\n")
		return
	}

	err = f.SetCellStyle(sheetName, axis, axis, styleId)

	if err != nil {
		fmt.Printf(errStr + err.Error() + "\n")
	}
}

// GetStyleFromCell 获取样式
func (f *File) GetStyleFromCell(sheet, row, col int) int {

	errStr := fmt.Sprintf("设置值到[%s-表%d-行%d-列%d]失败: ", f.Path, sheet, row, col)
	sheetName, axis, err := f.GetSheetNameAndAxis(sheet, row, col)
	if err != nil {
		fmt.Printf(errStr + err.Error() + "\n")
		return -1
	}

	id, err := f.GetCellStyle(sheetName, axis)

	if err != nil {
		fmt.Printf(errStr + err.Error() + "\n")
		return -1
	}

	return id
}

// NewDefaultStyle 创建默认样式
// horizontal: 0-left, 1-center, 2-right
// vertical: 0-left, 1-center, 2-right
// shading: (叠加白色)0-横向，1-纵向，2-对角线向上，3-对角线向下，4-从对角线向内，5-从中心向外, 其他纯色填充
// boards:0-5 上左下右,斜上，斜下
// 返回 styleID
func (f *File) NewDefaultStyle(bgColor, textColor uint64, borders []int, horizontal, vertical int, shading int) int {

	style := NewStyle(bgColor, textColor, borders, horizontal, vertical, shading)
	i, err := f.NewStyle(style)
	if err != nil {
		fmt.Printf("创建样式失败：" + err.Error() + "\n")
		return -1
	}

	return i
}

// NewStyle 创建默认样式
// horizontal: 0-left, 1-center, 2-right
// vertical: 0-left, 1-center, 2-right
// shading: (叠加白色)0-横向，1-纵向，2-对角线向上，3-对角线向下，4-从对角线向内，5-从中心向外, 其他纯色填充
// boards:0-5 上左下右,斜上，斜下
// 返回 style对象
func NewStyle(bgColor, textColor uint64, borders []int, horizontal, vertical int, shading int) *excelize.Style {
	if bgColor >= NOColorHex {
		bgColor = WhiteColor
	}
	if textColor >= NOColorHex {
		textColor = BlackColor
	}

	if horizontal < 0 || horizontal > 2 {
		horizontal = 0
	}

	bs := make([]excelize.Border, 0)
	borderKeys := []string{"top", "left", "bottom", "right", "diagonalUp", "diagonalDown"}
	for _, index := range borders {
		if index >= 0 && index < len(borderKeys) {
			b := excelize.Border{Type: borderKeys[index], Color: HexToRgbStr(HaftGrayColor), Style: DefaultBorderType}
			bs = append(bs, b)
		}
	}

	var fill excelize.Fill

	if shading >= 0 && shading <= 5 {
		fill = excelize.Fill{
			Type:    "gradient",
			Color:   []string{HexToRgbStr(WhiteColor), HexToRgbStr(bgColor)},
			Shading: shading,
		}
	} else {
		fill = excelize.Fill{
			Type:    "pattern",
			Pattern: 1,
			Color:   []string{HexToRgbStr(bgColor)},
		}
	}

	style := &excelize.Style{
		Border: bs,
		Fill:   fill,
		Alignment: &excelize.Alignment{
			Horizontal: []string{"left", "center", "right"}[horizontal],
			Vertical:   []string{"left", "center", "right"}[vertical],
			WrapText:   true,
		},
		Font: &excelize.Font{
			Bold:      false,
			Italic:    false,
			Underline: "none",
			Family:    DefaultFont,
			Size:      DefaultFontSize,
			Strike:    false,
			Color:     HexToRgbStr(textColor),
		},
	}

	return style
}
