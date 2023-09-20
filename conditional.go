/*
@Time : 2021/1/27 11:44
@Author : LiuKun
@File : conditional
@Software: GoLand
@Description:
*/

package excel

import (
	"errors"
	"fmt"
	"github.com/xuri/excelize/v2"
)

const (
	RedText        = "RedText"
	GreenText      = "GreenText"
	BlueText       = "BlueText"
	BlackText      = "BlackText"
	NormalBlueText = "NormalBlueText"
	PurPleText     = "PurPleText"

	OrangeBg = "OrangeBg"
	GrayBg   = "GrayBg"
)

// ConditionalFormat 条件格式
type ConditionalFormat struct {
	SheetIndex    int
	StartRow      int
	StartCol      int
	EndRow        int
	EndCol        int
	FormatOptions []excelize.ConditionalFormatOptions //type, format, criteria、value、minimum、maximum, 单元格的值
}

// TextContain 文本包含参数
type TextContain struct {
	Text      string
	IsContain bool
	Style     string
}

// ValueCompareModel  值比较包含参数
type ValueCompareModel struct {
	Value    float64
	Criteria string
	Style    string
}

// SetConditionalToCells 设置多组数据; sheetIndex:表索引，从0开始; startRow:开始行, >=1; startCol:开始列，>=1; values:数据数组
func (f *File) SetConditionalToCells(c *ConditionalFormat) error {
	sheetName := f.GetSheetName(c.SheetIndex)
	if len(sheetName) < 1 {
		return errors.New("找不到工作表")
	}
	startAxis, err := excelize.CoordinatesToCellName(c.StartCol, c.StartRow)
	if err != nil {
		return err
	}
	endAxis, err := excelize.CoordinatesToCellName(c.EndCol, c.EndRow)
	if err != nil {
		return err
	}

	axis := ""
	if startAxis == endAxis {
		axis = startAxis
	} else {
		axis = fmt.Sprintf("%s:%s", startAxis, endAxis)
	}
	return f.SetConditionalFormat(sheetName, axis, c.FormatOptions)
}

func (f *File) initConditionalStyleMap() {
	f.ConditionalStyleMap = make(map[string]int)

	keys := []string{
		RedText, GreenText, BlueText, BlackText, NormalBlueText, PurPleText,
		OrangeBg, GrayBg,
	}
	textColors := []uint64{
		RedColor, GreenColor, BlueColor, BlackColor, NormalBlueColor, PurPleColor,
		NOColorHex, NOColorHex,
	}
	bgColors := []uint64{
		NOColorHex, NOColorHex, NOColorHex, NOColorHex, NOColorHex, NOColorHex,
		DarkOrangeColor, GrayColor,
	}

	for i, k := range keys {
		style, err := f.createConditionalStyle(textColors[i], bgColors[i])
		if err == nil {
			f.ConditionalStyleMap[k] = style
		}
	}
}

/*
	创建条件格式

textColorHex文本颜色， fillColorHex填充颜色
*/
func (f *File) createConditionalStyle(textColorHex, fillColorHex uint64) (int, error) {

	style := NewStyle(fillColorHex, textColorHex, []int{0, 1, 2, 3}, 1, 1, 0)
	return f.NewConditionalStyle(style)
}

// SetValueCompareCF 值比较 条件格式 [value,
func (f *File) SetValueCompareCF(sheetIndex, startRow, rowLen, startCol, colLen int, cfs ...ValueCompareModel) {
	c := new(ConditionalFormat)
	c.SheetIndex = sheetIndex
	c.StartRow = startRow
	c.EndRow = startRow + rowLen - 1
	c.StartCol = startCol
	c.EndCol = startCol + colLen - 1
	c.FormatOptions = make([]excelize.ConditionalFormatOptions, 0)
	for _, cf := range cfs {
		s, o := f.ConditionalStyleMap[cf.Style]
		if o {
			m := excelize.ConditionalFormatOptions{
				Type:     "cell",
				Format:   s,
				Criteria: cf.Criteria,
				Value:    fmt.Sprintf("%f", cf.Value),
			}
			c.FormatOptions = append(c.FormatOptions, m)
		}
	}
	err := f.SetConditionalToCells(c)
	if err != nil {
		fmt.Printf("条件格式设置失败：%s\n", err.Error())
	}
}

// SetTextContainsCF 文本包含条件格式 cfs [text, contains, style]
func (f *File) SetTextContainsCF(sheetIndex, startRow, rowLen, startCol, colLen int, cfs ...TextContain) {
	c := new(ConditionalFormat)
	c.SheetIndex = sheetIndex
	c.StartRow = startRow
	c.EndRow = startRow + rowLen - 1
	c.StartCol = startCol
	c.EndCol = startCol + colLen - 1
	c.FormatOptions = make([]excelize.ConditionalFormatOptions, 0)

	for _, cf := range cfs {
		if len(cf.Text) < 1 {
			continue
		}
		s, o := f.ConditionalStyleMap[cf.Style]
		if o {
			v := "0,1"
			if cf.IsContain {
				v = "1,0"
			}
			m := excelize.ConditionalFormatOptions{
				Type:     "formula",
				Format:   s,
				Criteria: fmt.Sprintf("IF(ISNUMBER(FIND(\"%s\",INDIRECT(ADDRESS(ROW(),COLUMN())))),%s)", cf.Text, v),
			}
			c.FormatOptions = append(c.FormatOptions, m)
		}
	}
	err := f.SetConditionalToCells(c)
	if err != nil {
		fmt.Printf("条件格式设置失败：%s\n", err.Error())
	}
}

func (f *File) DeleteConditionalFromCells(sheetIndex, startRow, endRow, startCol, endCol int) {

	errStr := fmt.Sprintf("删除条件格式[%s-表%d-位置[(%d,%d)-(%d-%d)]失败: ", f.Path, sheetIndex, startRow, startCol, endRow, endCol)

	sheetName := f.GetSheetName(sheetIndex)
	if len(sheetName) < 1 {
		fmt.Printf(errStr + "找不到工作表")
	}
	startAxis, err := excelize.CoordinatesToCellName(startCol, startRow)
	if err != nil {
		fmt.Printf(errStr + err.Error() + "\n")
	}
	endAxis, err := excelize.CoordinatesToCellName(endCol, endRow)
	if err != nil {
		fmt.Printf(errStr + err.Error() + "\n")
	}

	axis := ""
	if startAxis == endAxis {
		axis = startAxis
	} else {
		axis = fmt.Sprintf("%s:%s", startAxis, endAxis)
	}

	err = f.UnsetConditionalFormat(sheetName, axis)
	if err != nil {
		fmt.Printf(errStr + err.Error() + "\n")
	}
}
