/*
@Time : 2021/1/16 10:20
@Author : LiuKun
@File : base
@Software: GoLand
@Description:
*/

package excel

import (
	"errors"
	"fmt"
	"github.com/xuri/excelize/v2"
	"strconv"
	"strings"
)

/* 颜色 */
const (
	WhiteColor        uint64 = 0xFFFFFF
	RedColor          uint64 = 0xFF0000
	GreenColor        uint64 = 0x00B050
	NormalGreenColor  uint64 = 0x00FF00
	DarkGreenColor    uint64 = 0x007F00
	BlueColor         uint64 = 0x00B0F0
	BlackColor        uint64 = 0x000000
	NormalBlueColor   uint64 = 0x0000FF
	PurPleColor       uint64 = 0xFF00FF
	BluePurPleColor   uint64 = 0x7030A0
	NormalOrangeColor uint64 = 0xFFA500
	DarkOrangeColor   uint64 = 0xFF3F00
	GrayColor         uint64 = 0x7F9696
	LightGrayColor    uint64 = 0x808080
	DarkGrayColor     uint64 = 0x404040
	DarkRedColor      uint64 = 0x800000
	ShadeRedColor     uint64 = 0xBF0000 //暗红

	HaftGrayColor uint64 = 0x7F7F7F //边框灰色
)

const (
	NOColorHex      uint64 = 0xFFFFFF + 1
	DefaultFont            = "微软雅黑"
	DefaultFontSize        = 10.0
)

type NoValueType struct{}
type FormulaType string

var NoValue = NoValueType{}

// File excel文件模型
type File struct {
	*excelize.File
	ConditionalStyleMap map[string]int
	CellStyleMap        map[string]*StyleModel
}

// NewExcelFile 新建对象
func NewExcelFile(filename string) (*File, error) {
	f, err := excelize.OpenFile(filename)
	if err != nil {
		return nil, err
	}
	nf := new(File)
	nf.File = f
	nf.initConditionalStyleMap()
	nf.initCellStyleMap()
	return nf, nil
}

// IsSheetExistWithName 是否存在表
func (f *File) IsSheetExistWithName(name string) bool {
	sheetIndex, err := f.GetSheetIndex(name)
	if err != nil {
		return false
	}
	return sheetIndex >= 0
}

/*
	设置Cell的公式

表Index(从0开始)， 行index(从1开始， 列Index从1
*/
func (f *File) SetFormulaToCell(sheet, row, col int, v string) {

	errStr := fmt.Sprintf("设置公式到[%s-表%d-行%d-列%d]失败: ", f.Path, sheet, row, col)
	if len(v) < 1 {
		fmt.Printf(errStr + "公式为空")
		return
	}

	sheetName, axis, err := f.GetSheetNameAndAxis(sheet, row, col)
	if err != nil {
		fmt.Printf(errStr + err.Error() + "\n")
		return
	}

	err = f.SetCellFormula(sheetName, axis, v)

	if err != nil {
		fmt.Printf(errStr + err.Error() + "\n")
	}

}

// GetValueFromCell 获取Cell的值, 表Index(从0开始), 行index(从1开始), 列Index(从1开始)
func (f *File) GetValueFromCell(sheet, row, col int) string {
	errStr := fmt.Sprintf("获取值[%s-表%d-行%d-列%d]失败: ", f.Path, sheet, row, col)

	sheetName, axis, err := f.GetSheetNameAndAxis(sheet, row, col)
	if err != nil {
		fmt.Printf(errStr + err.Error() + "\n")
		return ""
	}

	v, err := f.GetCellValue(sheetName, axis)

	if err != nil {
		fmt.Printf(errStr + err.Error() + "\n")
		return ""
	}
	return v
}

// SetValueToCell 设置Cell的值, 表Index(从0开始), 行index(从1开始), 列Index(从1开始)
func (f *File) SetValueToCell(sheet, row, col int, v interface{}) {

	errStr := fmt.Sprintf("设置值到[%s-表%d-行%d-列%d]失败: ", f.Path, sheet, row, col)
	if v == nil {
		fmt.Printf(errStr + "值为空")
		return
	}

	_, ok := v.(NoValueType)
	if ok {
		return
	}

	sheetName, axis, err := f.GetSheetNameAndAxis(sheet, row, col)
	if err != nil {
		fmt.Printf(errStr + err.Error() + "\n")
		return
	}

	s, ok := v.(FormulaType)
	if ok {
		err = f.SetCellFormula(sheetName, axis, string(s))
	} else {
		r, ok1 := v.(RichText)
		if ok1 {
			v = RichTextSlice{r}
		}
		rs, ok2 := v.(RichTextSlice)
		if ok2 {
			err = f.SetCellRichText(sheetName, axis, rs.toRichTextRunSlice())
		} else {
			err = f.SetCellValue(sheetName, axis, v)
		}

	}

	if err != nil {
		fmt.Printf(errStr + err.Error() + "\n")
	}

}

// SetCellsByValues 设置多组数据,,先行后列; sheetIndex:表索引，从0开始; startRow:开始行, >=1; params startCol:开始列，>=1; values:数据数组
func (f *File) SetCellsByValues(sheetIndex, startRow, startCol int, values [][]interface{}) {

	for row, colValues := range values {
		for col, value := range colValues {
			f.SetValueToCell(sheetIndex, startRow+row, startCol+col, value)
		}
	}
}

// SetCellsByValuesWithCol 设置多组数据,先列后行; sheetIndex:表索引，从0开始; startRow:开始行, >=1; params startCol:开始列，>=1; values:数据数组
func (f *File) SetCellsByValuesWithCol(sheetIndex, startRow, startCol int, values [][]interface{}) {

	for col, rowValues := range values {
		for row, value := range rowValues {
			f.SetValueToCell(sheetIndex, startRow+row, startCol+col, value)
		}
	}
}

// GetSheetNameAndAxis 获取表名和位置; sheetIndex:表索引; row:行号; col:列号; 返回sheetName, axis, err
func (f *File) GetSheetNameAndAxis(sheetIndex, row, col int) (string, string, error) {
	sheetName := f.GetSheetName(sheetIndex)
	if len(sheetName) < 1 {
		return "", "", errors.New("找不到工作表")
	}
	if row < 1 {
		return sheetName, "", errors.New("行号不存在")
	}
	if col < 1 {
		return sheetName, "", errors.New("列号不存在")
	}

	axis, err := excelize.CoordinatesToCellName(col, row)

	return sheetName, axis, err

}

// CopyRowTo 复制行
func (f *File) CopyRowTo(sheet, row, toRow int) {

	errStr := fmt.Sprintf("复制行[%s-表%d-行%d->行%d]失败: ", f.Path, sheet, row, toRow)
	sheetName := f.GetSheetName(sheet)
	if len(sheetName) < 1 {
		fmt.Printf(errStr + "找不到工作表")
		return
	}
	err := f.DuplicateRowTo(sheetName, row, toRow)
	if err != nil {
		fmt.Printf(errStr + err.Error() + "\n")
	}
}

// InsertColsBefore 插入多列
func (f *File) InsertColsBefore(sheet, col int, n int) {

	errStr := fmt.Sprintf("插入列[%s-表%d-列%d]失败: ", f.Path, sheet, col)
	sheetName := f.GetSheetName(sheet)
	if len(sheetName) < 1 {
		fmt.Printf(errStr + "找不到工作表")
		return
	}
	colStr, err := excelize.ColumnNumberToName(col)
	if err != nil {
		fmt.Printf(errStr + err.Error() + "\n")
		return
	}
	err = f.InsertCols(sheetName, colStr, n)
	if err != nil {
		fmt.Printf(errStr + err.Error() + "\n")
	}
}

// DeleteChartAtCell 删除图表
func (f *File) DeleteChartAtCell(sheet, row, col int) {
	errStr := fmt.Sprintf("删除图表[%s-表%d-行%d-列%d]失败: ", f.Path, sheet, row, col)
	sheetName, axis, err := f.GetSheetNameAndAxis(sheet, row, col)
	if err != nil {
		fmt.Printf(errStr + err.Error() + "\n")
		return
	}
	err = f.DeleteChart(sheetName, axis)
	if err != nil {
		fmt.Printf(errStr + err.Error() + "\n")
	}
}

// DeleteCol 删除列
func (f *File) DeleteCol(sheet, col int) {

	errStr := fmt.Sprintf("删除列[%s-表%d-列%d]失败: ", f.Path, sheet, col)
	sheetName := f.GetSheetName(sheet)
	if len(sheetName) < 1 {
		fmt.Printf(errStr + "找不到工作表")
		return
	}
	colStr, err := excelize.ColumnNumberToName(col)
	if err != nil {
		fmt.Printf(errStr + err.Error() + "\n")
		return
	}
	err = f.RemoveCol(sheetName, colStr)
	if err != nil {
		fmt.Printf(errStr + err.Error() + "\n")
	}
}

// DeleteCols 删除多列
func (f *File) DeleteCols(sheet, startCol, l int) {
	for i := l - 1; i >= 0; i-- {
		f.DeleteCol(sheet, startCol+i)
	}
}

// DeleteRow 删除行
func (f *File) DeleteRow(sheet, row int) {

	errStr := fmt.Sprintf("删除行[%s-表%d-行%d]失败: ", f.Path, sheet, row)
	sheetName := f.GetSheetName(sheet)
	if len(sheetName) < 1 {
		fmt.Printf(errStr + "找不到工作表")
		return
	}
	err := f.RemoveRow(sheetName, row)
	if err != nil {
		fmt.Printf(errStr + err.Error() + "\n")
	}
}

// DeleteRows 删除多行
func (f *File) DeleteRows(sheet, startRow, l int) {
	for i := l - 1; i >= 0; i-- {
		f.DeleteRow(sheet, startRow+i)
	}
}

// MergeCellAt 合并单元格
func (f *File) MergeCellAt(sheetIndex, startRow, startCol, endRow, endCol int) {

	errStr := fmt.Sprintf("合并单元格[%s-表%d-位置(%d,%d)-(%d-%d)]失败: ", f.Path, sheetIndex, startRow, startCol, endRow, endCol)
	sheetName := f.GetSheetName(sheetIndex)
	if len(sheetName) < 1 {
		fmt.Printf(errStr + "找不到工作表")
		return
	}
	axisStart, err := excelize.CoordinatesToCellName(startCol, startRow)
	if err != nil {
		fmt.Printf(errStr + err.Error() + "\n")
		return
	}

	axisEnd, err := excelize.CoordinatesToCellName(endCol, endRow)
	if err != nil {
		fmt.Printf(errStr + err.Error() + "\n")
		return
	}

	err = f.MergeCell(sheetName, axisStart, axisEnd)
	if err != nil {
		fmt.Printf(errStr + err.Error() + "\n")
	}
}

// UnMergeCellAt 取消合并单元格
func (f *File) UnMergeCellAt(sheetIndex, startRow, startCol, endRow, endCol int) {

	errStr := fmt.Sprintf("取消合并单元格[%s-表%d-位置(%d,%d)-(%d-%d)]失败: ", f.Path, sheetIndex, startRow, startCol, endRow, endCol)
	sheetName := f.GetSheetName(sheetIndex)
	if len(sheetName) < 1 {
		fmt.Printf(errStr + "找不到工作表")
		return
	}
	axisStart, err := excelize.CoordinatesToCellName(startCol, startRow)
	if err != nil {
		fmt.Printf(errStr + err.Error() + "\n")
		return
	}

	axisEnd, err := excelize.CoordinatesToCellName(endCol, endRow)
	if err != nil {
		fmt.Printf(errStr + err.Error() + "\n")
		return
	}

	err = f.UnmergeCell(sheetName, axisStart, axisEnd)
	if err != nil {
		fmt.Printf(errStr + err.Error())
	}
}

// SetLocationLinkToCell 设置本地连接 [sheetIndex, row, col] 连接到[toSheetIndex,toRow,toCol]
func (f *File) SetLocationLinkToCell(sheetIndex, row, col, toSheetIndex, toRow, toCol int) {

	errStr := fmt.Sprintf("设置连接[%s-表%d-(%d行,%d列)]失败: ", f.Path, sheetIndex, row, col)
	sheetName := f.GetSheetName(sheetIndex)
	if len(sheetName) < 1 {
		fmt.Printf(errStr + "找不到工作表")
		return
	}

	toSheetName := f.GetSheetName(toSheetIndex)
	if len(toSheetName) < 1 {
		fmt.Printf(errStr + "找不到连接到的工作表")
		return
	}
	axis, err := excelize.CoordinatesToCellName(row, col)
	if err != nil {
		fmt.Printf(errStr + err.Error() + "\n")
		return
	}

	toAxis, err := excelize.CoordinatesToCellName(toRow, toCol)
	if err != nil {
		fmt.Printf(errStr + err.Error() + "\n")
		return
	}

	err = f.SetCellHyperLink(sheetName, axis, fmt.Sprintf("%s!%s", toSheetName, toAxis), "Location")
	if err != nil {
		fmt.Printf(errStr + err.Error())
	}
}

// SetRefName 设置名称
func (f *File) SetRefName(name, ref string) {
	_ = f.DeleteDefinedName(&excelize.DefinedName{
		Name: name,
	})

	err := f.SetDefinedName(&excelize.DefinedName{
		Name:     name,
		RefersTo: ref,
	})

	if err != nil {
		fmt.Printf(fmt.Sprintf("设置名称[%s]%s失败：%s\n", f.Path, name, err.Error()))
	}

}

// GetRefName 获取名称引用
func (f *File) GetRefName(name string) string {
	ns := f.GetDefinedName()
	for _, n := range ns {
		if n.Name == name {
			return n.RefersTo
		}
	}
	return ""
}

// HexToRgbStr hex -> str
func HexToRgbStr(hex uint64) string {
	if hex >= NOColorHex {
		hex = WhiteColor
	}
	return fmt.Sprintf("%06x", hex)
}

// RgbStrToHex str -> Hex
func RgbStrToHex(rgbStr string) uint64 {
	s := strings.ReplaceAll(rgbStr, "#", "")
	if len(s) != 6 {
		return NOColorHex
	}

	hex, e := strconv.ParseInt(s, 16, 64)
	if e != nil {
		return NOColorHex
	}
	return uint64(hex)
}

// GetColumnNameFromNumber 从列数获取列名, 失败时返回空字符串
func GetColumnNameFromNumber(num int) string {
	s, e := excelize.ColumnNumberToName(num)
	if e != nil {
		fmt.Printf("从列数获取列名失败%s\n", e.Error())
		return ""
	}
	return s
}

// GetColumnNumberFromName 从列数获取列名, 失败时返回-1
func GetColumnNumberFromName(name string) int {
	n, e := excelize.ColumnNameToNumber(name)
	if e != nil {
		fmt.Printf("从列名获取列数失败%s\n", e.Error())
		return -1
	}
	return n
}
