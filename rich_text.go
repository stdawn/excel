/*
@Time : 2021/2/24 11:15
@Author : LiuKun
@File : rich_text
@Software: GoLand
@Description:
*/

package excel

import "github.com/xuri/excelize/v2"

// RichText 富文本
type RichText struct {
	Text      string
	Color     uint64
	Size      float64
	Bold      bool
	Italic    bool
	Underline bool
}

type RichTextSlice []RichText

// ToString 获取文本
func (rs RichTextSlice) ToString() string {
	s := ""
	for _, r := range rs {
		s += r.Text
	}
	return s
}

func (r RichText) toRichTextRun() excelize.RichTextRun {
	underline := "none"
	if r.Underline {
		underline = "single"
	}
	size := r.Size
	if size < 1 {
		size = DefaultFontSize
	}
	return excelize.RichTextRun{
		Text: r.Text,
		Font: &excelize.Font{
			Bold:      r.Bold,
			Color:     HexToRgbStr(r.Color),
			Italic:    r.Italic,
			Family:    DefaultFont,
			Size:      size,
			Underline: underline,
		},
	}
}

func (rs RichTextSlice) toRichTextRunSlice() []excelize.RichTextRun {
	runSlice := make([]excelize.RichTextRun, 0)
	for _, r := range rs {
		runSlice = append(runSlice, r.toRichTextRun())
	}
	return runSlice
}
