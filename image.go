/*
@Time : 2021/5/7 9:58
@Author : LiuKun
@File : chart
@Software: GoLand
@Description:
*/

package excel

import (
	"fmt"
	"github.com/xuri/excelize/v2"
)

// AddImage 添加图片
func (f *File) AddImage(sheet, row, col int, xyScale [2]float64, xyOffset [2]int, imageName, imageExt string, imageBytes []byte) {
	errStr := fmt.Sprintf("添加图片[%s-表%d-行%d-列%d]失败: ", f.Path, sheet, row, col)

	sheetName, axis, err := f.GetSheetNameAndAxis(sheet, row, col)
	if err != nil {
		fmt.Printf(errStr + err.Error() + "\n")
		return
	}

	format := &excelize.GraphicOptions{
		AltText: imageName,
		OffsetX: xyOffset[0],
		OffsetY: xyOffset[1],
		ScaleX:  xyScale[0],
		ScaleY:  xyScale[1],
	}

	err = f.AddPictureFromBytes(sheetName, axis, &excelize.Picture{
		Extension: imageExt,
		File:      imageBytes,
		Format:    format,
	})

	if err != nil {
		fmt.Printf("%s:%s\n", errStr, err.Error())
	}
}
