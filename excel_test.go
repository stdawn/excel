/**
 * @Time: 2023/9/20 16:32
 * @Author: LiuKun
 * @File: excel_test.go
 * @Description:
 */

package excel

import "testing"

func TestExcel(t *testing.T) {

	n := GetColumnNumberFromName("AA")
	t.Log(n)

	s := GetColumnNameFromNumber(21)
	t.Log(s)
}
