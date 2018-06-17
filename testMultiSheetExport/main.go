package main

import (
	"strconv"
	"strings"

	"github.com/eaciit/hoboexcel"
)

type SimpleSheetFetcher struct {
	curId      int
	sheets     []*SimpleSheet
	sheetNames []string
}
type SimpleSheet struct {
	name   string
	curRow int
	maxRow int
	maxCol int
}

func (a *SimpleSheetFetcher) GetSheetNames() []string {
	if len(a.sheetNames) > 0 {
		return a.sheetNames
	} else {
		hasil := []string{}
		for _, c := range a.sheets {
			hasil = append(hasil, c.GetSheetName())
		}
		return hasil
	}
}
func colCountToAlphaabet(idx int) string {
	var colName string
	if idx >= 26 {
		firstLetter := (idx / 26) - 1
		secondLetter := (idx % 26)
		colName = string(65+firstLetter) + string(65+secondLetter)
	} else {
		colName = string(65 + idx)
	}
	return strings.ToUpper(colName)
}
func (a *SimpleSheetFetcher) NextSheet() hoboexcel.Sheet {
	if a.curId < len(a.sheets) {
		p := a.sheets[a.curId]
		a.curId++
		return p
	} else {
		return nil
	}
}

func (a *SimpleSheet) NextRow() []string {
	if a.curRow == a.maxRow {
		return nil
	}
	results := []string{}
	for i := 0; i < a.maxCol; i++ {
		results = append(results, "Writing "+a.GetSheetName()+" row "+strconv.Itoa(a.curRow+1)+" Col "+colCountToAlphaabet(i))
	}
	a.curRow++
	return results
}
func (a *SimpleSheet) GetSheetName() string {
	return a.name
}
func main() {
	sheetsSource := &SimpleSheetFetcher{}
	sheet1 := &SimpleSheet{}
	sheet1.maxRow = 1
	sheet1.maxCol = 3
	sheet1.name = "SheetA"
	sheet2 := &SimpleSheet{}
	sheet2.name = "SheetB"
	sheet2.maxRow = 1
	sheet2.maxCol = 3
	sheetsSource.sheets = append(sheetsSource.sheets, sheet1)
	sheetsSource.sheets = append(sheetsSource.sheets, sheet2)
	hoboexcel.ExportMultisheet("dd.xlsx", sheetsSource)
}
