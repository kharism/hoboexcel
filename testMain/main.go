package main

import (
	"fmt"
	"strings"

	ee "github.com/eaciit/hoboexcel"
)

type DummyDataFetcher struct {
	CurRow int
	MaxRow int
}

func (d *DummyDataFetcher) NextRow() []string {
	if d.CurRow <= d.MaxRow {
		res := []string{}
		for i := 0; i < 20; i++ {
			t := colCountToAlphaabet(i)
			//fmt.Println(fmt.Sprintf("Hola%s%d", t, i))
			res = append(res, fmt.Sprintf("Hola%s%d & \u0016 世界", t, d.CurRow))
		}
		d.CurRow++
		fmt.Println(d.CurRow)
		return res
	}
	return nil
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
func main() {
	fetcher := DummyDataFetcher{1, 10}
	ee.Export("dd.xlsx", &fetcher)

	// ee.ExportWorksheet("sheet1.xml", &fetcher)
	// outputFile := "dd.xlsx"
	// file := make(map[string]io.Reader)
	// file["_rels/.rels"] = "_rels/.rels"
	// file["docProps/app.xml"] = "docProps/app.xml"
	// file["docProps/core.xml"] = "docProps/core.xml"
	// file["xl/_rels/workbook.xml.rels"] = "xl/_rels/workbook.xml.rels"
	// file["xl/theme/theme1.xml"] = "xl/theme/theme1.xml"
	// file["xl/worksheets/sheet1.xml"] = "xl/worksheets/sheet1.xml"
	// file["xl/styles.xml"] = "xl/styles.xml"
	// file["xl/workbook.xml"] = "xl/workbook.xml"
	// file["xl/sharedStrings.xml"] = "xl/sharedStrings.xml"
	// file["[Content_Types].xml"] = "[Content_Types].xml"
	// of, _ := os.Create(outputFile)
	// defer of.Close()
	// zipWriter := zip.NewWriter(of)
	// for k, v := range file {
	// 	fWriter, _ := zipWriter.Create(k)
	// 	fh, _ := os.Open(v)
	// 	io.Copy(fWriter, fh)
	// 	fh.Close()
	// }
	// zipWriter.Close()
}
