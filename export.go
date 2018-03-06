package hoboexcel

import (
	"archive/zip"
	"bufio"
	"encoding/xml"
	"fmt"
	"html"
	"io"
	"os"
	"strconv"
	"strings"
	"time"
	//"golang.org/x/text/unicode/norm"
)

var TempDir = "./xl/worksheets/"

func CleanNonUtfAndControlChar(s string) string {
	s = strings.Map(func(r rune) rune {
		if r <= 31 {
			return -1
		}
		return r
	}, s)
	return s
}
func ExportWorksheet(filename string, rows RowFetcher) {
	file, _ := os.Create(filename)
	defer file.Close()
	shaStr, _ := os.Create(filename + ".ss")
	defer shaStr.Close()
	Writer := bufio.NewWriter(file)
	SharedStrWriter := bufio.NewWriter(shaStr)

	Writer.WriteString("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"x14ac\" xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\">")
	Writer.WriteString("<sheetViews><sheetView tabSelected=\"1\" workbookViewId=\"0\"><selection activeCell=\"A1\" sqref=\"A1\"/></sheetView></sheetViews>")
	Writer.WriteString("<sheetFormatPr defaultRowHeight=\"15\" x14ac:dyDescent=\"0.25\"/>")
	Writer.WriteString("<sheetData>")

	SharedStrWriter.WriteString("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>")
	SharedStrWriter.WriteString("<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"0\" uniqueCount=\"0\">")

	rowCount := 1
	//uniqueString := map[string]int{}
	//sortedUsedStr := []string{}
	cellsCount := 0
	for {
		raw_row := rows.NextRow()
		if raw_row == nil {
			break
		}
		rr := row{}
		rr.R = rowCount
		for idx, val := range raw_row {
			colName := colCountToAlphaabet(idx)
			newCol := XlsxC{}
			newCol.T = "s"
			newCol.R = fmt.Sprintf("%s%d", colName, rowCount)
			// idxStr, ok := uniqueString[val]
			// if ok {
			// 	newCol.V = strconv.Itoa(idxStr)
			// } else {
			// 	uniqueString[val] = len(uniqueString)
			// 	newCol.V = strconv.Itoa(uniqueString[val])
			// 	sortedUsedStr = append(sortedUsedStr, val)
			// }
			newCol.V = strconv.Itoa(cellsCount)
			cellsCount++
			rr.C = append(rr.C, newCol)
			fmt.Println(val, html.EscapeString(CleanNonUtfAndControlChar(val)))
			SharedStrWriter.WriteString(fmt.Sprintf("<si><t>%s</t></si>", html.EscapeString(CleanNonUtfAndControlChar(val))))
		}
		rr.Spans = "1:10"
		rr.Descent = "0.25"
		bb, e := xml.Marshal(rr)
		if e != nil {
			fmt.Println("Encoder error", e.Error())
			fmt.Println(rr)
			os.Exit(1)
		}
		//fmt.Println(string(bb))
		pp, e := Writer.Write(bb)
		if e != nil {
			fmt.Println("Writer error", e.Error())
			fmt.Println(rr)
			os.Exit(1)
		}
		if pp != len(bb) {
			fmt.Println("Writer error2")
		}
		if rowCount%1000 == 0 {
			SharedStrWriter.Flush()
			Writer.Flush()
		}
		rowCount++

	}
	Writer.WriteString("</sheetData>")
	Writer.WriteString("<pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/>")
	Writer.WriteString("</worksheet>")
	Writer.Flush()
	SharedStrWriter.WriteString("</sst>")
	SharedStrWriter.Flush()
	//write shared strings
	//sharedString := xlsxSST{}
	//sharedString.Count = len(sortedUsedStr)
	//sharedString.UniqueCount = len(sortedUsedStr)
	// for _, val := range sortedUsedStr {
	// 	ss := xlsxSI{}
	// 	ss.T = val
	// 	sharedString.SI = append(sharedString.SI, ss)
	// }

	// encoder := xml.NewEncoder(shaStr)
	// e := encoder.Encode(sharedString)
	// if e != nil {
	// 	fmt.Println(e.Error())
	// }

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
func Export(filename string, fetcher RowFetcher) {
	now := time.Now()
	sheetName := now.Format("20060102150405") //filename should be (pseudo)random
	ExportWorksheet(sheetName, fetcher)
	outputFile := filename
	file := make(map[string]io.Reader)
	file["_rels/.rels"] = DummyRelsDotRels()
	file["docProps/app.xml"] = DummyAppXml()
	file["docProps/core.xml"] = DummyCoreXml()
	file["xl/_rels/workbook.xml.rels"] = DummyWorkbookRels()
	file["xl/theme/theme1.xml"] = DummyThemeXml()
	file["xl/worksheets/sheet1.xml"], _ = os.Open(sheetName)
	file["xl/styles.xml"] = DummyStyleXml()
	file["xl/workbook.xml"] = DummyWorkbookXml()
	file["xl/sharedStrings.xml"], _ = os.Open(sheetName + ".ss")
	file["[Content_Types].xml"] = DummyContentTypes()
	of, _ := os.Create(outputFile)
	defer of.Close()
	zipWriter := zip.NewWriter(of)
	for k, v := range file {
		fWriter, _ := zipWriter.Create(k)
		io.Copy(fWriter, v)

	}
	zipWriter.Close()
	(file["xl/sharedStrings.xml"].(*os.File)).Close()
	(file["xl/worksheets/sheet1.xml"].(*os.File)).Close()
	e := os.Remove("./" + now.Format("20060102150405"))
	if e != nil {
		fmt.Println(e.Error())
	}
	e = os.Remove("./" + now.Format("20060102150405") + ".ss")
	if e != nil {
		fmt.Println(e.Error())
	}
}
