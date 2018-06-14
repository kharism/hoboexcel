package hoboexcel

import (
	"bytes"
	"fmt"
	"io"
	"strconv"
)

func AppXmlGenerator(sheetNames []string) io.Reader {
	buff := bytes.Buffer{}
	size := len(sheetNames)
	buff.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>Microsoft Excel</Application><DocSecurity>0</DocSecurity><ScaleCrop>false</ScaleCrop><HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant>
		<vt:variant><vt:i4>`)
	buff.WriteString(strconv.Itoa(size))
	buff.WriteString(`</vt:i4></vt:variant></vt:vector></HeadingPairs><TitlesOfParts>`)
	buff.WriteString(fmt.Sprintf("<vt:vector size=\"%d\" baseType=\"lpstr\">", size))
	for _, v := range sheetNames {
		buff.WriteString(fmt.Sprintf("<vt:lpstr>%s</vt:lpstr>", v))
	}
	buff.WriteString(`</vt:vector></TitlesOfParts><Company></Company><LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc><HyperlinksChanged>false</HyperlinksChanged><AppVersion>15.0300</AppVersion></Properties>`)
	return &buff
}
func WorkbookXMLGenerator(sheetNames []string) io.Reader {
	buff := bytes.Buffer{}
	//size := len(sheetNames)
	//fmt.Println(size)
	buff.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x15" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"><fileVersion appName="xl" lastEdited="6" lowestEdited="6" rupBuild="14420"/><workbookPr defaultThemeVersion="153222"/><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="x15"><x15ac:absPath url="D:\mygo\src\github.com\kharism\effectiveexport\sample\" xmlns:x15ac="http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac"/></mc:Choice></mc:AlternateContent><bookViews><workbookView xWindow="0" yWindow="0" windowWidth="20490" windowHeight="7755" activeTab="1"/></bookViews><sheets>`)
	for i, val := range sheetNames {
		buff.WriteString(fmt.Sprintf(`<sheet name="%s" sheetId="%d" r:id="rId%d"/>`, val, i+1, i+1))
		//fmt.Println("Writing", fmt.Sprintf(`<sheet name="%s" sheetId="%d" r:id="rId%d"/>`, val, i, i))
	}
	buff.WriteString(`</sheets><calcPr calcId="152511"/><extLst><ext uri="{140A7094-0E35-4892-8432-C4D2E57EDEB5}" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"><x15:workbookPr chartTrackingRefBase="1"/></ext></extLst></workbook>`)
	return &buff
}
func WorkbookRelGenerator(sheetNames []string) io.Reader {
	buff := bytes.Buffer{}
	buff.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`)
	relId := 1
	for sheetNum, _ := range sheetNames {
		buff.WriteString(fmt.Sprintf(`<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet%d.xml"/>`, relId, sheetNum))
		relId++
	}
	buff.WriteString(fmt.Sprintf(`<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>`, relId))
	relId++
	buff.WriteString(fmt.Sprintf(`<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>`, relId))
	relId++
	buff.WriteString(fmt.Sprintf(`<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>`, relId))
	relId++
	buff.WriteString(`</Relationships>`)
	return &buff
}
func ContentTypeGenerator(sheetNames []string) io.Reader {
	buff := bytes.Buffer{}
	buff.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="bin" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>`)
	//todo:ganti ke sheet1,sheet2,sheet3 bukan sheetA.xml
	for ids, _ := range sheetNames {
		buff.WriteString(fmt.Sprintf(`<Override PartName="/xl/worksheets/sheet%d.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`, (ids + 1)))
	}
	buff.WriteString(`<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
		<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
		<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
		<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>`)
	buff.WriteString("</Types>")
	return &buff
}
