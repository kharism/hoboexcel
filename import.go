package hoboexcel

import (
	"archive/zip"
	"bufio"
	"encoding/xml"
	"errors"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"strconv"
	"strings"
)

var READ_TEMP_DIR = "./" //dont forget to end it with path separator
var PARTITION_SIZE = 300 //the smaller the faster but it will produce more temporary file

type XlsxRowFetcher struct {
	Filename        string
	ZipFile         *zip.ReadCloser
	Decoder         *xml.Decoder
	CurSheet        io.ReadCloser
	IsUsingRamCache bool //set this to true if your sharedstring is relatively small
	curPartitionId  int
	cacheSharedStr  []string
}

//seek string with some caching mechanism
func (r *XlsxRowFetcher) SeekString(index int) string {
	fileId := index / PARTITION_SIZE
	if index >= PARTITION_SIZE {
		index = index % PARTITION_SIZE
	}
	if fileId == r.curPartitionId && len(r.cacheSharedStr) > 0 {
		return r.cacheSharedStr[index]
	} else {
		curFile, _ := os.Open(READ_TEMP_DIR + r.Filename + "ss" + strconv.Itoa(fileId))
		defer curFile.Close()
		decoder := xml.NewDecoder(curFile)
		//curIdx := 0
		tempStr := []string{}
		for {
			tok, _ := decoder.Token()
			if tok == nil {
				break
			}
			switch se := tok.(type) {
			case xml.StartElement:
				if se.Name.Local == "t" {
					tok2, _ := decoder.Token()
					cd := tok2.(xml.CharData)
					//fmt.Println("%d,%s", preIdx, string(cd))
					tempStr = append(tempStr, string(cd))
				}
			}
		}
		//fmt.Println(tempStr)
		r.cacheSharedStr = tempStr
		r.curPartitionId = fileId
		return r.cacheSharedStr[index]
	}
	return ""
}
func SeekString(filename string, index int) string {
	fileId := index / PARTITION_SIZE
	//preIdx := index
	if index >= PARTITION_SIZE {
		index = index % PARTITION_SIZE
	}
	curFile, _ := os.Open(READ_TEMP_DIR + filename + "ss" + strconv.Itoa(fileId))
	defer curFile.Close()
	decoder := xml.NewDecoder(curFile)
	curIdx := 0
	for {
		tok, _ := decoder.Token()
		if tok == nil {
			break
		}
		switch se := tok.(type) {
		case xml.StartElement:
			if se.Name.Local == "t" {
				if curIdx == index {
					tok2, _ := decoder.Token()
					cd := tok2.(xml.CharData)
					//fmt.Println("%d,%s", preIdx, string(cd))
					return string(cd)
				}
				curIdx++
			}
		default:
			break
		}
	}
	//fmt.Println(preIdx, index)
	return ""
}
func (s *XlsxRowFetcher) Close() error {
	e := s.ZipFile.Close()
	if e != nil {
		return e
	}
	//fmt.Println(TempDir, s.Filename+"ss*")
	sharedStringTemps, _ := filepath.Glob(READ_TEMP_DIR + s.Filename + "ss*")
	for _, f := range sharedStringTemps {
		//fmt.Println("Removing", f)
		os.Remove(f)
	}
	return nil
}

func PartitionSharedString(filename string) error {
	rr, err := zip.OpenReader(filename)
	baseFilename := filepath.Base(filename)
	if err != nil {
		fmt.Println(err.Error())
		return err
	}
	defer rr.Close()
	var sharedStrFile *zip.File
	for _, f := range rr.File {
		if strings.Contains(f.Name, "sharedStrings.xml") {
			sharedStrFile = f
			break
		}
	}
	ss, err := sharedStrFile.Open()
	if err != nil {
		return err
	}
	defer ss.Close()
	idx := 0
	curFile, err := os.Create(READ_TEMP_DIR + baseFilename + "ss0")
	if err != nil {
		return err
	}
	curBuffer := bufio.NewWriter(curFile)
	decoder := xml.NewDecoder(ss)
	for {
		tok, _ := decoder.Token()
		if tok == nil {
			break
		}
		switch se := tok.(type) {
		case xml.StartElement:
			if se.Name.Local == "t" {
				val, _ := decoder.Token()
				str := val.(xml.CharData)
				curBuffer.WriteString("<t>" + string(str) + "</t>")
				idx++
				if idx%PARTITION_SIZE == 0 {
					curBuffer.Flush()
					curFile.Close()
					curFile, err = os.Create(READ_TEMP_DIR + baseFilename + "ss" + fmt.Sprintf("%d", idx/PARTITION_SIZE))
					if err != nil {
						return err
					}
					curBuffer = bufio.NewWriter(curFile)
				}
			}
			break
		default:
			break
		}
	}
	err = curBuffer.Flush()
	if err != nil {
		return err
	}
	err = curFile.Close()
	if err != nil {
		return err
	}
	return nil
}

type Column struct {
	IsString bool
	val      string
}

func (self *XlsxRowFetcher) NextRow() []string {

	for {
		tok, _ := self.Decoder.Token()
		if tok == nil {
			return nil
		}
		switch se := tok.(type) {
		case xml.StartElement:
			if se.Name.Local == "row" {
				fmt.Println("New Row")
				cols := []Column{}
				for {
					s, _ := self.Decoder.Token()
					if cc, ok := s.(xml.StartElement); ok {
						if cc.Name.Local == "c" {
							//fmt.Println("Col", cc.Attr)
							isString := false
							for _, kk := range cc.Attr {
								if kk.Name.Local == "t" && kk.Value == "s" {
									isString = true
								}
							}
							for {
								ss, _ := self.Decoder.Token()
								if cc2, ok := ss.(xml.StartElement); ok {
									if cc2.Name.Local == "v" {
										cont, _ := self.Decoder.Token()
										if cd, ok := cont.(xml.CharData); ok {
											if isString {
												//fmt.Println("CharData String", string(cd))
												cols = append(cols, Column{true, string(cd)})
											} else {
												//fmt.Println("CharData", string(cd))
												cols = append(cols, Column{false, string(cd)})
											}
										}
										break
									}
								}
							}
						}
					}
					if cc, ok := s.(xml.EndElement); ok {
						if cc.Name.Local == "row" {
							strCols := []string{}
							for _, c := range cols {
								if c.IsString {
									idx, _ := strconv.Atoi(c.val)
									//fmt.Println(idx)
									if self.IsUsingRamCache {
										c.val = self.SeekString(idx)
									} else {
										c.val = SeekString(self.Filename, idx)
									}

								}
								strCols = append(strCols, c.val)
							}
							//fmt.Println(strCols)
							//break
							return strCols
						}
					}
				}
			}
		}
	}
	return nil
}

var SheetNotFoundError = errors.New("Sheet Not Found")

func Import(filename string, sheetname string) (*XlsxRowFetcher, error) {
	res := &XlsxRowFetcher{}
	res.Filename = filepath.Base(filename)
	xlsxFile, err := zip.OpenReader(filename)
	if err != nil {
		return nil, err
	}
	res.ZipFile = xlsxFile
	var curSheet *zip.File
	for _, f := range xlsxFile.File {
		if strings.HasSuffix(f.Name, sheetname+".xml") {
			curSheet = f
			break
		}
	}
	if curSheet == nil {
		return nil, SheetNotFoundError
	}
	file, err := curSheet.Open()
	if err != nil {
		return nil, err
	}
	//defer file.Close()
	res.CurSheet = file
	decoder := xml.NewDecoder(file)
	res.Decoder = decoder
	PartitionSharedString(filename)
	return res, nil
}
