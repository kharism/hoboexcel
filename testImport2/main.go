package main

import (
	"fmt"
	"os"
	"strings"

	ee "github.com/eaciit/hoboexcel"
)

func main() {
	//set up import environment variable
	ee.READ_TEMP_DIR = "./temp/"
	ee.PARTITION_SIZE = 300

	XlsxRowFetcher, err := ee.Import("./Book1.xlsx", "Sheet1")
	//don't forget to use ram cache if you can afford it, it really helps in some cases, otherwise turn it off
	XlsxRowFetcher.IsUsingRamCache = true
	if err != nil {
		fmt.Println(err.Error())
		os.Exit(1)
	}
	fmt.Println("Start Fetching")
	for {
		cols := XlsxRowFetcher.NextRow()
		if cols == nil {
			XlsxRowFetcher.Close()
			break
		}
		fmt.Println(strings.Join(cols, "|"))
	}
}
