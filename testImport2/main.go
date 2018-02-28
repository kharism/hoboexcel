package main

import (
	"fmt"
	"os"

	ee "github.com/kharism/hoboexcel"
)

func main() {
	//set up import environment variable
	ee.READ_TEMP_DIR = "./temp/"
	ee.PARTITION_SIZE = 300

	XlsxRowFetcher, err := ee.Import("./Book1.xlsx", "sheet1")
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
		fmt.Println(cols)
	}
}
