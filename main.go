package main

import (
	"encoding/csv"
	"flag"
	"fmt"
	"os"

	"github.com/360EntSecGroup-Skylar/excelize"
)

func main() {
	flag.Parse()
	args := flag.Args()
	if len(args) != 2 {
		fmt.Printf("syntax: tsv2xslx INPUT.tsv OUTPUT.xslx (%d arguments given)\n", len(args))
		return
	}

	// Open file
	fp, err := os.Open(args[0])
	if err != nil {
		panic(err)
	}
	defer fp.Close()
	r := csv.NewReader(fp)
	r.Comma = '\t'
	r.LazyQuotes = true

	// Read header
	_, err = r.Read()
	if err != nil {
		panic(err)
	}

	xlsx := excelize.NewFile()

	// Iterate
	i := 0
	record, err := r.Read()
	for err == nil {
		i++
		for k, v := range record {
			//fmt.Printf("[%d][%v] => %s\n", k, v, fmt.Sprintf("%s%d", excelize.ToAlphaString(k), i))
			xlsx.SetCellValue("Sheet1", fmt.Sprintf("%s%d", excelize.ToAlphaString(k), i), v)
		}
		record, err = r.Read()
	}
	fmt.Printf("%d records processed\n", i-1)
	if err.Error() != "EOF" {
		fmt.Printf("Terminating error: %s\n", err.Error())
	}

	// Set active sheet of the workbook.
	xlsx.SetActiveSheet(0)

	// Save xlsx file by the given path.
	err = xlsx.SaveAs(args[1])
	if err != nil {
		panic(err)
	}
}
