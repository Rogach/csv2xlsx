package main

import (
	"bufio"
	"fmt"
	"os"
	"strings"
	"unicode/utf8"
	"github.com/tealeg/xlsx"
	"gopkg.in/alecthomas/kingpin.v2"
)

var (
	separator = kingpin.Flag("separator", "separator used in input data (use \\t for tab)").Default(",").Short('s').String()
	outputFile = kingpin.Arg("destination", "destination file - leave empty to use stdout").String()
)

func handleError(err error) {
	if err != nil {
		fmt.Fprintln(os.Stderr, err)
		os.Exit(1)
	}
}

func main() {
	kingpin.Parse()

	if (*separator == "\\t") {
		*separator = "\t"
	}

	scanner := bufio.NewScanner(os.Stdin)
	defer os.Stdin.Close()

	var records [][]string


	for scanner.Scan() {
		line := scanner.Text()
		if len(line) != 0 {
			records = append(records, strings.Split(scanner.Text(), *separator))
		}
	}

	columnCount := 0
	for _, record := range records {
		if columnCount < len(record) {
			columnCount = len(record)
		}
	}

	workbook := xlsx.NewFile()
	sheet, err := workbook.AddSheet("Sheet1")
	handleError(err)

	// auto-fit columns
	for c := 0; c < columnCount; c++ {
		maxCellSize := 0
		for _, record := range records {
			if len(record) > c {
				l := utf8.RuneCountInString(record[c])
				if maxCellSize < l {
					maxCellSize = l
				}
			}
		}
		sheet.SetColWidth(c, c, float64(maxCellSize))
	}

	for _, record := range records {
		row := sheet.AddRow()
		for _, s := range record {
			cell := row.AddCell()
			cell.Value = s
		}
	}

	if (*outputFile == "" || *outputFile == "-") {
		err = workbook.Write(os.Stdout)
		handleError(err)
	} else {
		err = workbook.Save(*outputFile)
		handleError(err)
	}

}
