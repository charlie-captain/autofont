package main

import (
	"baliance.com/gooxml/color"
	"baliance.com/gooxml/document"
	"baliance.com/gooxml/measurement"
	"baliance.com/gooxml/schema/soo/wml"
	"fmt"
	"github.com/sirupsen/logrus"
	"github.com/urfave/cli/v2"
	"math"
	"os"
)

func main() {

	app := &cli.App{
		Name:            "AutoFont",
		Version:         "1.0.0",
		Usage:           "",
		HideHelp:        false,
		HideHelpCommand: false,
		Flags: []cli.Flag{&cli.StringFlag{Name: "input", Aliases: []string{"i"}, Usage: "input file"},
			&cli.StringFlag{Name: "output", Aliases: []string{"o"}, Usage: "output file"}},
		Action: func(context *cli.Context) error {
			inputPath := context.String("input")
			outputPath := context.String("output")
			if inputPath != "" && outputPath != "" {
				return handle(inputPath, outputPath)
			}
			return context.App.Command("help").Run(context)
		},
	}

	err := app.Run(os.Args)
	if err != nil {
		fmt.Println(err)
		return
	}

}

func handle(input string, output string) error {
	inputFile, err := document.Open(input)
	if err != nil {
		logrus.Error(err)
	}
	p := inputFile.Paragraphs()
	text := ""
	for _, paragraph := range p {
		for _, line := range paragraph.Runs() {
			if len(line.Text()) > 0 {
				text += line.Text()
				break
			}
		}
	}
	charArray := []rune(text)
	newDoc := document.New()
	newDoc.Settings.SetUpdateFieldsOnOpen(true)
	table := newDoc.AddTable()
	pp := table.Properties()
	pp.SetAlignment(wml.ST_JcTableCenter)
	pp.Borders().SetAll(wml.ST_BorderSingle, color.Red, measurement.Point*0.5)
	textCount := len(charArray)
	// 4a纸大小
	var pageHeight float64 = 24
	// 一页可以放多少行
	var pageRow = int(pageHeight / 1.5)
	columnCount := 6
	var pageTextCount = pageRow * columnCount
	page := int(math.Ceil(float64(textCount) / float64(pageTextCount)))
	rowCount := page * pageRow
	rows := []document.Row{}
	cells := [][]document.Cell{}
	for row := 0; row < rowCount; row++ {
		row := table.AddRow()
		row.Properties().SetHeight(measurement.Centimeter*1.5, wml.ST_HeightRuleExact)
		rows = append(rows, row)
		perCells := []document.Cell{}
		for column := 0; column < columnCount; column++ {
			cell := row.AddCell()
			cell.Properties().SetWidth(measurement.Centimeter * 1.5)
			cell.Properties().SetVerticalAlignment(wml.ST_VerticalJcCenter)
			perCells = append(perCells, cell)
		}
		cells = append(cells, perCells)
	}
	index := 0
	for p := 1; p <= page; p++ {
		maxRow := pageRow
		for column := columnCount - 1; column >= 0; column-- {
			for row := 0; row < maxRow; row++ {
				if index > textCount-1 {
					break
				}
				pp := row + (p-1)*maxRow
				rowCells := cells[pp]
				curCell := rowCells[column]
				p1 := curCell.AddParagraph()
				p1.Properties().SetAlignment(wml.ST_JcCenter)
				char := charArray[index]
				r := p1.AddRun()
				r.Properties().SetSize(14)
				r.AddText(string(char))
				index++
			}
		}
	}

	err = newDoc.SaveToFile(output)
	if err != nil {
		return err
	}
	fmt.Println("导出成功")
	return nil
}
