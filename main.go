package main

import (
	"embed"
	"excel2img/lib"
	"fmt"
	"github.com/spf13/cobra"
	"github.com/xuri/excelize/v2"
	"log"
	"strings"
)

var rootCmd = &cobra.Command{
	Use:     "excel2img",
	Short:   "excel2img:  {excelPath} {output}",
	Long:    "excel2img:  {excelPath} {output}",
	Version: "0.0.1",
	Run:     drawExcelToPng,
}

//go:embed fonts
var fonts embed.FS

func main() {
	err := lib.Init(fonts)
	if err != nil {
		panic(err)
	}
	if err := rootCmd.Execute(); err != nil {
		log.Fatal(err)
	}
}

func drawExcelToPng(cmd *cobra.Command, args []string) {
	if len(args) != 2 {
		log.Fatal("please input {excelPath} {output}")
	}
	e2i := lib.Ex2Img{}
	excelFile := args[0]
	output := args[1]
	file, err := excelize.OpenFile(excelFile)
	if err != nil {
		log.Fatal(err)
		return
	}
	if !strings.HasSuffix(output, ".png") && !strings.HasSuffix(output, ".PNG") {
		output = fmt.Sprintf("%s.png", output)
	}
	e2i.DrawExcelToPngFile(file, output)
}
