package main

import (
	"bytes"
	"fmt"
	"html/template"
	"io/ioutil"
	"log"
	"os"
	"os/exec"
	"strconv"
	"time"

	"github.com/SebastiaanKlippert/go-wkhtmltopdf"
	"github.com/tealeg/xlsx"
)

var (
	inFile       = "example-excel.xlsx"
	templatePath = "templateHtml/templateHtml.html"
)

type generatePDF struct {
	*requestPdf
	body     []byte
	extName  string
	pdfName  string
	tempHTML string
}

type requestPdf struct {
	Name           string
	EmployeeNumber string
	Gender         string
	Education      string
	Nationality    string
	Password       string
}

func main() {
	xlFile, err := xlsx.OpenFile(inFile)
	if err != nil {
		fmt.Println(err.Error())
		return
	}

	if xlFile.Sheets == nil {
		fmt.Println(err.Error())
		return
	}

	sheet := xlFile.Sheets[0]
	rows := sheet.Rows[1:]
	for _, row := range rows {
		gp := &generatePDF{
			// To be optimized...
			requestPdf: &requestPdf{
				Name:           row.Cells[0].String(),
				EmployeeNumber: row.Cells[1].String(),
				Gender:         row.Cells[2].String(),
				Education:      row.Cells[3].String(),
				Nationality:    row.Cells[4].String(),
				Password:       row.Cells[5].String(),
			},
			extName:  ".pdf",
			pdfName:  fmt.Sprintf("%s-%s", row.Cells[0].String(), row.Cells[1].String()),
			tempHTML: fmt.Sprintf("pdfGenerator/cloneTemplate/%s.html", strconv.FormatInt(int64(time.Now().Unix()), 10)),
		}
		err := gp.parseTemplate(templatePath)
		if err != nil {
			log.Fatal(err)
		}

		defer os.Remove(gp.tempHTML)
		outputPath := fmt.Sprintf("pdfFolder/%s%s", gp.pdfName, gp.extName)

		if gp.generatePDF(outputPath) != nil {
			log.Fatal(err)
		}

		if gp.protectionPDF(outputPath) != nil {
			log.Fatal(err)
		}

		fmt.Println(fmt.Sprintf("%s is success", fmt.Sprintf("%s%s", gp.pdfName, gp.extName)))
	}
	fmt.Println("import success")
}

//parsing template
func (gp *generatePDF) parseTemplate(templateFileName string) (err error) {

	t, err := template.ParseFiles(templateFileName)
	if err != nil {
		return err
	}
	buf := new(bytes.Buffer)
	if err = t.Execute(buf, gp.requestPdf); err != nil {
		return err
	}

	gp.body = buf.Bytes()
	return nil
}

//generatePDF means generate pdf
func (gp *generatePDF) generatePDF(pdfPath string) (err error) {
	err1 := ioutil.WriteFile(gp.tempHTML, gp.body, 0644)
	if err1 != nil {
		return err
	}

	tempHTMLFile, err := os.Open(gp.tempHTML)
	if tempHTMLFile != nil {
		defer tempHTMLFile.Close()
	}
	if err != nil {
		return err
	}

	wkhtmltopdf.SetPath("pdfGenerator/wkhtmltopdf")
	pdfg, err := wkhtmltopdf.NewPDFGenerator()
	if err != nil {
		return err
	}

	pdfg.AddPage(wkhtmltopdf.NewPageReader(tempHTMLFile))

	pdfg.PageSize.Set(wkhtmltopdf.PageSizeA4)

	pdfg.Dpi.Set(300)

	err = pdfg.Create()
	if err != nil {
		return err
	}

	err = pdfg.WriteFile(pdfPath)
	if err != nil {
		return err
	}
	return nil
}

//protectionPDF means provider password pdf
func (gp *generatePDF) protectionPDF(pdfPath string) (err error) {
	cmd := exec.Command("pdfcpuGenerator/pdfcpu", "encrypt", "-upw", gp.Password, "-opw", gp.Password, pdfPath)
	err = cmd.Run()
	if err != nil {
		return err
	}
	return nil
}
