package cmd

import (
	"bytes"
	"fmt"
	"html/template"
	"io/ioutil"
	"log"
	"os"
	"os/exec"
	"runtime"
	"strconv"
	"sync"
	"time"

	"github.com/SebastiaanKlippert/go-wkhtmltopdf"
	"github.com/tealeg/xlsx"
)

var (
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

type jobChannel struct {
	index       int
	fileContent *generatePDF
}

func (fs *flags) generateMultiplePdf() error {
	fmt.Println("--------------- start work ---------------")

	xlFile, err := xlsx.OpenFile(fs.readfile)
	if err != nil {
		return err
	}

	if xlFile.Sheets == nil {
		return err
	}

	sheet := xlFile.Sheets[0]
	rows := sheet.Rows[1:]
	// channel for job
	jobChans := make(chan jobChannel, len(rows))
	os.Mkdir(fs.folder, os.ModePerm)

	wg := &sync.WaitGroup{}
	wg.Add(len(rows))

	// start workers
	for i := 1; i <= runtime.NumCPU(); i++ {
		go func(i int) {
			for job := range jobChans {
				fs.work(job)
				wg.Done()
			}
		}(i)
	}

	// collect job
	for i, row := range rows {
		jobChans <- jobChannel{
			index: i,
			fileContent: &generatePDF{
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
			},
		}
	}

	close(jobChans)

	wg.Wait()

	fmt.Println("--------------- finish work ---------------")
	return nil
}

// work
func (fs *flags) work(job jobChannel) {

	err := job.fileContent.parseTemplate(templatePath)
	if err != nil {
		log.Fatal(err)
	}

	outputPath := fmt.Sprintf("%s/%s%s", fs.folder, job.fileContent.pdfName, job.fileContent.extName)

	if job.fileContent.generatePDF(outputPath) != nil {
		log.Fatal(err)
	}

	if fs.protection == true && job.fileContent.protectionPDF(outputPath) != nil {
		log.Fatal(err)
	}

	os.Remove(job.fileContent.tempHTML)
	fmt.Println(fmt.Sprintf("%s is success", fmt.Sprintf("%s%s", job.fileContent.pdfName, job.fileContent.extName)))
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
