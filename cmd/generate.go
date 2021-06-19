package cmd

import (
	"bytes"
	"errors"
	"fmt"
	"html/template"
	"io/ioutil"
	"log"
	"os"
	"os/exec"
	"path/filepath"
	"runtime"
	"strconv"
	"sync"
	"time"

	"github.com/SebastiaanKlippert/go-wkhtmltopdf"
	"github.com/tealeg/xlsx"
)

var (
	extName = ".pdf"

	templatePath = "templateHtml/pdf/index-3.html"
)

type generatePDF struct {
	*requestPdf
	body     []byte
	pdfName  string
	tempHTML string
}

type requestPdf struct {
	EmployeeNumber string
	Name           string
	Department     string
	JobLevel       string
	JobGrade       string
	JobTitle       string
	Password       string
	Email          string
}

type jobChannel struct {
	index       int
	fileContent *generatePDF
}

func (fs *flags) GenerateMultiplePdf() error {
	fmt.Println("--------------- start work ---------------")

	xlFile, err := xlsx.OpenFile(fs.readfile)
	if err != nil {
		return err
	}

	if xlFile.Sheets == nil {
		return fmt.Errorf("sheets is empty")
	}

	// only get first sheet
	sheet := xlFile.Sheets[0]
	// remove title
	rows := sheet.Rows[1:]

	if rows == nil {
		return fmt.Errorf("sheets rows is empty")
	}

	// create folder
	os.MkdirAll(fs.folder, os.ModePerm)
	os.MkdirAll(fmt.Sprintf("%s-copy", fs.folder), os.ModePerm)

	// channel for job
	fs.jobs(rows)

	fmt.Println("--------------- finish work ---------------")
	return nil
}

// jobs
func (fs *flags) jobs(rows []*xlsx.Row) {
	jobChans := make(chan *jobChannel, len(rows))

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
		pdfName := fmt.Sprintf("%s-%s", row.Cells[0].String(), row.Cells[1].String())
		jobChans <- &jobChannel{
			index: i,
			fileContent: &generatePDF{
				// To be optimized...
				requestPdf: &requestPdf{
					EmployeeNumber: row.Cells[0].String(),
					Name:           row.Cells[1].String(),
					Department:     row.Cells[2].String(),
					JobLevel:       row.Cells[3].String(),
					JobGrade:       row.Cells[4].String(),
					JobTitle:       row.Cells[5].String(),
					Password:       row.Cells[6].String(),
					Email:          row.Cells[7].String(),
				},
				pdfName:  pdfName,
				tempHTML: fmt.Sprintf("pdfGenerator/cloneTemplate/%s-%s.html", pdfName, strconv.FormatInt(int64(time.Now().Unix()), 10)),
			},
		}
	}

	close(jobChans)

	wg.Wait()
}

// work
func (fs *flags) work(job *jobChannel) {

	defer func() {
		job = nil
	}()

	defer os.Remove(job.fileContent.tempHTML)

	err := job.fileContent.parseTemplate(templatePath)
	if err != nil {
		fmt.Println("parseTemplate failure" + templatePath)
		log.Fatal(err)
	}

	outputPath := fmt.Sprintf("%s/%s%s", fs.folder, job.fileContent.pdfName, extName)
	outputPathCopy := fmt.Sprintf("%s-copy/%s%s", fs.folder, job.fileContent.pdfName, extName)

	if job.fileContent.generatePDF(outputPath) != nil {
		fmt.Println("parseTemplate failure" + outputPath)
		log.Fatal(err)
	}

	if fs.protection == true && job.fileContent.protectionPDF(outputPath) != nil {
		fmt.Println("protection pdf failure" + outputPath)
		log.Fatal(err)
	}

	err = job.fileContent.copyFile(outputPath, outputPathCopy)
	if err != nil {
		fmt.Println("copy file")
		fmt.Println(err)
	}

	fmt.Println(fmt.Sprintf("%s is success", fmt.Sprintf("%s%s", job.fileContent.pdfName, extName)))
}

func (gp *generatePDF) copyFile(src, dst string) (err error) {
	//Read all the contents of the  original file
	bytesRead, err := ioutil.ReadFile(src)
	if err != nil {
		log.Fatal(err)
		return err
	}

	//Copy all the contents to the desitination file
	err = ioutil.WriteFile(dst, bytesRead, 0755)
	if err != nil {
		log.Fatal(err)
		return err
	}
	return nil
}

//parsing template
func (gp *generatePDF) parseTemplate(templateFileName string) (err error) {
	_, err = filepath.Abs(templateFileName)
	if err != nil {
		return errors.New("invalid template name")
	}

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
		fmt.Println("create pdf failure")
		fmt.Println(err)
		return err
	}

	err = pdfg.WriteFile(pdfPath)
	if err != nil {
		fmt.Println("WriteFile pdf failure")
		fmt.Println(err)
		return err
	}
	return nil
}

//protectionPDF means provider password pdf
func (gp *generatePDF) protectionPDF(pdfPath string) (err error) {
	cmd := exec.Command("pdfcpuGenerator/pdfcpu", "encrypt", "-upw", gp.Password, "-opw", gp.Password, pdfPath)
	err = cmd.Run()
	if err != nil {
		fmt.Println("exec protection pdf failure")
		fmt.Println(err)
		return err
	}
	return nil
}
