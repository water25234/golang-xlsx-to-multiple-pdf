package cmd

import (
	"fmt"
	"log"
	"time"

	"github.com/tealeg/xlsx"
	"github.com/water25234/golang-xlsx-to-multiple-pdf/mail"
)

// SendMail
func (fs *flags) SendMail() (err error) {

	mail.OAuthGmailService()

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

	fmt.Println("--------------- start send email work ---------------")

	for _, row := range rows {
		toEmail := row.Cells[7].String()
		pdfName := fmt.Sprintf("%s-%s.pdf", row.Cells[0].String(), row.Cells[1].String())
		_, err = mail.SendEmailOAUTH2(toEmail, fs.folder, pdfName)
		if err != nil {
			log.Println(err)
		}

		log.Println(fmt.Sprintf("Email sent successfully email: %s", toEmail))

		time.Sleep(1 * time.Second)
	}

	fmt.Println("--------------- finish send email work ---------------")
	return nil
}
