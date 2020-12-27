package cmd

import (
	"fmt"
	"log"

	"github.com/water25234/golang-xlsx-to-multiple-pdf/mail"
)

// SendMail
func (fs *flags) SendMail() (err error) {

	mail.OAuthGmailService()

	// xlFile, err := xlsx.OpenFile(fs.readfile)
	// if err != nil {
	// 	return err
	// }

	// if xlFile.Sheets == nil {
	// 	return fmt.Errorf("sheets is empty")
	// }

	// // only get first sheet
	// sheet := xlFile.Sheets[0]
	// // remove title
	// rows := sheet.Rows[1:]

	// if rows == nil {
	// 	return fmt.Errorf("sheets rows is empty")
	// }

	// row := rows[0]

	toEmail := "123@gmail.com"

	fmt.Println("--------------- start send email work ---------------")

	data := struct {
		ReceiverName string
		SenderName   string
	}{
		ReceiverName: "David Gilmour",
		SenderName:   "Binod Kafle",
	}

	_, err = mail.SendEmailOAUTH2(toEmail, data)
	if err != nil {
		log.Println(err)
	}

	log.Println(fmt.Sprintf("Email sent successfully email: %s", toEmail))

	fmt.Println("--------------- finish send email work ---------------")
	return nil
}
