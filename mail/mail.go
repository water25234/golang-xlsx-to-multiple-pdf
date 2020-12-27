package mail

import (
	"context"
	"crypto/rand"
	"encoding/base64"
	"errors"
	"fmt"
	"io/ioutil"
	"log"
	"net/http"
	"os"
	"time"

	parse "github.com/water25234/golang-xlsx-to-multiple-pdf/parseTemplate"
	"golang.org/x/oauth2"
	"golang.org/x/oauth2/google"
	"google.golang.org/api/gmail/v1"
	"google.golang.org/api/option"
)

// GmailService : Gmail client for sending email
var (
	GmailService *gmail.Service

	template = "templateHtml/mail/templateHtml.html"
)

// OAuthGmailService means
func OAuthGmailService() {
	config := oauth2.Config{
		ClientID:     os.Getenv("CLIENT_ID"),
		ClientSecret: os.Getenv("CLIENT_SECRET"),
		Endpoint:     google.Endpoint,
		RedirectURL:  "http://localhost",
	}

	token := oauth2.Token{
		AccessToken:  os.Getenv("ACCESS_TOKEN"),
		RefreshToken: os.Getenv("REFRESH_TOKEN"),
		TokenType:    "Bearer",
		Expiry:       time.Now(),
	}

	var tokenSource = config.TokenSource(context.Background(), &token)

	srv, err := gmail.NewService(context.Background(), option.WithTokenSource(tokenSource))
	if err != nil {
		log.Fatalf("Unable to retrieve Gmail client: %v", err)
	}

	GmailService = srv
	if GmailService != nil {
		fmt.Println("Email service is initialized \n")
	}
}

// SendEmailOAUTH2 means
func SendEmailOAUTH2(to string, data interface{}) (bool, error) {

	emailBodyByte, err := parse.Template(template, data)
	if err != nil {
		return false, errors.New("unable to parse email template")
	}

	var message gmail.Message

	filename := "林黛玉-KK88892.pdf"
	fileBytes, err := ioutil.ReadFile(fmt.Sprintf("%s/%s", "pdfFolder", "林黛玉-KK88892.pdf"))
	if err != nil {
		log.Fatalf("Error: %v", err)
	}

	boundary := randStr(32, "alphanum")

	fileMIMEType := http.DetectContentType(fileBytes)

	fileData := base64.StdEncoding.EncodeToString(fileBytes)

	content := string(emailBodyByte)

	msg := []byte("Content-Type: multipart/mixed; boundary=" + boundary + " \n" +
		"MIME-Version: 1.0\n" +
		"to: " + to + "\n" +
		"subject: " + "Test Email form Gmail API using OAuth" + "\n\n" +

		"--" + boundary + "\n" +
		"Content-Type: text/plain; charset=" + string('"') + "UTF-8" + string('"') + "\n" +
		"MIME-Version: 1.0\n" +
		"Content-Transfer-Encoding: 7bit\n\n" +
		content + "\n\n" +
		"--" + boundary + "\n" +

		"Content-Type: " + fileMIMEType + "; name=" + string('"') + filename + string('"') + " \n" +
		"MIME-Version: 1.0\n" +
		"Content-Transfer-Encoding: base64\n" +
		"Content-Disposition: attachment; filename=" + string('"') + filename + string('"') + " \n\n" +
		chunkSplit(fileData, 76, "\n") +
		"--" + boundary + "--")

	message.Raw = base64.URLEncoding.EncodeToString(msg)

	// Send the message
	_, err = GmailService.Users.Messages.Send("me", &message).Do()
	if err != nil {
		return false, err
	}
	return true, nil
}

func chunkSplit(body string, limit int, end string) string {
	var charSlice []rune

	// push characters to slice
	for _, char := range body {
		charSlice = append(charSlice, char)
	}

	var result = ""

	for len(charSlice) >= 1 {
		// convert slice/array back to string
		// but insert end at specified limit
		result = result + string(charSlice[:limit]) + end

		// discard the elements that were copied over to result
		charSlice = charSlice[limit:]

		// change the limit
		// to cater for the last few words in
		if len(charSlice) < limit {
			limit = len(charSlice)
		}
	}
	return result
}

func randStr(strSize int, randType string) string {

	var dictionary string

	if randType == "alphanum" {
		dictionary = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"
	}

	var strBytes = make([]byte, strSize)
	_, _ = rand.Read(strBytes)
	for k, v := range strBytes {
		strBytes[k] = dictionary[v%byte(len(dictionary))]
	}
	return string(strBytes)
}
