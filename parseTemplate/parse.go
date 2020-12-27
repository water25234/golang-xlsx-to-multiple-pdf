package parse

import (
	"bytes"
	"fmt"
	"path/filepath"
	"text/template"
)

// Template means parse template
func Template(templateFileName string, data interface{}) (templateByte []byte, err error) {
	templatePath, err := filepath.Abs(templateFileName)
	if err != nil {
		return nil, fmt.Errorf("invalid template name")
	}

	t, err := template.ParseFiles(templatePath)
	if err != nil {
		return nil, err
	}

	buf := new(bytes.Buffer)
	if err = t.Execute(buf, data); err != nil {
		return nil, err
	}

	return buf.Bytes(), nil
}
