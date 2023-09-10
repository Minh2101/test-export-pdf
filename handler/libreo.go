package handler

import (
	"fmt"
	"io"
	"log"
	"net/http"
	"os"
	"os/exec"
	"strings"
	"text/template"
)

type TemplateData struct {
	Name    string
	Mail    string
	Age     int64
	TestArr []int64
}

func convertDocxToHtml(docxFilePath string) string {
	libreOfficePath := "C:\\Program Files\\LibreOffice\\program\\soffice.exe"

	cmd := exec.Command(libreOfficePath, "--headless", "--convert-to", "html", "--outdir", ".", docxFilePath)
	cmd.Stdout = os.Stdout
	cmd.Stderr = os.Stderr

	if err := cmd.Run(); err != nil {
		log.Fatalf("Error converting DOCX to HTML: %v", err)
	}

	htmlFilePath := strings.ReplaceAll(docxFilePath, "docx", "html")

	return htmlFilePath
}
func parseDataToHtml(htmlFilePath string) string {
	arr := []int64{1, 2, 3, 4, 5, 6, 7, 8, 9, 0, 0, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 2, 2, 22, 2, 2, 2, 2, 2, 2, 2, 2, 3, 4, 4, 5, 5, 5, 6, 6, 6}
	data := TemplateData{
		Name:    "<b><em>Minh</em></b>",
		Mail:    "abc@gmail.com",
		Age:     3,
		TestArr: arr,
	}

	tmpl, err := template.ParseFiles(htmlFilePath)
	if err != nil {
		panic(err)
	}

	templateFilePath := "result.html"
	outputFile, err := os.Create(templateFilePath)
	if err != nil {
		panic(err)
	}
	defer outputFile.Close()

	err = tmpl.Execute(outputFile, data)
	if err != nil {
		panic(err)
	}

	return templateFilePath
}

func ExportHTMLToPDF(htmlFilePath string) error {
	libreofficePath := "C:\\Program Files\\LibreOffice\\program\\soffice.exe"
	if _, err := os.Stat(libreofficePath); os.IsNotExist(err) {
		return fmt.Errorf("LibreOffice executable not found")
	}

	cmd := exec.Command(libreofficePath, "--headless", "--convert-to", "pdf", "--outdir", ".", htmlFilePath)
	output, err := cmd.CombinedOutput()
	if err != nil {
		return fmt.Errorf("Error exporting HTML to PDF: %s\nOutput: %s", err, string(output))
	}

	return nil
}
func HandleConvertAndDownload(w http.ResponseWriter, r *http.Request) {
	if r.Method != http.MethodPost {
		http.Error(w, "Method not allowed", http.StatusMethodNotAllowed)
		return
	}

	wordFilePath := "test.docx"
	pdfOutputDir := "result.pdf"

	htmlFilePath := convertDocxToHtml(wordFilePath)
	defer func() {
		if err := os.Remove(htmlFilePath); err != nil {
			log.Printf("Error removing temporary PDF file: %s", err)
		}
	}()
	resultPath := parseDataToHtml(htmlFilePath)
	defer func() {
		if err := os.Remove(resultPath); err != nil {
			log.Printf("Error removing temporary PDF file: %s", err)
		}
	}()
	err := ExportHTMLToPDF(resultPath)
	if err != nil {
		http.Error(w, "Error exporting Word to PDF", http.StatusInternalServerError)
		return
	}

	pdfFile, err := os.Open(pdfOutputDir)
	if err != nil {
		http.Error(w, "Error opening PDF file", http.StatusInternalServerError)
		return
	}
	defer func() {
		if err = os.Remove(pdfOutputDir); err != nil {
			log.Printf("Error removing temporary PDF file: %s", err)
		}
	}()
	defer pdfFile.Close()

	w.Header().Set("Content-Type", "application/pdf")
	w.Header().Set("Content-Disposition", "attachment; filename=output.pdf")

	_, err = io.Copy(w, pdfFile)
	if err != nil {
		http.Error(w, "Error sending PDF content", http.StatusInternalServerError)
		return
	}

	fmt.Println("Converted and sent PDF successfully")
}
