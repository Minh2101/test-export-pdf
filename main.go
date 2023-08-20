package main

import (
	"fmt"
	"io"
	"log"
	"net/http"
	"os"
	"os/exec"
	"path/filepath"
)

type TemplateData struct {
	Name string `json:"name"`
}

func ExportWordToPDF(wordFilePath, pdfOutputDir string) (string, error) {
	libreofficePath := "C:\\Program Files\\LibreOffice\\program\\swriter.exe"
	if _, err := os.Stat(libreofficePath); os.IsNotExist(err) {
		return "", fmt.Errorf("LibreOffice executable not found")
	}

	pdfFilePath := filepath.Join(pdfOutputDir, "test.pdf")

	cmd := exec.Command(libreofficePath, "--convert-to", "pdf", "--outdir", pdfOutputDir, wordFilePath)
	output, err := cmd.CombinedOutput()
	if err != nil {
		return "", fmt.Errorf("Error exporting Word to PDF: %s\nOutput: %s", err, string(output))
	}

	return pdfFilePath, nil
}

func HandleConvertAndDownload(w http.ResponseWriter, r *http.Request) {
	if r.Method != http.MethodPost {
		http.Error(w, "Method not allowed", http.StatusMethodNotAllowed)
		return
	}

	wordFilePath := "./test.docx"
	pdfOutputDir := "./"

	//template, err := docxt.OpenTemplate(wordFilePath)
	//if err != nil {
	//	fmt.Println(err)
	//	return
	//}
	//
	//data := new(TemplateData)
	//data = &TemplateData{
	//	Name: "tessttttttt",
	//}
	//if err = template.RenderTemplate(data); err != nil {
	//	fmt.Println(err)
	//	return
	//}
	//
	//if err = template.Save("result.docx"); err != nil {
	//	fmt.Println(err)
	//	return
	//}

	pdfFilePath, err := ExportWordToPDF(wordFilePath, pdfOutputDir)
	if err != nil {
		http.Error(w, "Error exporting Word to PDF", http.StatusInternalServerError)
		return
	}

	pdfFile, err := os.Open(pdfFilePath)
	if err != nil {
		http.Error(w, "Error opening PDF file", http.StatusInternalServerError)
		return
	}
	defer func() {
		if err = os.Remove(pdfFilePath); err != nil {
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

func main() {
	http.HandleFunc("/export-pdf", HandleConvertAndDownload)

	fmt.Println("Server is listening on port 8080...")
	http.ListenAndServe(":8080", nil)
}
