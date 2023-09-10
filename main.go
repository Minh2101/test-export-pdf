package main

import (
	"fmt"
	"net/http"
	"test-export-pdf/handler"
)

func main() {
	http.HandleFunc("/export-pdf", handler.HandleConvertAndDownload)
	http.HandleFunc("/export-pdf-ole", handler.ExportPDF)

	fmt.Println("Server is listening on port 8080...")
	http.ListenAndServe(":8080", nil)
}

//package main
//
//import (
//	"fmt"
//	"log"
//	"os"
//	"os/exec"
//)
//
//func main() {
//	// Define the input DOCX file path.
//	docxFilePath := "test.docx"
//
//	// Define the output HTML file path.
//	htmlFilePath := "output.html"
//
//	// Specify the full path to the Pandoc executable (pandoc.exe).
//	pandocPath := "C:\\Users\\nguye\\AppData\\Local\\Pandoc\\pandoc.exe" // Replace with the actual path to pandoc.exe
//
//	// Use Pandoc to convert DOCX to HTML.
//	cmd := exec.Command(pandocPath, docxFilePath, "-o", htmlFilePath)
//	cmd.Stdout = os.Stdout
//	cmd.Stderr = os.Stderr
//
//	if err := cmd.Run(); err != nil {
//		log.Fatalf("Error converting DOCX to HTML: %v", err)
//	}
//
//	fmt.Printf("DOCX file '%s' converted to HTML and saved as '%s'\n", docxFilePath, htmlFilePath)
//}
