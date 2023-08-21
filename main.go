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
