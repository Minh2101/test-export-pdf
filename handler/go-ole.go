package handler

import (
	"fmt"
	"io"
	"log"
	"net/http"
	"os"

	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

func ExportPDF(w http.ResponseWriter, r *http.Request) {
	// Initialize OLE
	ole.CoInitializeEx(0, ole.COINIT_APARTMENTTHREADED|ole.COINIT_SPEED_OVER_MEMORY)
	defer ole.CoUninitialize()

	// Create a new instance of Word
	unknown, err := oleutil.CreateObject("Word.Application")
	if err != nil {
		log.Fatal(err)
	}
	defer unknown.Release()

	word, err := unknown.QueryInterface(ole.IID_IDispatch)
	if err != nil {
		log.Fatal(err)
	}
	defer word.Release()

	// Set Word to be visible (optional)
	oleutil.PutProperty(word, "Visible", true)

	// Get the 'Documents' object
	docs := oleutil.MustGetProperty(word, "Documents").ToIDispatch()

	// Open the Word document
	doc, err := oleutil.CallMethod(docs, "Open", "C:/Users/nguye/Downloads/test-export-pdf/test.docx")
	if err != nil {
		log.Fatal(err)
	}
	defer doc.Clear()
	defer oleutil.CallMethod(word, "Quit")

	pdfFilePath := "C:/Users/nguye/Downloads/test-export-pdf/testPDF.pdf"
	// Export the document to PDF
	_, err = oleutil.CallMethod(doc.ToIDispatch(), "SaveAs2", pdfFilePath, 17) // 17 for PDF format
	if err != nil {
		log.Fatal(err)
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

	fmt.Println("Exported to PDF successfully")
}
