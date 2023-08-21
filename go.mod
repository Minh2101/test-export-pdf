module test-export-pdf

go 1.18

replace (
	"github.com/go-ole/go-ole" => "./pkg/go-ole"
	"github.com/go-ole/go-ole/oleutil" => "./pkg/go-ole/oleutil"
)

require github.com/go-ole/go-ole v1.3.0

require golang.org/x/sys v0.1.0 // indirect
