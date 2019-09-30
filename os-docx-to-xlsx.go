//main is main package in os-docx-to-xsl project
package main

import (
	"flag"
	"fmt"
	"io/ioutil"

	"os"
	"path/filepath"
	"unicode/utf8"

	log "github.com/sirupsen/logrus"
	"github.com/unidoc/unioffice/document"
)

const ftdIncomAttrErr = "Incoming arguments are invalid"
const ftdIncomAttrErr1 = "Incoming arguments are invalid %s - \"%v\""
const ftdIncomAttrErr2 = "Incoming arguments are invalid  %s - \"%v\",  %s - \"%v\""
const ftdIncomAttrErr3 = "Incoming arguments are invalid  %s - \"%v\",  %s - \"%v\",  %s - \"%v\""
const ftdAttrShouldBeInfo = "Attributes should be %v"
const errorFoundReturn = "error found exit with return "

func init() {
	// Log as JSON instead of the default ASCII formatter.
	log.SetFormatter(&log.JSONFormatter{})

	// Output to stdout instead of the default stderr
	// Can be any io.Writer, see below for File example
	log.SetOutput(os.Stdout)

	// Only log the warning severity or above.
	log.SetLevel(log.WarnLevel)

	// calling method as a field, instruct the logger
	log.SetReportCaller(true)

}
func main() {
	fdirptr := flag.String("fdirpath", ".", "file path to read file")
	flag.Parse()
	dirName := *fdirptr
	files, err := ioutil.ReadDir(dirName)
	if err != nil {
		log.Fatal(err)
	}
	var totalList []string
	for _, f := range files {
		callPattern := fmt.Sprintf("%s%c%s%c%s",
			dirName,
			os.PathSeparator,
			f.Name(),
			os.PathSeparator,
			"Заявка*.docx")
		resultNames, err := ReturnNames(callPattern)
		if err != nil {
			log.Fatal(err)
		}

		if len(resultNames) > 0 {
			totalList = append(totalList, resultNames...)
			// for _, names := range resultNames {
			// 	fmt.Printf("\"%s\" ", names)
			// }

		}

	}
	// fmt.Printf("%q", totalList)
	fmt.Println("Total number of items is ", len(totalList))
	var recordsCollection []string
	for _, zaya := range totalList {
		doc, err := document.Open(zaya)
		if err != nil {
			log.Fatalf("error opening document: %s", err)
		}
		var doRecord bool = false

		var record string
		for _, tbl := range doc.Tables() {
			for _, row := range tbl.Rows() {
				for _, cell := range row.Cells() {
					// fmt.Println("cell")
					for _, p := range cell.Paragraphs() {
						// fmt.Println("paragraph")
						for _, r := range p.Runs() {
							// fmt.Println(r.Text())
							switch r.Text() {
							case "Фамилия", "Имя", "e-mail", "Номер телефона":
								doRecord = true
							case " ":
								if doRecord {
									record = record + "; "
								}
								doRecord = false

							default:
								if doRecord {
									record = record + r.Text()
								}
							}
						}
					}
				}
			}
		}
		recordsCollection = append(recordsCollection, record)

	}
	fmt.Printf("%q", recordsCollection)
	log.Trace("Create file start")
	f, err := os.Create("list.csv")
	if err != nil {
		log.Error(err)
		log.Traceln(errorFoundReturn, err)
	}
	log.Trace("Create file end")

	log.Trace("Start write to file the incoming string ...")
	for _, s := range recordsCollection {
		_, err := f.WriteString(s+";\n")
		if err != nil {
			log.Traceln(errorFoundReturn, err)
			log.Error(err)
		}
	}
	log.Traceln("close file start ...")
	err = f.Close()
	if err != nil {
		log.Traceln(errorFoundReturn, err)
		log.Error(err)
	}
	log.Traceln("close file end")
}

//ReturnNames return the names of all files that matches a patter
func ReturnNames(pattern string) (fileNames []string, err error) {
	if utf8.RuneCountInString(pattern) <= 0 {
		return nil, nil
	}

	filesList, error := filepath.Glob(pattern)
	if error != nil && error == filepath.ErrBadPattern {
		fmt.Println("Bad pattern -->", error)
		return nil, error
	}
	log.Tracef(fmt.Sprint(filesList))
	return filesList, nil
}

func getMultyArg(args ...interface{}) []interface{} {
	return args
}
