// BonusCertOverview keeps an eye on your Bonus Certificates
package main

import (
	"encoding/json"
	"fmt"
	"io/ioutil"
	"log"
	"math"
	"os"
	"strconv"
	"strings"
	"time"

	"github.com/rs/zerolog"
	"github.com/xuri/excelize/v2"
)

type ConfigValues struct {
	Language string
	DataFile string
}

const (
	_           = iota
	Bezeichnung // 1
	ISIN        // 2
	_
	EinkaufsWert       // 4
	Anzahl             // 5
	Spesen             // 6
	EinkaufswertGesamt // 7
	Einkaufsdatum
	EndeDatum
	Gesamtlaufzeit
	CAP
	ISINBasiswert
	Barriere
	_
	AktDatum
	AktWert
	AktWertBasiswert
	AktuellerErtragInProzent
	AktuellerErtragSumme
	ErwarteterErtragInProzent
	ErwarteterErtrag
	GesamtertragInProzentProJahr
	Restlaufzeit
	ErtragInProzentRestlaufzeit
	ErtragInProzentProJahr
)

func main() {

	// Define global variables
	activeDataSheet := "Laufende Aktionen"
	zerolog.TimeFieldFormat = time.RFC3339
	zerolog.SetGlobalLevel(zerolog.DebugLevel)
	timeStamp := time.Now().Format("20060102_150405")

	// Let's first read the `config.json` file
	content, err := ioutil.ReadFile("./config.json")
	if err != nil {
		log.Fatal("Error when opening file: ", err)
	}

	// Now let's unmarshall the data into `payload`
	var config ConfigValues
	err = json.Unmarshal(content, &config)
	if err != nil {
		log.Fatal("Error during Unmarshal(): ", err)
	}

	// Print the unmarshalled data!
	log.Printf("language: %s\n", config.Language)
	log.Printf("dataFile: %s\n", config.DataFile)

	// Create a log file
	posOfPoint := strings.LastIndex(config.DataFile, ".")
	logFileName := config.DataFile[0:posOfPoint] + "_" + timeStamp + ".log"
	logFile, err := os.Create(logFileName)

	if err != nil {
		log.Fatal(err)
	}

	defer logFile.Close()

	zLog := zerolog.New(logFile).With().Timestamp().Logger()

	// Read the Excel data file
	zLog.Info().Msg("read the BonusCertificate file " + config.DataFile)
	dataFile, err := excelize.OpenFile(config.DataFile)
	if err != nil {
		zLog.Fatal().Err(err).Msg("")
	}

	defer func() {
		// Close the spreadsheet.
		if err := dataFile.Close(); err != nil {
			zLog.Fatal().Err(err).Msg("")
		}
	}()

	// Get all the rows in the Sheet1.
	rows, err := dataFile.GetRows(activeDataSheet)
	if err != nil {
		zLog.Fatal().Err(err).Msg("")
	}

	// Get the number of rows
	numberOfRows := len(rows)
	fmt.Printf("Anzahl Rows: %v", numberOfRows)

	// Get the number of columns
	numberOfColumns := len(rows[0])
	fmt.Printf("Anzahl Columns: %v", numberOfColumns)

	// var AllCertificateData [numberOfColumns][numberOfRows]string
	// var AllCertificateData [100][100]string
	AllCertificateData := [100][100]string{}

	// Iterate over all culomns and rows and save the data into the slice AllCertificateData
	for currColumn := 1; currColumn <= numberOfColumns; currColumn++ {

		currColumnName, _ := excelize.ColumnNumberToName(currColumn)

		for currRow := 1; currRow <= numberOfRows; currRow++ {

			currAxis := currColumnName + strconv.Itoa(currRow)
			zwerg, _ := dataFile.GetCellType(activeDataSheet, currAxis)
			zwerg2, _ := dataFile.GetCellValue(activeDataSheet, currAxis)
			AllCertificateData[currColumn][currRow] = zwerg2
			fmt.Printf("Wert ist: %v\t", zwerg)
			fmt.Printf("Wert ist: %v\n", zwerg2)
		}
	}

	fmt.Printf("B5 ist %v", AllCertificateData[2][5])

	// Re-Calculate the values
	zLog.Debug().Msg("Re-Calculate the values")

	// Iterate over all columns
	zLog.Debug().Msg("Iterate over all columns")

	for currColumn := 2; currColumn <= numberOfColumns; currColumn++ {
		zLog.Debug().Msg("Work on column " + strconv.Itoa(currColumn))
		zLog.Debug().Msg("Work on column " + AllCertificateData[currColumn][Bezeichnung])

		zLog.Debug().Msg("Calculate GesamtEinkaufswert")

		currEinkaufswert, _ := strconv.ParseFloat(AllCertificateData[currColumn][EinkaufsWert], 32)
		currEinkaufswert = math.Round((currEinkaufswert * 100)) / 100
		currAnzahl, _ := strconv.ParseFloat(AllCertificateData[currColumn][Anzahl], 32)
		currSpesen, _ := strconv.ParseFloat(AllCertificateData[currColumn][Spesen], 32)

		currEinkaufswertGesamt := (currEinkaufswert * currAnzahl) + currSpesen
		currEinkaufswertGesamt = math.Round((currEinkaufswertGesamt * 100)) / 100

		AllCertificateData[currColumn][EinkaufswertGesamt] = fmt.Sprintf("%.2f", currEinkaufswertGesamt)

		zLog.Debug().Msg("GesamtEinkaufswert ist: " + AllCertificateData[currColumn][EinkaufswertGesamt])

		zLog.Debug().Msg("Calculate Laufzeit in Tagen")

		tmpEinkaufsDatum := strings.Replace(AllCertificateData[currColumn][Einkaufsdatum], "/", "-", -1)
		currEinkaufsDatum, _ := time.Parse("2-1-2006", tmpEinkaufsDatum)
		zLog.Debug().Msgf("Einkaufsdatum %v", currEinkaufsDatum)

		tmpEndeDatum := strings.Replace(AllCertificateData[currColumn][EndeDatum], "/", "-", -1)
		currEndeDatum, _ := time.Parse("2-1-2006", tmpEndeDatum)
		fmt.Println(currEndeDatum)
		zLog.Debug().Msgf("Enddatum %v", currEndeDatum)

		tmpDiffTime := currEndeDatum.Sub(currEinkaufsDatum).Hours() / 24

		AllCertificateData[currColumn][Gesamtlaufzeit] = fmt.Sprintf("%.0f", tmpDiffTime)

		zLog.Debug().Msgf("Laufzeit in Tagen: %v", AllCertificateData[currColumn][Gesamtlaufzeit])

	}

}
