// BonusCertOverview keeps an eye on your Bonus Certificates
package main

import (
	"encoding/json"
	"fmt"
	"io/ioutil"
	"log"
	"os"
	"strings"
	"time"

	"github.com/rs/zerolog"
	"github.com/xuri/excelize/v2"
)

type ConfigValues struct {
	Language string
	DataFile string
}

func main() {

	// Define global variables
	// AllCertificateData := make(map[string]map[string]string)
	zerolog.TimeFieldFormat = time.RFC3339
	zerolog.SetGlobalLevel(zerolog.InfoLevel)
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
	rows, err := dataFile.GetRows("Laufende Aktionen")
	if err != nil {
		zLog.Fatal().Err(err).Msg("")
	}

	// Get the number of rows
	numberOfRows := len(rows)
	fmt.Printf("Anzahl Rows: %v", numberOfRows)
	capOfRows := len(rows)
	fmt.Printf("cap Rows: %v", capOfRows)

	// Get the number of columns
	numberOfColumns := len(rows[0])
	fmt.Printf("Anzahl Columns: %v", numberOfColumns)

	// Iterate over all rows
	for _, row := range rows {
		for _, colCell := range row {
			fmt.Print(colCell, "\t")
		}
		fmt.Println()
	}

	/*

		csvReader := csv.NewReader(dataFile)
		for {

			rec, err := csvReader.Read()
			if err == io.EOF {
				break
			}
			if err != nil {
				log.Fatal(err)
			}
			// do something with read line
			fmt.Printf("%+v\n", rec)

			laenge := len(rec)
			fmt.Printf("Länge %+v\n", laenge)

			zwergName := rec[0]

			AllCertificateData[zwergName] = make(map[string]string)
			AllCertificateData[zwergName]["Einkaufswert"] = "133.20"

			fmt.Printf("EinkaufsWert von Allianz: %s", AllCertificateData[zwergName]["Einkaufswert"])

		}

	*/

}

/*
Bezeichnung
ISIN

Einkaufswert
Anzahl
Spesen
Einkaufswert gesamt
Einkaufsdatum
Ende Datum
Gesamtlaufzeit in Tagen ab Kauf
Restlaufzeit in Tagen
CAP
ISIN Basiswert
Barriere

Aktuelles Datum
Aktueller Wert
Aktueller Wert Basiswert
Aktueller Ertrag in %
Aktueller Ertrag
Erwarteter Ertrag in %
Erwarteter Ertrag
Gesamtertrag in % pro Jahr
Restlaufzeit in Tagen
Ertrag in % bis Restlaufzeit
Ertrag in % pro Jahr bis Restlaufzeit
*/
