package main

import (
	"errors"
	"flag"
	"fmt"
	ical "github.com/arran4/golang-ical"
	"github.com/segmentio/ksuid"
	"github.com/tealeg/xlsx"
	"log"
	"os"
	"strconv"
	"strings"
	"time"
)

func main() {
	var employeeName, filePath string

	flag.StringVar(&employeeName, "name", "Jane Doe", "name of the employee")
	flag.StringVar(&filePath, "file-path", ".", "path to the excel sheet")
	flag.Parse()

	fmt.Println(employeeName)
	fmt.Println(filePath)

	xlFile, err := xlsx.OpenFile(filePath)
	if err != nil {
		log.Fatalf("error openeing the excel sheet %s: %s", filePath, err)
	}

	// Iterate over each sheet in the Excel file
	var dateRow, employeeRow *xlsx.Row
	for _, sheet := range xlFile.Sheets {
		// Iterate over each row in the sheet
		for _, row := range sheet.Rows {
			currentRow := row
			val, err := currentRow.Cells[0].FormattedValue()
			if err != nil {
				log.Fatalf("unsupported format of excel sheet: %s", err)
			}

			if strings.EqualFold(val, "date:") {
				dateRow = currentRow
			} else if strings.EqualFold(employeeName, val) {
				employeeRow = currentRow
			}
		}
	}

	if dateRow == nil {
		errorMsg := "could not find a row with the dates"
		log.Fatalf("unsupported format of excel sheet: %s", errorMsg)
	}

	if employeeRow == nil {
		errorMsg := "could not find a row with employee's name"
		log.Fatalf("unsupported format of excel sheet: %s", errorMsg)
	}

	if len(dateRow.Cells) != len(employeeRow.Cells) {
		errorMsg := "didn't find the same number of columns for date and employee name rows"
		log.Fatalf("unsupported format of excel sheet: %s", errorMsg)
	}

	// Start from 3rd column.
	cal := ical.NewCalendarFor("pharm-schedule")
	cal.SetMethod(ical.MethodPublish)

	for i := 2; i < len(dateRow.Cells); i++ {
		dateVal, err := dateRow.Cells[i].FormattedValue()
		if err != nil {
			errorMsg := "unable to parse date"
			log.Fatalf("unsupported format of excel sheet: %s", errorMsg)
		}
		scheduleVal, err := employeeRow.Cells[i].FormattedValue()
		if err != nil {
			errorMsg := "unable to parse schedule"
			log.Fatalf("unsupported format of excel sheet: %s", errorMsg)
		}

		// If the employee is scheduled that day ...
		if scheduleVal != "" && dateVal != "" {
			startTime, err := parseDate(dateVal)
			if err != nil {
				log.Fatalf("failed to parse date: %s", err)
			}
			endTime := startTime.Add(510 * time.Minute) // 8.5 hours

			event := cal.AddEvent(ksuid.New().String())
			event.SetCreatedTime(time.Now())
			event.SetDtStampTime(time.Now())
			event.SetStartAt(startTime)
			event.SetEndAt(endTime)
			event.SetSummary(scheduleVal)
			event.SetDescription(scheduleVal)
		}
	}

	// Generate the .ics file
	outputFileName := fmt.Sprintf("%s schedule.ics", employeeName)
	file, err := os.Create(outputFileName)
	if err != nil {
		log.Fatal(err)
	}
	defer file.Close()

	// Write the calendar to the .ics file
	err = cal.SerializeTo(file)
	if err != nil {
		log.Fatal(err)
	}

	log.Println("Employee schedule converted to .ics format successfully!")
}

func parseDate(dateVal string) (time.Time, error) {
	d := strings.Split(dateVal, "-")
	if len(d) != 2 {
		return time.Time{}, errors.New("could not understand the date")
	}

	// TODO: make day parsing more robust
	day, _ := strconv.Atoi(d[0])
	var month time.Month
	// TODO: make month parsing logic more robust
	switch d[1] {
	case "Jan":
		month = time.January
	case "Feb":
		month = time.February
	case "Mar":
		month = time.March
	case "Apr":
		month = time.April
	case "May":
		month = time.May
	case "Jun":
		month = time.June
	case "Jul":
		month = time.July
	case "Aug":
		month = time.August
	case "Sept":
		month = time.September
	case "Oct":
		month = time.October
	case "Nov":
		month = time.November
	case "Dec":
		month = time.December
	default:
		return time.Time{}, errors.New("unknown month in date")
	}

	return time.Date(time.Now().Year(), month, day, 7, 0, 0, 0, time.Local), nil
}
