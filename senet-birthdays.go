package main

import (
	"fmt"
	"strconv"
	"time"
	"io"
	"net/http"
	"os"

	"github.com/360EntSecGroup-Skylar/excelize"
	"github.com/dex35/smsru"
	"gopkg.in/ini.v1"
)

type Client struct {
	Login     string
	FirstName string
	Birthday  time.Time
}

// note: this is off by two days on the real epoch (1/1/1900) because
// - the days are 1 indexed so 1/1/1900 is 1 not 0
// - Excel pretends that Feb 29, 1900 existed even though it did not
// The following function will fail for dates before March 1st 1900
// Before that date the Julian calendar was used so a conversion would be necessary
var excelEpoch = time.Date(1899, time.December, 30, 0, 0, 0, 0, time.UTC)
var smsClient *smsru.SmsClient
const configFilename = "config.ini"
const reportFileName = "report.xlsx"
const senetReportUrl = "/api/v2/stat/reports/?format=xlsx&office_id=1&type=created_users&from_date=2019-08-15T00%3A00%3A00.000Z&to_date=2021-01-01T00%3A00%3A00.999"

var smsruApiToken = ""
var senetDomain = ""
var announceText = ""

func main() {
	fmt.Println("Started")

	initConfig()

	smsClient = smsru.CreateClient(smsruApiToken)
	now := time.Now()

	downloadFile(reportFileName, fmt.Sprintf("https://%s%s", senetDomain, senetReportUrl))
	fmt.Println("Downloaded", reportFileName)

	f, err := excelize.OpenFile(reportFileName)
	if err != nil {
		fmt.Println(err)
		return
	}

	fmt.Println("Parse file...")
	rows := f.GetRows("Sheet1")
	for k, row := range rows {
		if k == 0 {
			continue
		}

		clientData := parseClientXLSData(row)

		// Skip bad birthday
		if clientData.Birthday.Year() < 1900 {
			continue
		}

		if clientData.Birthday.Day() == now.Day() && clientData.Birthday.Month() == now.Month() {
			informClient(clientData)
		}
	}
}

func parseClientXLSData(row []string) Client {
	clientData := Client{}

	for i, colCell := range row {
		if i == 0 {
			clientData.Login = colCell
		}
		if i == 1 {
			clientData.FirstName = colCell
		}
		if i == 3 {
			clientData.Birthday = excelDateToDate(colCell)
		}
	}

	return clientData
}

func initConfig() {
	cfg, err := ini.ShadowLoad(configFilename)
	if err != nil {
		panic("Fail to read config file!")
	}

	smsruApiToken = cfg.Section("").Key("smsruApiToken").Value()
	senetDomain = cfg.Section("").Key("senetDomain").Value()
	announceText = cfg.Section("").Key("announceText").Value()
}

func informClient(clientData Client) {
	fmt.Println("Wohoo! Birthday:", clientData.FirstName, "(", clientData.Birthday, ")", "( Login:", clientData.Login, ")")

	// checkCode := "A1B2C3D4"
	phone := clientData.Login

	// 4PROD 
	if true {
		sms := smsru.CreateSMS(phone, fmt.Sprintf(announceText, clientData.FirstName))
		sendedsms, _ := smsClient.SmsSend(sms)

		fmt.Println(" > Sms sent. Status:", sendedsms)	
	}
}

func excelDateToDate(excelDate string) time.Time {
	var days, _ = strconv.Atoi(excelDate)
	return excelEpoch.Add(time.Second * time.Duration(days*86400))
}

func downloadFile(filepath string, url string) error {

	// Get the data
	resp, err := http.Get(url)
	if err != nil {
		return err
	}
	defer resp.Body.Close()

	// Create the file
	out, err := os.Create(filepath)
	if err != nil {
		return err
	}
	defer out.Close()

	// Write the body to file
	_, err = io.Copy(out, resp.Body)
	return err
}