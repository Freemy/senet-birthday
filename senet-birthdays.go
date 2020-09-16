package main

import (
	"fmt"
	"strconv"
	"time"

	"github.com/360EntSecGroup-Skylar/excelize"
	"github.com/dex35/smsru"
)

type Client struct {
	Login     string
	FirstName string
	Birthday  time.Time
}

const announceText = "%s, CyberLine поздравляет вас с днем рождения и дарит бесплатные 5 часов в подарок: %s"

var smsClient SmsClient

func main() {
	smsClient := smsru.CreateClient("API_KEY")

	now := time.Now()

	f, err := excelize.OpenFile("report.xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}

	rows, err := f.GetRows("Sheet1")
	for k, row := range rows {
		if k == 0 {
			continue
		}

		clientData := parseClientXLSData(row)

		// Кривые ДР пропускаем
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
			colCellFloat, _ := strconv.ParseFloat(colCell, 64)
			clientData.Birthday, _ = excelize.ExcelDateToTime(colCellFloat, false)
		}
	}

	return clientData
}

func informClient(clientData Client) {
	fmt.Println("Wohoo! Birthday:", clientData.FirstName, "(", clientData.Birthday, ")", "( Login:", clientData.Login, ")")

	checkCode := "A1B2C3D4"
	phone := clientData.Login

	sms := smsru.CreateSMS(phone, fmt.Sprintf(announceText, clientData.FirstName, checkCode))
	sendedsms, err := smsClient.SmsSend(sms)
}
