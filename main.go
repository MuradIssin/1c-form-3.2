package main

import (
	"fmt"
	"log"
	"strings"
	"time"

	"github.com/xuri/excelize/v2"
)

func main() {
	fmt.Println("start 1c - 3.1 версия")
	cur := []string{"KZT", "USD"}

	var periodCol, debitAccountCol, creditAccountCol, debitValueCol, creditValueCol, analitikaDTCol, analitikaKTCol int

	file, err := excelize.OpenFile("1000.xlsx")
	if err != nil {
		fmt.Println(err)
		fmt.Println("Используетй файл 1000.xlsx в том же директории есть запускаете программу")
		timer := time.NewTimer(5 * time.Second)
		<-timer.C
		log.Fatal(err)
	}
	defer file.Close()

	mainListName := file.GetSheetName(0)
	if mainListName == "" {
		log.Panicln("ошибка с исходым файлом Excel")
	}

	nameCompany := ""

	FormRows := []FormRow{}

	rows, err := file.Rows(mainListName)
	if err != nil {
		fmt.Println(err)
		return
	}
	i := 0 //   счетчик строк
	foundHeader := false
outerLoop:
	//цикл строк
	for rows.Next() {
		rowExcel, err := rows.Columns()
		if err != nil {
			fmt.Println(err)
		}

		// счетчик на экран
		if i == i/10*10 { // экранный счетчик
			fmt.Printf("\rотработано записей : %d", i)
		}
		// fmt.Printf("отработано записей : %d\n", i)

		var dateCell, debitAccount, analitikaDT, analitikaDT2, analitikaDT3, analitikaDT4, creditAccount, analitikaKT, analitikaKT2, analitikaKT3, analitikaKT4 string
		var debitValueNum, creditValueNum float64
		// outRow:
		// цикл для стобцов
		for j, cellValue := range rowExcel {
			if i == 0 { // поиск имени
				nameCompany = cellValue
				nameCompany = strings.ReplaceAll(nameCompany, `"`, "")
				fmt.Println(nameCompany)
				i++
				continue
			}
			// поиск заголовков
			if !foundHeader {
				if cellValue == "Период" {
					log.Println(i, j, "Период нашелся")
					periodCol = j
				}
				if cellValue == "Аналитика Дт" {
					log.Println(i, j, "Аналитика Дт нашелся")
					analitikaDTCol = j
				}
				if cellValue == "Аналитика Кт" {
					log.Println(i, j, "Аналитика Кт нашелся")
					analitikaKTCol = j
				}
				if cellValue == "Дебет" {
					log.Println(i, j, "Дебет нашелся")
					debitAccountCol = j
					debitValueCol = j + 1
				}
				if cellValue == "Кредит" {
					log.Println(i, j, "Кредит нашелся")
					foundHeader = true
					creditAccountCol = j
					creditValueCol = j + 1
				}

				continue
			}
			// fmt.Printf("строка %d столбец %d : %s\n", i, j, colCell)
			switch j {
			case periodCol:
				if cellValue == FINALROW {
					fmt.Println("\nобработка листа завершена")
					break outerLoop
				}
				dateCell = cellValue
				dateCell, _ = formatDate(dateCell)
			case debitAccountCol:
				debitAccount = cellValue
			case creditAccountCol:
				creditAccount = cellValue
			case debitValueCol:
				debitValue := cellValue
				if contains(cur, debitValue) {
					// row++
					continue
				}
				debitValueNum = convToFloat(debitValue)
			case creditValueCol:
				creditValue := cellValue
				if contains(cur, creditValue) {
					// row++
					continue
				}
				creditValueNum = convToFloat(creditValue)
			case analitikaDTCol:
				analitikaDT = takeRow(cellValue, 0)
				analitikaDT2 = takeRow(cellValue, 1)
				analitikaDT3 = takeRow(cellValue, 2)
				analitikaDT4 = takeRow(cellValue, 3)
			case analitikaKTCol:
				analitikaKT = takeRow(cellValue, 0)
				analitikaKT2 = takeRow(cellValue, 1)
				analitikaKT3 = takeRow(cellValue, 2)
				analitikaKT4 = takeRow(cellValue, 3)

			}

		}

		// присвоить значение мапе
		if dateCell != "" {
			FormRows = append(FormRows, FormRow{
				Data:        dateCell,
				DebitAcc:    debitAccount,
				Debit:       debitValueNum,
				DebitText:   analitikaDT,
				DebitText2:  analitikaDT2,
				DebitText3:  analitikaDT3,
				DebitText4:  analitikaDT4,
				CreditAcc:   creditAccount,
				Credit:      creditValueNum,
				CreditText:  analitikaKT,
				CreditText2: analitikaKT2,
				CreditText3: analitikaKT3,
				CreditText4: analitikaKT4,
			})
		}
		// fmt.Println(FormRows)
		i++
		// if i > 27 {
		// 	break outerLoop
		// }
	}
	if err = rows.Close(); err != nil {
		fmt.Println(err)
	}

	buildExcel(FormRows, nameCompany)

	fmt.Println("Если есть ошибки прочитайте инструкцию ")
	fmt.Println("или пишите на muradi@freedombank.kz")
	timer2 := time.NewTimer(5 * time.Second)
	// Ожидаем, пока таймер не истечет
	<-timer2.C
	fmt.Println("завершено excel")

}
