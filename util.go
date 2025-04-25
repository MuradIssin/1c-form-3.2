package main

import (
	"fmt"
	"log"
	"strconv"
	"strings"
	"time"

	"github.com/xuri/excelize/v2"
)

const (
	KO       = "Кредитовые обороты"
	CKO      = "Чистые кредитовые обороты"
	POS      = "Поставщики"
	POK      = "Покупатели"
	STARTROW = "Сальдо на начало"
	FINALROW = "Обороты за период и сальдо на конец"
)

type FormRow struct {
	Data        string // дата
	DebitAcc    string
	Debit       float64 //
	DebitText   string
	DebitText2  string
	DebitText3  string
	DebitText4  string
	CreditAcc   string
	Credit      float64
	CreditText  string
	CreditText2 string
	CreditText3 string
	CreditText4 string
}

func formatDate(input string) (string, error) {
	// Парсим строку в объект time
	date, err := time.Parse("02.01.2006", input)
	if err != nil {
		return "", err
	}

	// Форматируем дату в нужный формат "годмесяц"
	return date.Format("2006_01"), nil
}

func contains(arr []string, value string) bool {
	for _, v := range arr {
		if v == value {
			return true
		}
	}
	return false
}

func convToFloat(text string) float64 {
	text = strings.ReplaceAll(text, string(","), "")
	text = strings.ReplaceAll(text, string(" "), "")
	resFloat, err := strconv.ParseFloat(text, 64)
	if err != nil {
		return 0
	}
	return resFloat
}

// func takeFirstRow(s string) string {
// 	newS := strings.Split(s, "\n")
// 	if len(newS) < 2 {
// 		return s
// 	}
// 	return newS[0]
// }

// func takeSecondRow(s string) string {
// 	newS := strings.Split(s, "\n")
// 	if len(newS) < 2 {
// 		return s
// 	}
// 	return newS[1]
// }

func takeRow(s string, n int) string {
	lines := strings.Split(strings.ReplaceAll(s, "\r\n", "\n"), "\n")
	if n < 0 || n >= len(lines) {
		return ""
	}
	return lines[n]
}

func buildExcel(data []FormRow, nameFile string) {
	f := excelize.NewFile()
	defer func() {
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()
	listName := "1"
	_, err := f.NewSheet(listName)
	if err != nil {
		fmt.Println(err)
		return
	}

	err = f.DeleteSheet("sheet1")
	if err != nil {
		log.Println("log_1111")
		fmt.Println(err)
		return
	}

	// Записываем заголовки
	f.SetCellValue(listName, "A1", "год_месяц")
	f.SetCellValue(listName, "B1", "Дебит")
	f.SetCellValue(listName, "C1", "Сумма_Дебит")
	f.SetCellValue(listName, "D1", "Аналитика_Дебит")
	f.SetCellValue(listName, "E1", "Аналитика_Дебит2")
	f.SetCellValue(listName, "F1", "Аналитика_Дебит3")
	f.SetCellValue(listName, "G1", "Аналитика_Дебит4")

	f.SetCellValue(listName, "H1", "Кредит")
	f.SetCellValue(listName, "I1", "Сумма_Кредит")
	f.SetCellValue(listName, "J1", "Аналитика_Кредит")
	f.SetCellValue(listName, "K1", "Аналитика_Кредит2")
	f.SetCellValue(listName, "L1", "Аналитика_Кредит3")
	f.SetCellValue(listName, "M1", "Аналитика_Кредит4")

	err = f.SetColWidth(listName, "D", "G", 50)
	if err != nil {
		log.Println(err)
	}
	err = f.SetColWidth(listName, "J", "M", 50)
	if err != nil {
		log.Println(err)
	}
	err = f.SetColWidth(listName, "C", "C", 25)
	if err != nil {
		log.Println(err)
	}
	err = f.SetColWidth(listName, "I", "I", 25)
	if err != nil {
		log.Println(err)
	}

	styleNum, err := f.NewStyle(&excelize.Style{
		NumFmt: 3,
	})
	if err != nil {
		fmt.Println(err)
	}
	err = f.SetCellStyle(listName, "C2", "C1000", styleNum)
	if err != nil {
		fmt.Println("ошибка со стилями эксель")
	}
	err = f.SetCellStyle(listName, "I2", "I1000", styleNum)
	if err != nil {
		fmt.Println("ошибка со стилями эксель")
	}

	arrDebit := []string{
		"1030",
		"1022",
		"1030"}

	row := 1
	for _, rowForm := range data {
		// Индексы строк в Excel начинаются с 2, т.к. строка 1 - это заголовки

		if contains(arrDebit, rowForm.DebitAcc) {
			row += 1
			f.SetCellValue(listName, fmt.Sprintf("A%d", row), rowForm.Data)
			f.SetCellValue(listName, fmt.Sprintf("B%d", row), rowForm.DebitAcc)
			f.SetCellValue(listName, fmt.Sprintf("C%d", row), rowForm.Debit)
			f.SetCellValue(listName, fmt.Sprintf("D%d", row), rowForm.DebitText)
			f.SetCellValue(listName, fmt.Sprintf("E%d", row), rowForm.DebitText2)
			f.SetCellValue(listName, fmt.Sprintf("F%d", row), rowForm.DebitText3)
			f.SetCellValue(listName, fmt.Sprintf("G%d", row), rowForm.DebitText4)
			f.SetCellValue(listName, fmt.Sprintf("H%d", row), rowForm.CreditAcc)
			f.SetCellValue(listName, fmt.Sprintf("I%d", row), rowForm.Credit)
			f.SetCellValue(listName, fmt.Sprintf("J%d", row), rowForm.CreditText)
			f.SetCellValue(listName, fmt.Sprintf("K%d", row), rowForm.CreditText2)
			f.SetCellValue(listName, fmt.Sprintf("L%d", row), rowForm.CreditText3)
			f.SetCellValue(listName, fmt.Sprintf("M%d", row), rowForm.CreditText4)
		}

	}

	// // Добавляем сводную таблицу - кредитовые обороты
	// // f.NewSheet(KO)
	f.NewSheet("Поступления")

	if err := f.AddPivotTable(&excelize.PivotTableOptions{
		DataRange:       listName + "!A1:M" + strconv.Itoa(row), // исходные данные
		PivotTableRange: "Поступления!A4:M34",                   // куда положить сводную таблицу
		Rows: []excelize.PivotTableField{
			{Data: "Кредит", DefaultSubtotal: true}, {Data: "Аналитика_Кредит2"},
		},
		Data: []excelize.PivotTableField{
			{Data: "Сумма_Дебит", Name: "Сумма_Дебит", Subtotal: "Sum"}},
		Columns: []excelize.PivotTableField{
			{Data: "год_месяц", DefaultSubtotal: true}},

		RowGrandTotals: true,
		ColGrandTotals: true,
		ShowDrill:      true,
		ShowRowHeaders: true,
		ShowColHeaders: true,
		ShowLastColumn: true,
	}); err != nil {
		fmt.Println(err)
		return
	}

	// 2 лист

	listName = "2"
	_, err = f.NewSheet(listName)
	if err != nil {
		fmt.Println(err)
		return
	}

	// Записываем заголовки
	f.SetCellValue(listName, "A1", "год_месяц")
	f.SetCellValue(listName, "B1", "Дебит")
	f.SetCellValue(listName, "C1", "Сумма_Дебит")
	f.SetCellValue(listName, "D1", "Аналитика_Дебит")
	f.SetCellValue(listName, "E1", "Аналитика_Дебит2")
	f.SetCellValue(listName, "F1", "Аналитика_Дебит3")
	f.SetCellValue(listName, "G1", "Аналитика_Дебит4")

	f.SetCellValue(listName, "H1", "Кредит")
	f.SetCellValue(listName, "I1", "Сумма_Кредит")
	f.SetCellValue(listName, "J1", "Аналитика_Кредит")
	f.SetCellValue(listName, "K1", "Аналитика_Кредит2")
	f.SetCellValue(listName, "L1", "Аналитика_Кредит3")
	f.SetCellValue(listName, "M1", "Аналитика_Кредит4")

	err = f.SetColWidth(listName, "D", "G", 50)
	if err != nil {
		log.Println(err)
	}
	err = f.SetColWidth(listName, "J", "M", 50)
	if err != nil {
		log.Println(err)
	}
	err = f.SetColWidth(listName, "C", "C", 25)
	if err != nil {
		log.Println(err)
	}
	err = f.SetColWidth(listName, "I", "I", 25)
	if err != nil {
		log.Println(err)
	}

	err = f.SetCellStyle(listName, "C2", "C1000", styleNum)
	if err != nil {
		fmt.Println("ошибка со стилями эксель")
	}
	err = f.SetCellStyle(listName, "I2", "I1000", styleNum)
	if err != nil {
		fmt.Println("ошибка со стилями эксель")
	}

	arrCredit := []string{
		"1010",
		"1022",
		"1030"}

	row = 1
	for _, rowForm := range data {
		// Индексы строк в Excel начинаются с 2, т.к. строка 1 - это заголовки

		if contains(arrCredit, rowForm.CreditAcc) {
			row += 1
			f.SetCellValue(listName, fmt.Sprintf("A%d", row), rowForm.Data)
			f.SetCellValue(listName, fmt.Sprintf("B%d", row), rowForm.DebitAcc)
			f.SetCellValue(listName, fmt.Sprintf("C%d", row), rowForm.Debit)
			f.SetCellValue(listName, fmt.Sprintf("D%d", row), rowForm.DebitText)
			f.SetCellValue(listName, fmt.Sprintf("E%d", row), rowForm.DebitText2)
			f.SetCellValue(listName, fmt.Sprintf("F%d", row), rowForm.DebitText3)
			f.SetCellValue(listName, fmt.Sprintf("G%d", row), rowForm.DebitText4)
			f.SetCellValue(listName, fmt.Sprintf("H%d", row), rowForm.CreditAcc)
			f.SetCellValue(listName, fmt.Sprintf("I%d", row), rowForm.Credit)
			f.SetCellValue(listName, fmt.Sprintf("J%d", row), rowForm.CreditText)
			f.SetCellValue(listName, fmt.Sprintf("K%d", row), rowForm.CreditText2)
			f.SetCellValue(listName, fmt.Sprintf("L%d", row), rowForm.CreditText3)
			f.SetCellValue(listName, fmt.Sprintf("M%d", row), rowForm.CreditText4)

		}
	}
	f.NewSheet("Выбытия")

	if err := f.AddPivotTable(&excelize.PivotTableOptions{
		DataRange:       listName + "!A1:M" + strconv.Itoa(row), // исходные данные
		PivotTableRange: "Выбытия!A4:M34",                       // куда положить сводную таблицу
		Rows: []excelize.PivotTableField{
			{Data: "Дебит", DefaultSubtotal: true}, {Data: "Аналитика_Дебит2"},
		},
		Data: []excelize.PivotTableField{
			{Data: "Сумма_Кредит", Name: "Сумма_Кредит", Subtotal: "Sum"}},
		Columns: []excelize.PivotTableField{
			{Data: "год_месяц", DefaultSubtotal: true}},

		RowGrandTotals: true,
		ColGrandTotals: true,
		ShowDrill:      true,
		ShowRowHeaders: true,
		ShowColHeaders: true,
		ShowLastColumn: true,
	}); err != nil {
		fmt.Println(err)
		return
	}

	// 3 лист ++++++++++++++++++++++++++++++++++++++++++++++++++++

	listName = "3"
	_, err = f.NewSheet(listName)
	if err != nil {
		fmt.Println(err)
		return
	}

	// Записываем заголовки
	f.SetCellValue(listName, "A1", "год_месяц")
	f.SetCellValue(listName, "B1", "Дебит")
	f.SetCellValue(listName, "C1", "Сумма_Дебит")
	f.SetCellValue(listName, "D1", "Аналитика_Дебит")
	f.SetCellValue(listName, "E1", "Аналитика_Дебит2")
	f.SetCellValue(listName, "F1", "Аналитика_Дебит3")
	f.SetCellValue(listName, "G1", "Аналитика_Дебит4")

	f.SetCellValue(listName, "H1", "Кредит")
	f.SetCellValue(listName, "I1", "Сумма_Кредит")
	f.SetCellValue(listName, "J1", "Аналитика_Кредит")
	f.SetCellValue(listName, "K1", "Аналитика_Кредит2")
	f.SetCellValue(listName, "L1", "Аналитика_Кредит3")
	f.SetCellValue(listName, "M1", "Аналитика_Кредит4")

	err = f.SetColWidth(listName, "D", "G", 50)
	if err != nil {
		log.Println(err)
	}
	err = f.SetColWidth(listName, "J", "M", 50)
	if err != nil {
		log.Println(err)
	}
	err = f.SetColWidth(listName, "C", "C", 25)
	if err != nil {
		log.Println(err)
	}
	err = f.SetColWidth(listName, "I", "I", 25)
	if err != nil {
		log.Println(err)
	}

	err = f.SetCellStyle(listName, "C2", "C1000", styleNum)
	if err != nil {
		fmt.Println("ошибка со стилями эксель")
	}
	err = f.SetCellStyle(listName, "I2", "I1000", styleNum)
	if err != nil {
		fmt.Println("ошибка со стилями эксель")
	}

	arrDebit3 := []string{
		"1022",
		"1010",
		"1030"}

	arrCredit3 := []string{
		"1210",

		"1211", // MURAD
		"1212", // MURAD
		"1213", // MURAD
		"1213", // MURAD
		"1214", // MURAD
		"1215", // MURAD
		"1216", // MURAD
		"1217", // MURAD
		"1218", // MURAD
		"1219", // MURAD

		"1251",
		"1254",
		"3510",
	}

	row = 1
	for _, rowForm := range data {
		// Индексы строк в Excel начинаются с 2, т.к. строка 1 - это заголовки

		if contains(arrCredit3, rowForm.CreditAcc) && contains(arrDebit3, rowForm.DebitAcc) {
			row += 1
			f.SetCellValue(listName, fmt.Sprintf("A%d", row), rowForm.Data)
			f.SetCellValue(listName, fmt.Sprintf("B%d", row), rowForm.DebitAcc)
			f.SetCellValue(listName, fmt.Sprintf("C%d", row), rowForm.Debit)
			f.SetCellValue(listName, fmt.Sprintf("D%d", row), rowForm.DebitText)
			f.SetCellValue(listName, fmt.Sprintf("E%d", row), rowForm.DebitText2)
			f.SetCellValue(listName, fmt.Sprintf("F%d", row), rowForm.DebitText3)
			f.SetCellValue(listName, fmt.Sprintf("G%d", row), rowForm.DebitText4)
			f.SetCellValue(listName, fmt.Sprintf("H%d", row), rowForm.CreditAcc)
			f.SetCellValue(listName, fmt.Sprintf("I%d", row), rowForm.Credit)
			f.SetCellValue(listName, fmt.Sprintf("J%d", row), rowForm.CreditText)
			f.SetCellValue(listName, fmt.Sprintf("K%d", row), rowForm.CreditText2)
			f.SetCellValue(listName, fmt.Sprintf("L%d", row), rowForm.CreditText3)
			f.SetCellValue(listName, fmt.Sprintf("M%d", row), rowForm.CreditText4)
		}

	}

	f.NewSheet("ЧКО")

	if err := f.AddPivotTable(&excelize.PivotTableOptions{
		DataRange:       listName + "!A1:M" + strconv.Itoa(row), // исходные данные
		PivotTableRange: "ЧКО!A4:M34",                           // куда положить сводную таблицу
		Rows: []excelize.PivotTableField{
			{Data: "Аналитика_Кредит2", DefaultSubtotal: true}},
		Data: []excelize.PivotTableField{
			{Data: "Сумма_Дебит", Name: "Сумма_Дебит"}},
		Columns: []excelize.PivotTableField{
			{Data: "год_месяц", DefaultSubtotal: true}},
		RowGrandTotals: true,
		ColGrandTotals: true,
		ShowDrill:      true,
		ShowRowHeaders: true,
		ShowColHeaders: true,
		ShowLastColumn: true,
	}); err != nil {
		fmt.Println(err)
		return
	}

	// 4 лист ++++++++++++++++++++++++++++++++++++++++++++++++++++

	listName = "4"
	_, err = f.NewSheet(listName)
	if err != nil {
		fmt.Println(err)
		return
	}

	// Записываем заголовки
	f.SetCellValue(listName, "A1", "год_месяц")
	f.SetCellValue(listName, "B1", "Дебит")
	f.SetCellValue(listName, "C1", "Сумма_Дебит")
	f.SetCellValue(listName, "D1", "Аналитика_Дебит")
	f.SetCellValue(listName, "E1", "Аналитика_Дебит2")
	f.SetCellValue(listName, "F1", "Аналитика_Дебит3")
	f.SetCellValue(listName, "G1", "Аналитика_Дебит4")

	f.SetCellValue(listName, "H1", "Кредит")
	f.SetCellValue(listName, "I1", "Сумма_Кредит")
	f.SetCellValue(listName, "J1", "Аналитика_Кредит")
	f.SetCellValue(listName, "K1", "Аналитика_Кредит2")
	f.SetCellValue(listName, "L1", "Аналитика_Кредит3")
	f.SetCellValue(listName, "M1", "Аналитика_Кредит4")

	err = f.SetColWidth(listName, "D", "G", 50)
	if err != nil {
		log.Println(err)
	}
	err = f.SetColWidth(listName, "J", "M", 50)
	if err != nil {
		log.Println(err)
	}
	err = f.SetColWidth(listName, "C", "C", 25)
	if err != nil {
		log.Println(err)
	}
	err = f.SetColWidth(listName, "I", "I", 25)
	if err != nil {
		log.Println(err)
	}

	err = f.SetCellStyle(listName, "C2", "C1000", styleNum)
	if err != nil {
		fmt.Println("ошибка со стилями эксель")
	}
	err = f.SetCellStyle(listName, "I2", "I1000", styleNum)
	if err != nil {
		fmt.Println("ошибка со стилями эксель")
	}

	arrCredit4 := []string{
		"1010",
		"1022",
		"1030",
	}
	arrDebit4 := []string{
		"1710",
		"3233",
		"3310",

		"3311", // MURAD
		"3312", // MURAD
		"3313", // MURAD
		"3314", // MURAD
		"3315", // MURAD
		"3316", // MURAD
		"3317", // MURAD
		"3318", // MURAD
		"3319", // MURAD
	}

	row = 1
	for _, rowForm := range data {
		// Индексы строк в Excel начинаются с 2, т.к. строка 1 - это заголовки

		if contains(arrCredit4, rowForm.CreditAcc) && contains(arrDebit4, rowForm.DebitAcc) {
			row += 1
			f.SetCellValue(listName, fmt.Sprintf("A%d", row), rowForm.Data)
			f.SetCellValue(listName, fmt.Sprintf("B%d", row), rowForm.DebitAcc)
			f.SetCellValue(listName, fmt.Sprintf("C%d", row), rowForm.Debit)
			f.SetCellValue(listName, fmt.Sprintf("D%d", row), rowForm.DebitText)
			f.SetCellValue(listName, fmt.Sprintf("E%d", row), rowForm.DebitText2)
			f.SetCellValue(listName, fmt.Sprintf("F%d", row), rowForm.DebitText3)
			f.SetCellValue(listName, fmt.Sprintf("G%d", row), rowForm.DebitText4)
			f.SetCellValue(listName, fmt.Sprintf("H%d", row), rowForm.CreditAcc)
			f.SetCellValue(listName, fmt.Sprintf("I%d", row), rowForm.Credit)
			f.SetCellValue(listName, fmt.Sprintf("J%d", row), rowForm.CreditText)
			f.SetCellValue(listName, fmt.Sprintf("K%d", row), rowForm.CreditText2)
			f.SetCellValue(listName, fmt.Sprintf("L%d", row), rowForm.CreditText3)
			f.SetCellValue(listName, fmt.Sprintf("M%d", row), rowForm.CreditText4)

		}
	}

	f.NewSheet("ЧДО")

	if err := f.AddPivotTable(&excelize.PivotTableOptions{
		DataRange:       listName + "!A1:M" + strconv.Itoa(row), // исходные данные
		PivotTableRange: "ЧДО!A4:M34",                           // куда положить сводную таблицу
		Rows: []excelize.PivotTableField{
			{Data: "Аналитика_Дебит2", DefaultSubtotal: true}},
		Data: []excelize.PivotTableField{
			{Data: "Сумма_Кредит", Name: "Сумма_Кредит"}},
		Columns: []excelize.PivotTableField{
			{Data: "год_месяц", DefaultSubtotal: true}},
		RowGrandTotals: true,
		ColGrandTotals: true,
		ShowDrill:      true,
		ShowRowHeaders: true,
		ShowColHeaders: true,
		ShowLastColumn: true,
	}); err != nil {
		fmt.Println(err)
		return
	}

	/// все данные

	listName = "all"
	_, err = f.NewSheet(listName)
	if err != nil {
		fmt.Println(err)
		return
	}

	f.SetCellValue(listName, "A1", "год_месяц")
	f.SetCellValue(listName, "B1", "Дебит")
	f.SetCellValue(listName, "C1", "Сумма_Дебит")
	f.SetCellValue(listName, "D1", "Аналитика_Дебит")
	f.SetCellValue(listName, "E1", "Аналитика_Дебит2")
	f.SetCellValue(listName, "F1", "Аналитика_Дебит3")
	f.SetCellValue(listName, "G1", "Аналитика_Дебит4")

	f.SetCellValue(listName, "H1", "Кредит")
	f.SetCellValue(listName, "I1", "Сумма_Кредит")
	f.SetCellValue(listName, "J1", "Аналитика_Кредит")
	f.SetCellValue(listName, "K1", "Аналитика_Кредит2")
	f.SetCellValue(listName, "L1", "Аналитика_Кредит3")
	f.SetCellValue(listName, "M1", "Аналитика_Кредит4")

	err = f.SetColWidth(listName, "D", "G", 50)
	if err != nil {
		log.Println(err)
	}
	err = f.SetColWidth(listName, "J", "M", 50)
	if err != nil {
		log.Println(err)
	}
	err = f.SetColWidth(listName, "C", "C", 25)
	if err != nil {
		log.Println(err)
	}
	err = f.SetColWidth(listName, "I", "I", 25)
	if err != nil {
		log.Println(err)
	}

	err = f.SetCellStyle(listName, "C2", "C1000", styleNum)
	if err != nil {
		fmt.Println("ошибка со стилями эксель")
	}
	err = f.SetCellStyle(listName, "I2", "I1000", styleNum)
	if err != nil {
		fmt.Println("ошибка со стилями эксель")
	}

	row = 1
	for _, rowForm := range data {
		// Индексы строк в Excel начинаются с 2, т.к. строка 1 - это заголовки

		// if contains(arrCredit4, rowForm.CreditAcc) && contains(arrDebit4, rowForm.DebitAcc) {
		row += 1
		f.SetCellValue(listName, fmt.Sprintf("A%d", row), rowForm.Data)
		f.SetCellValue(listName, fmt.Sprintf("B%d", row), rowForm.DebitAcc)
		f.SetCellValue(listName, fmt.Sprintf("C%d", row), rowForm.Debit)
		f.SetCellValue(listName, fmt.Sprintf("D%d", row), rowForm.DebitText)
		f.SetCellValue(listName, fmt.Sprintf("E%d", row), rowForm.DebitText2)
		f.SetCellValue(listName, fmt.Sprintf("F%d", row), rowForm.DebitText3)
		f.SetCellValue(listName, fmt.Sprintf("G%d", row), rowForm.DebitText4)
		f.SetCellValue(listName, fmt.Sprintf("H%d", row), rowForm.CreditAcc)
		f.SetCellValue(listName, fmt.Sprintf("I%d", row), rowForm.Credit)
		f.SetCellValue(listName, fmt.Sprintf("J%d", row), rowForm.CreditText)
		f.SetCellValue(listName, fmt.Sprintf("K%d", row), rowForm.CreditText2)
		f.SetCellValue(listName, fmt.Sprintf("L%d", row), rowForm.CreditText3)
		f.SetCellValue(listName, fmt.Sprintf("M%d", row), rowForm.CreditText4)

	}

	fileName := "1000_v3.1_" + nameFile + ".xlsx"
	if err := f.SaveAs(fileName); err != nil {
		fmt.Println(err)
	}
}
