package main

import (
	"fmt"
	"github.com/xuri/excelize/v2"
	"strconv"
)

func main() {
	f, err := excelize.OpenFile("Ввод1.xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}
	defer func() {
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()
	// Получить значение из ячейки по заданному имени и оси листа
	cell, err := f.GetCellValue("Sheet1", "A1")
	if err != nil {
		fmt.Println(err)
		return
	}
	fmt.Println(cell)

	// Получить все строки в Sheet1
	rows, err := f.GetRows("Sheet1")
	if err != nil {
		fmt.Println(err)
		return
	}
	//Заполняем шапку
	f.SetCellValue("Sheet2", "A1", "№ п/п")
	f.MergeCell("Sheet2", "A1", "A2")
	f.SetCellValue("Sheet2", "B1", "Система")
	f.MergeCell("Sheet2", "B1", "B2")
	f.SetCellValue("Sheet2", "C1", "Оборудование")
	f.MergeCell("Sheet2", "C1", "C2")
	f.SetCellValue("Sheet2", "D1", "Описание сигнала/ алгоритма")
	f.MergeCell("Sheet2", "D1", "D2")
	f.SetCellValue("Sheet2", "E1", "Тип сигнала")
	f.MergeCell("Sheet2", "E1", "H1")
	f.SetCellValue("Sheet2", "E2", "AI")
	f.SetCellValue("Sheet2", "F2", "DI")
	f.SetCellValue("Sheet2", "G2", "DO")
	f.SetCellValue("Sheet2", "H2", "AI")
	currentRow := 3
	currentRowStr := ""
	//addressSheet := ""

	for _, row := range rows {

		//for _, colCell := range row {
		//for _, colCell := range row {
		//Если ячейка содержит ЩД-
		//if strings.Contains(colCell, "ЩД") {
		//	addressSheet = "A" + strconv.Itoa(currentRow)
		//	f.SetCellValue("Sheet2", addressSheet, row[i])
		//	continue
		//}

		if row[0] != "ЩД" {
			currentRowStr = strconv.Itoa(currentRow + 1)
			switch len(row) {
			case 1:
			case 2:
				fmt.Println("case 2")
				f.SetCellValue("Sheet2", "B"+currentRowStr, row[1])

			case 3:
				fmt.Println("case 3")
				f.SetCellValue("Sheet2", "B"+currentRowStr, row[1]+" "+row[2])
			}
		} else {
			currentRowStr = strconv.Itoa(currentRow)
			f.SetCellValue("Sheet2", "A"+currentRowStr, row[1])
			f.MergeCell("Sheet2", "A"+currentRowStr, "H"+currentRowStr)
		}

		switch row[0] {
		case "ЩД":
			currentRow++
			currentRowStr = strconv.Itoa(currentRow)
			f.SetCellValue("Sheet2", "B"+currentRowStr, row[1])
			f.SetCellValue("Sheet2", "C"+currentRowStr, "Вводной рубильник QS")
			f.SetCellValue("Sheet2", "D"+currentRowStr, "Вкл/выкл")
			f.SetCellValue("Sheet2", "F"+currentRowStr, "1")
			currentRow++
			currentRowStr = strconv.Itoa(currentRow)
			f.SetCellValue("Sheet2", "C"+currentRowStr, "Реле контроля напряжения")
			f.SetCellValue("Sheet2", "D"+currentRowStr, "Норма/Авария")
			f.SetCellValue("Sheet2", "F"+currentRowStr, "1")
			currentRow++
			currentRowStr = strconv.Itoa(currentRow)
			f.SetCellValue("Sheet2", "C"+currentRowStr, "Групповые автоматические выключатели QF")
			f.SetCellValue("Sheet2", "D"+currentRowStr, "Норма/Авария")
			f.SetCellValue("Sheet2", "F"+currentRowStr, "1")
		case "Узел ввода ГВС":
			//Заполняем таблицу по типу 1
			currentRow++
			currentRowStr = strconv.Itoa(currentRow)
			f.SetCellValue("Sheet2", "C"+currentRowStr, "Датчик протечки LS1")
			f.SetCellValue("Sheet2", "D"+currentRowStr, "Норма/Авария")
			f.SetCellValue("Sheet2", "E"+currentRowStr, "1")
			currentRow++
			currentRowStr = strconv.Itoa(currentRow)
			f.SetCellValue("Sheet2", "C"+currentRowStr, "Температура ГВС")
			f.SetCellValue("Sheet2", "D"+currentRowStr, "Значение температуры")
			f.SetCellValue("Sheet2", "F"+currentRowStr, "1")
			currentRow++
			currentRowStr = strconv.Itoa(currentRow)
			f.SetCellValue("Sheet2", "C"+currentRowStr, "Давление в системе ГВС HW-PE1")
			f.SetCellValue("Sheet2", "D"+currentRowStr, "Значение давления")
			f.SetCellValue("Sheet2", "F"+currentRowStr, "1")
			currentRow++
			currentRowStr = strconv.Itoa(currentRow)
			f.SetCellValue("Sheet2", "C"+currentRowStr, "Давление в системе ХВС CW-PE1")
			f.SetCellValue("Sheet2", "D"+currentRowStr, "Значение давления")
			f.SetCellValue("Sheet2", "F"+currentRowStr, "1")
		case "Коллектор отопления":
			//Заполняем таблицу по типу 2
			currentRow++
			currentRowStr = strconv.Itoa(currentRow)
			f.SetCellValue("Sheet2", "C"+currentRowStr, "Протечка")
			f.SetCellValue("Sheet2", "D"+currentRowStr, "Норма/Авария")
			f.SetCellValue("Sheet2", "E"+currentRowStr, "1")
			currentRow++
			currentRowStr = strconv.Itoa(currentRow)
			f.SetCellValue("Sheet2", "C"+currentRowStr, "Давление холодной воды")
			f.SetCellValue("Sheet2", "D"+currentRowStr, "Значение давления")
			f.SetCellValue("Sheet2", "F"+currentRowStr, "1")
			currentRow++
			currentRowStr = strconv.Itoa(currentRow)
			f.SetCellValue("Sheet2", "C"+currentRowStr, "Давление горячей воды")
			f.SetCellValue("Sheet2", "D"+currentRowStr, "Значение давления")
			f.SetCellValue("Sheet2", "F"+currentRowStr, "1")
			currentRow++
			currentRowStr = strconv.Itoa(currentRow)
			f.SetCellValue("Sheet2", "C"+currentRowStr, "Температура ХВС")
			f.SetCellValue("Sheet2", "D"+currentRowStr, "Значение температуры")
			f.SetCellValue("Sheet2", "F"+currentRowStr, "1")
			currentRow++
			currentRowStr = strconv.Itoa(currentRow)
			f.SetCellValue("Sheet2", "C"+currentRowStr, "Температура ГВС")
			f.SetCellValue("Sheet2", "D"+currentRowStr, "Значение температуры")
			f.SetCellValue("Sheet2", "F"+currentRowStr, "1")
			currentRow++
			currentRowStr = strconv.Itoa(currentRow)
			f.SetCellValue("Sheet2", "C"+currentRowStr, "Управление отсечным клапаном ХВС")
			f.SetCellValue("Sheet2", "D"+currentRowStr, "Регулирование")
			f.SetCellValue("Sheet2", "H"+currentRowStr, "1")
			currentRow++
			currentRowStr = strconv.Itoa(currentRow)
			f.SetCellValue("Sheet2", "D"+currentRowStr, "Измерение")
			f.SetCellValue("Sheet2", "F"+currentRowStr, "1")
			currentRow++
			currentRowStr = strconv.Itoa(currentRow)
			f.SetCellValue("Sheet2", "C"+currentRowStr, "Управление отсечным клапаном ГВС")
			f.MergeCell("Sheet2", "C"+currentRowStr, "C"+strconv.Itoa(currentRow+1))
			f.SetCellValue("Sheet2", "D"+currentRowStr, "Регулирование")
			f.SetCellValue("Sheet2", "H"+currentRowStr, "1")
			currentRow++
			currentRowStr = strconv.Itoa(currentRow)
			f.SetCellValue("Sheet2", "D"+currentRowStr, "Измерение")
			f.SetCellValue("Sheet2", "F"+currentRowStr, "1")
		case "Щит КН":
			//Заполняем таблицу по типу 1
			currentRow++
			currentRowStr = strconv.Itoa(currentRow)
			f.SetCellValue("Sheet2", "C"+currentRowStr, "Вводной рубильник QS")
			f.SetCellValue("Sheet2", "D"+currentRowStr, "Вкл/выкл")
			f.SetCellValue("Sheet2", "F"+currentRowStr, "1")
			currentRow++
			currentRowStr = strconv.Itoa(currentRow)
			f.SetCellValue("Sheet2", "C"+currentRowStr, "Реле контроля фаз KV")
			f.SetCellValue("Sheet2", "D"+currentRowStr, "Вкл/выкл")
			f.SetCellValue("Sheet2", "F"+currentRowStr, "1")
			currentRow++
			currentRowStr = strconv.Itoa(currentRow)
			f.SetCellValue("Sheet2", "C"+currentRowStr, "Групповой автоматический выключатель QF")
			f.MergeCell("Sheet2", "C"+currentRowStr, "C"+strconv.Itoa(currentRow+1))
			f.SetCellValue("Sheet2", "D"+currentRowStr, "Вкл/выкл")
			f.SetCellValue("Sheet2", "F"+currentRowStr, "1")
			currentRow++
			currentRowStr = strconv.Itoa(currentRow)
			f.SetCellValue("Sheet2", "D"+currentRowStr, "Норма/Авария")
			f.SetCellValue("Sheet2", "F"+currentRowStr, "1")
		case "ГВС кофе":
			//Заполняем таблицу по типу 1
			currentRow++
			currentRowStr = strconv.Itoa(currentRow)
			f.SetCellValue("Sheet2", "C"+currentRowStr, "Управление отсечным клапаном ХВС")
			f.MergeCell("Sheet2", "C"+currentRowStr, "C"+strconv.Itoa(currentRow+1))
			f.SetCellValue("Sheet2", "D"+currentRowStr, "Регулирование")
			f.SetCellValue("Sheet2", "H"+currentRowStr, "1")
			currentRow++
			currentRowStr = strconv.Itoa(currentRow)
			f.SetCellValue("Sheet2", "D"+currentRowStr, "Измерение")
			f.SetCellValue("Sheet2", "E"+currentRowStr, "1")
			currentRow++
			currentRowStr = strconv.Itoa(currentRow)
			f.SetCellValue("Sheet2", "C"+currentRowStr, "Управление отсечным клапаном ГВС")
			f.MergeCell("Sheet2", "C"+currentRowStr, "C"+strconv.Itoa(currentRow+1))
			f.SetCellValue("Sheet2", "D"+currentRowStr, "Регулирование")
			f.SetCellValue("Sheet2", "H"+currentRowStr, "1")
			currentRow++
			currentRowStr = strconv.Itoa(currentRow)
			f.SetCellValue("Sheet2", "D"+currentRowStr, "Измерение")
			f.SetCellValue("Sheet2", "E"+currentRowStr, "1")
			currentRow++
			currentRowStr = strconv.Itoa(currentRow)
			f.SetCellValue("Sheet2", "C"+currentRowStr, "Протечка в мини-кухне")
			f.SetCellValue("Sheet2", "D"+currentRowStr, "Норма/Авария")
			f.SetCellValue("Sheet2", "F"+currentRowStr, "1")
		case "Щит КФ":
			//Заполняем таблицу по типу 1
			currentRow++
			currentRowStr = strconv.Itoa(currentRow)
			f.SetCellValue("Sheet2", "C"+currentRowStr, "Вводной рубильник QS")
			f.SetCellValue("Sheet2", "D"+currentRowStr, "Вкл/выкл")
			f.SetCellValue("Sheet2", "F"+currentRowStr, "1")
			currentRow++
			currentRowStr = strconv.Itoa(currentRow)
			f.SetCellValue("Sheet2", "C"+currentRowStr, "Реле контроля фаз KVS")
			f.SetCellValue("Sheet2", "D"+currentRowStr, "Норма/Авария")
			f.SetCellValue("Sheet2", "F"+currentRowStr, "1")
			currentRow++
			currentRowStr = strconv.Itoa(currentRow)
			f.SetCellValue("Sheet2", "C"+currentRowStr, "Групповой автоматический выключатель QF")
			f.MergeCell("Sheet2", "C"+currentRowStr, "C"+strconv.Itoa(currentRow+1))
			f.SetCellValue("Sheet2", "D"+currentRowStr, "Вкл/выкл")
			f.SetCellValue("Sheet2", "F"+currentRowStr, "1")
			currentRow++
			currentRowStr = strconv.Itoa(currentRow)
			f.SetCellValue("Sheet2", "D"+currentRowStr, "Норма/Авария")
			f.SetCellValue("Sheet2", "F"+currentRowStr, "1")
		case "АПГТ":
			//Заполняем таблицу по типу 1
			currentRow++
			currentRowStr = strconv.Itoa(currentRow)
			f.SetCellValue("Sheet2", "C"+currentRowStr, "Давление в баллоне газового пожаротушения мультиплексорной ГТИ")
			f.SetCellValue("Sheet2", "D"+currentRowStr, "Норма/Авария")
			f.SetCellValue("Sheet2", "F"+currentRowStr, "1")
		case "Пом СС":
			//Заполняем таблицу по типу 1
			currentRow++
			currentRowStr = strconv.Itoa(currentRow)
			f.SetCellValue("Sheet2", "C"+currentRowStr, "Датчик протечки LS1")
			f.SetCellValue("Sheet2", "D"+currentRowStr, "Норма/Авария")
			f.SetCellValue("Sheet2", "F"+currentRowStr, "1")
		case "Щит Б":
			//Заполняем таблицу по типу 1
			currentRow++
			currentRowStr = strconv.Itoa(currentRow)
			f.SetCellValue("Sheet2", "C"+currentRowStr, "Рубильник Питание ИБП")
			f.SetCellValue("Sheet2", "D"+currentRowStr, "Вкл/выкл")
			f.SetCellValue("Sheet2", "F"+currentRowStr, "1")
			currentRow++
			currentRowStr = strconv.Itoa(currentRow)
			f.SetCellValue("Sheet2", "C"+currentRowStr, "Байпасный рубильник Позиция I Питание от ИБП")
			f.SetCellValue("Sheet2", "D"+currentRowStr, "Вкл/выкл")
			f.SetCellValue("Sheet2", "F"+currentRowStr, "1")
			currentRow++
			currentRowStr = strconv.Itoa(currentRow)
			f.SetCellValue("Sheet2", "C"+currentRowStr, "Байпасный рубильник Позиция II Питание по байпасу")
			f.SetCellValue("Sheet2", "D"+currentRowStr, "Вкл/выкл")
			f.SetCellValue("Sheet2", "F"+currentRowStr, "1")
		case "ИБП":
			//Заполняем таблицу по типу 1
			currentRow++
			currentRowStr = strconv.Itoa(currentRow)
			f.SetCellValue("Sheet2", "C"+currentRowStr, "ИБП в норме")
			f.SetCellValue("Sheet2", "D"+currentRowStr, "Норма/Авария")
			f.SetCellValue("Sheet2", "F"+currentRowStr, "1")
			currentRow++
			currentRowStr = strconv.Itoa(currentRow)
			f.SetCellValue("Sheet2", "C"+currentRowStr, "Работа от инвертора")
			f.SetCellValue("Sheet2", "D"+currentRowStr, "Норма/Авария")
			f.SetCellValue("Sheet2", "F"+currentRowStr, "1")
			currentRow++
			currentRowStr = strconv.Itoa(currentRow)
			f.SetCellValue("Sheet2", "C"+currentRowStr, "Низкий заряд батареи")
			f.SetCellValue("Sheet2", "D"+currentRowStr, "Норма/Авария")
			f.SetCellValue("Sheet2", "F"+currentRowStr, "1")
			currentRow++
			currentRowStr = strconv.Itoa(currentRow)
			f.SetCellValue("Sheet2", "C"+currentRowStr, "Общая авария ИБП")
			f.SetCellValue("Sheet2", "D"+currentRowStr, "Норма/Авария")
			f.SetCellValue("Sheet2", "F"+currentRowStr, "1")
		case "Узел ввода УВ":
			//Заполняем таблицу по типу 1
			currentRow++
			currentRowStr = strconv.Itoa(currentRow)
			f.SetCellValue("Sheet2", "C"+currentRowStr, "Температура охлажденной воды")
			f.SetCellValue("Sheet2", "D"+currentRowStr, "Значение температуры")
			f.SetCellValue("Sheet2", "E"+currentRowStr, "1")
			currentRow++
			currentRowStr = strconv.Itoa(currentRow)
			f.SetCellValue("Sheet2", "C"+currentRowStr, "Температура нагретой воды")
			f.SetCellValue("Sheet2", "D"+currentRowStr, "Значение температуры")
			f.SetCellValue("Sheet2", "E"+currentRowStr, "1")
			currentRow++
			currentRowStr = strconv.Itoa(currentRow)
			f.SetCellValue("Sheet2", "C"+currentRowStr, "Давление охлажденной воды")
			f.SetCellValue("Sheet2", "D"+currentRowStr, "Значение давления")
			f.SetCellValue("Sheet2", "E"+currentRowStr, "1")
			currentRow++
			currentRowStr = strconv.Itoa(currentRow)
			f.SetCellValue("Sheet2", "C"+currentRowStr, "Давление нагретой воды")
			f.SetCellValue("Sheet2", "D"+currentRowStr, "Значение давления")
			f.SetCellValue("Sheet2", "E"+currentRowStr, "1")
		case "Система миникухни":
			//Заполняем таблицу по типу 1
			currentRow++
			currentRowStr = strconv.Itoa(currentRow)
			f.SetCellValue("Sheet2", "C"+currentRowStr, "Счетчик воды 1")
			f.SetCellValue("Sheet2", "D"+currentRowStr, "Значение л/ч")
			f.SetCellValue("Sheet2", "E"+currentRowStr, "1")
			currentRow++
			currentRowStr = strconv.Itoa(currentRow)
			f.SetCellValue("Sheet2", "C"+currentRowStr, "Счетчик воды 2")
			f.SetCellValue("Sheet2", "D"+currentRowStr, "Значение л/ч")
			f.SetCellValue("Sheet2", "E"+currentRowStr, "1")
			currentRow++
			currentRowStr = strconv.Itoa(currentRow)
			f.SetCellValue("Sheet2", "C"+currentRowStr, "Давление датчика давления 1")
			f.SetCellValue("Sheet2", "D"+currentRowStr, "Значение давления")
			f.SetCellValue("Sheet2", "E"+currentRowStr, "1")
			currentRow++
			currentRowStr = strconv.Itoa(currentRow)
			f.SetCellValue("Sheet2", "C"+currentRowStr, "Давление датчика давления 2")
			f.SetCellValue("Sheet2", "D"+currentRowStr, "Значение давления")
			f.SetCellValue("Sheet2", "E"+currentRowStr, "1")
			currentRow++
			currentRowStr = strconv.Itoa(currentRow)
			f.SetCellValue("Sheet2", "C"+currentRowStr, "Датчик протечки 1")
			f.SetCellValue("Sheet2", "D"+currentRowStr, "Норма/Авария")
			f.SetCellValue("Sheet2", "F"+currentRowStr, "1")
			currentRow++
			currentRowStr = strconv.Itoa(currentRow)
			f.SetCellValue("Sheet2", "C"+currentRowStr, "Датчик протечки 2")
			f.SetCellValue("Sheet2", "D"+currentRowStr, "Норма/Авария")
			f.SetCellValue("Sheet2", "F"+currentRowStr, "1")
			currentRow++
			currentRowStr = strconv.Itoa(currentRow)
			f.SetCellValue("Sheet2", "C"+currentRowStr, "Статус кран закрыт")
			f.SetCellValue("Sheet2", "D"+currentRowStr, "Норма/Авария")
			f.SetCellValue("Sheet2", "F"+currentRowStr, "1")
		case "Listcontroller":
			//Заполняем таблицу по типу 1
			currentRow++
			currentRowStr = strconv.Itoa(currentRow)
			f.SetCellValue("Sheet2", "C"+currentRowStr, "Температура в кабельном лотке")
			f.SetCellValue("Sheet2", "D"+currentRowStr, "Значение температуры")
			f.SetCellValue("Sheet2", "F"+currentRowStr, "1")
			currentRow++
			currentRowStr = strconv.Itoa(currentRow)
			f.SetCellValue("Sheet2", "C"+currentRowStr, "Обнаружение очага возгорания кабельных трасс")
			f.SetCellValue("Sheet2", "D"+currentRowStr, "Норма/Авария")
			f.SetCellValue("Sheet2", "F"+currentRowStr, "1")
			//через дефер добавить помещения, добавить всем левый столбик (B)
		}
		//fmt.Println(row[c])

		//if i == 2 {
		//	cell, err := f.GetCellValue("Sheet1", "A1")
		//	if err != nil {
		//		fmt.Println(err)
		//		return
		//	}
		//}

		//fmt.Print(row[i], "\t")

		// Сохранить файл xlsx по данному пути
		if err := f.SaveAs("Вывод.xlsx"); err != nil {
			fmt.Println(err)
		}
	}
	fmt.Println()
}
