package main

import (
	"fmt"
	"github.com/xuri/excelize/v2"
	"os"
	"strconv"
	"strings"
)

func main() {
	a := 1
	if a == 0 {
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
		f.SetCellValue("Sheet2", "H2", "AO")
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
				//f.MergeCell("Sheet2", "A"+currentRowStr, "H"+currentRowStr)
			}

			switch row[0] {
			case "ЩД":
				//
				currentRow++
				currentRowStr = strconv.Itoa(currentRow)
				f.SetCellValue("Sheet2", "C"+currentRowStr, "Состояние вводного рубильника")
				f.MergeCell("Sheet2", "C"+currentRowStr, "C"+strconv.Itoa(currentRow+1))
				f.SetCellValue("Sheet2", "D"+currentRowStr, "Вкл/выкл")
				f.SetCellValue("Sheet2", "F"+currentRowStr, "1")
				currentRow++
				currentRowStr = strconv.Itoa(currentRow)
				f.SetCellValue("Sheet2", "D"+currentRowStr, "Норма/Авария")
				f.SetCellValue("Sheet2", "F"+currentRowStr, "1")
				//
				currentRow++
				currentRowStr = strconv.Itoa(currentRow)
				f.SetCellValue("Sheet2", "C"+currentRowStr, "Состояние реле контроля напряжения")
				f.MergeCell("Sheet2", "C"+currentRowStr, "C"+strconv.Itoa(currentRow+1))
				f.SetCellValue("Sheet2", "D"+currentRowStr, "Вкл/выкл")
				f.SetCellValue("Sheet2", "F"+currentRowStr, "1")
				currentRow++
				currentRowStr = strconv.Itoa(currentRow)
				f.SetCellValue("Sheet2", "D"+currentRowStr, "Норма/Авария")
				f.SetCellValue("Sheet2", "F"+currentRowStr, "1")

				currentRow++
				currentRowStr = strconv.Itoa(currentRow)
				f.SetCellValue("Sheet2", "C"+currentRowStr, "Общий статус аварий групповых автоматов")
				f.SetCellValue("Sheet2", "D"+currentRowStr, "Норма/Авария")
				f.SetCellValue("Sheet2", "F"+currentRowStr, "1")
			case "гвс":
				//Заполняем таблицу по типу 1
				currentRow++
				currentRowStr = strconv.Itoa(currentRow)
				f.SetCellValue("Sheet2", "C"+currentRowStr, "Протечка")
				f.SetCellValue("Sheet2", "D"+currentRowStr, "Норма/Авария")
				f.SetCellValue("Sheet2", "F"+currentRowStr, "1")

				currentRow++
				currentRowStr = strconv.Itoa(currentRow)
				f.SetCellValue("Sheet2", "C"+currentRowStr, "Температура горячей воды")
				f.MergeCell("Sheet2", "C"+currentRowStr, "C"+strconv.Itoa(currentRow+1))
				f.SetCellValue("Sheet2", "D"+currentRowStr, "Значение температуры")
				f.SetCellValue("Sheet2", "E"+currentRowStr, "1")
				currentRow++
				currentRowStr = strconv.Itoa(currentRow)
				f.SetCellValue("Sheet2", "D"+currentRowStr, "Норма/Авария")
				f.SetCellValue("Sheet2", "F"+currentRowStr, "1")

				currentRow++
				currentRowStr = strconv.Itoa(currentRow)
				f.SetCellValue("Sheet2", "C"+currentRowStr, "Давление гор. воды")
				f.MergeCell("Sheet2", "C"+currentRowStr, "C"+strconv.Itoa(currentRow+1))
				f.SetCellValue("Sheet2", "D"+currentRowStr, "Значение давления")
				f.SetCellValue("Sheet2", "E"+currentRowStr, "1")
				currentRow++
				currentRowStr = strconv.Itoa(currentRow)
				f.SetCellValue("Sheet2", "D"+currentRowStr, "Норма/Авария")
				f.SetCellValue("Sheet2", "F"+currentRowStr, "1")

				currentRow++
				currentRowStr = strconv.Itoa(currentRow)
				f.SetCellValue("Sheet2", "C"+currentRowStr, "Давление холодной воды")
				f.MergeCell("Sheet2", "C"+currentRowStr, "C"+strconv.Itoa(currentRow+1))
				f.SetCellValue("Sheet2", "D"+currentRowStr, "Значение давления")
				f.SetCellValue("Sheet2", "E"+currentRowStr, "1")
				currentRow++
				currentRowStr = strconv.Itoa(currentRow)
				f.SetCellValue("Sheet2", "D"+currentRowStr, "Норма/Авария")
				f.SetCellValue("Sheet2", "F"+currentRowStr, "1")
			case "ко":
				//коллектор отопления
				currentRow++
				currentRowStr = strconv.Itoa(currentRow)
				f.SetCellValue("Sheet2", "C"+currentRowStr, "Протечка")
				f.SetCellValue("Sheet2", "D"+currentRowStr, "Норма/Авария")
				f.SetCellValue("Sheet2", "F"+currentRowStr, "1")

				currentRow++
				currentRowStr = strconv.Itoa(currentRow)
				f.SetCellValue("Sheet2", "C"+currentRowStr, "Давление холодной воды")
				f.MergeCell("Sheet2", "C"+currentRowStr, "C"+strconv.Itoa(currentRow+1))
				f.SetCellValue("Sheet2", "D"+currentRowStr, "Значение давления")
				f.SetCellValue("Sheet2", "E"+currentRowStr, "1")
				currentRow++
				currentRowStr = strconv.Itoa(currentRow)
				f.SetCellValue("Sheet2", "D"+currentRowStr, "Норма/Авария")
				f.SetCellValue("Sheet2", "F"+currentRowStr, "1")

				currentRow++
				currentRowStr = strconv.Itoa(currentRow)
				f.SetCellValue("Sheet2", "C"+currentRowStr, "Давление гор. воды")
				f.MergeCell("Sheet2", "C"+currentRowStr, "C"+strconv.Itoa(currentRow+1))
				f.SetCellValue("Sheet2", "D"+currentRowStr, "Значение давления")
				f.SetCellValue("Sheet2", "E"+currentRowStr, "1")
				currentRow++
				currentRowStr = strconv.Itoa(currentRow)
				f.SetCellValue("Sheet2", "D"+currentRowStr, "Норма/Авария")
				f.SetCellValue("Sheet2", "F"+currentRowStr, "1")
				currentRow++
				currentRowStr = strconv.Itoa(currentRow)
				f.SetCellValue("Sheet2", "C"+currentRowStr, "Температура холодной воды")
				f.MergeCell("Sheet2", "C"+currentRowStr, "C"+strconv.Itoa(currentRow+1))
				f.SetCellValue("Sheet2", "D"+currentRowStr, "Значение температуры")
				f.SetCellValue("Sheet2", "E"+currentRowStr, "1")
				currentRow++
				currentRowStr = strconv.Itoa(currentRow)
				f.SetCellValue("Sheet2", "D"+currentRowStr, "Норма/Авария")
				f.SetCellValue("Sheet2", "F"+currentRowStr, "1")

				currentRow++
				currentRowStr = strconv.Itoa(currentRow)
				f.SetCellValue("Sheet2", "C"+currentRowStr, "Температура гор. воды")
				f.MergeCell("Sheet2", "C"+currentRowStr, "C"+strconv.Itoa(currentRow+1))
				f.SetCellValue("Sheet2", "D"+currentRowStr, "Значение температуры")
				f.SetCellValue("Sheet2", "E"+currentRowStr, "1")
				currentRow++
				currentRowStr = strconv.Itoa(currentRow)
				f.SetCellValue("Sheet2", "D"+currentRowStr, "Норма/Авария")
				f.SetCellValue("Sheet2", "F"+currentRowStr, "1")
				currentRow++
				currentRowStr = strconv.Itoa(currentRow)
				f.SetCellValue("Sheet2", "C"+currentRowStr, "Управление отсечным клапаном холодной воды")
				f.MergeCell("Sheet2", "C"+currentRowStr, "C"+strconv.Itoa(currentRow+1))
				f.SetCellValue("Sheet2", "D"+currentRowStr, "Регулирование")
				f.SetCellValue("Sheet2", "H"+currentRowStr, "1")
				currentRow++
				currentRowStr = strconv.Itoa(currentRow)
				f.SetCellValue("Sheet2", "D"+currentRowStr, "Измерение")
				f.SetCellValue("Sheet2", "F"+currentRowStr, "1")
				currentRow++
				currentRowStr = strconv.Itoa(currentRow)
				f.SetCellValue("Sheet2", "C"+currentRowStr, "Управление отсечным клапаном гор. воды")
				f.MergeCell("Sheet2", "C"+currentRowStr, "C"+strconv.Itoa(currentRow+1))
				f.SetCellValue("Sheet2", "D"+currentRowStr, "Регулирование")
				f.SetCellValue("Sheet2", "H"+currentRowStr, "1")
				currentRow++
				currentRowStr = strconv.Itoa(currentRow)
				f.SetCellValue("Sheet2", "D"+currentRowStr, "Измерение")
				f.SetCellValue("Sheet2", "F"+currentRowStr, "1")
			case "р":
				//Щит КН
				currentRow++
				currentRowStr = strconv.Itoa(currentRow)
				f.SetCellValue("Sheet2", "C"+currentRowStr, "Состояние вводного рубильника")
				f.MergeCell("Sheet2", "C"+currentRowStr, "C"+strconv.Itoa(currentRow+1))
				f.SetCellValue("Sheet2", "D"+currentRowStr, "Вкл/выкл")
				f.SetCellValue("Sheet2", "F"+currentRowStr, "1")
				currentRow++
				currentRowStr = strconv.Itoa(currentRow)
				f.SetCellValue("Sheet2", "D"+currentRowStr, "Норма/Авария")
				f.SetCellValue("Sheet2", "F"+currentRowStr, "1")
				//
				currentRow++
				currentRowStr = strconv.Itoa(currentRow)
				f.SetCellValue("Sheet2", "C"+currentRowStr, "Состояние реле контроля фаз")
				f.MergeCell("Sheet2", "C"+currentRowStr, "C"+strconv.Itoa(currentRow+1))
				f.SetCellValue("Sheet2", "D"+currentRowStr, "Вкл/выкл")
				f.SetCellValue("Sheet2", "F"+currentRowStr, "1")
				currentRow++
				currentRowStr = strconv.Itoa(currentRow)
				f.SetCellValue("Sheet2", "D"+currentRowStr, "Норма/Авария")
				f.SetCellValue("Sheet2", "F"+currentRowStr, "1")
				currentRow++
				currentRowStr = strconv.Itoa(currentRow)
				f.SetCellValue("Sheet2", "C"+currentRowStr, "Общий статус состояния групповых авт. выключателей")
				f.MergeCell("Sheet2", "C"+currentRowStr, "C"+strconv.Itoa(currentRow+1))
				f.SetCellValue("Sheet2", "D"+currentRowStr, "Вкл/выкл")
				f.SetCellValue("Sheet2", "F"+currentRowStr, "1")
				currentRow++
				currentRowStr = strconv.Itoa(currentRow)
				f.SetCellValue("Sheet2", "D"+currentRowStr, "Норма/Авария")
				f.SetCellValue("Sheet2", "F"+currentRowStr, "1")

			case "гвсм":
				//Заполняем таблицу по типу 1
				currentRow++

				currentRowStr = strconv.Itoa(currentRow)
				f.SetCellValue("Sheet2", "C"+currentRowStr, "Управление отсечным клапаном ХВС")
				f.MergeCell("Sheet2", "C"+currentRowStr, "C"+strconv.Itoa(currentRow+2))
				f.SetCellValue("Sheet2", "D"+currentRowStr, "Регулирование")
				f.SetCellValue("Sheet2", "H"+currentRowStr, "1")
				currentRow++
				currentRowStr = strconv.Itoa(currentRow)
				f.SetCellValue("Sheet2", "D"+currentRowStr, "Измерение")
				f.SetCellValue("Sheet2", "E"+currentRowStr, "1")
				currentRow++
				currentRowStr = strconv.Itoa(currentRow)
				f.SetCellValue("Sheet2", "D"+currentRowStr, "Норма/Авария")
				f.SetCellValue("Sheet2", "F"+currentRowStr, "1")
				currentRow++
				currentRowStr = strconv.Itoa(currentRow)
				f.SetCellValue("Sheet2", "C"+currentRowStr, "Управление отсечным клапаном ГВС")
				f.MergeCell("Sheet2", "C"+currentRowStr, "C"+strconv.Itoa(currentRow+2))
				f.SetCellValue("Sheet2", "D"+currentRowStr, "Регулирование")
				f.SetCellValue("Sheet2", "H"+currentRowStr, "1")
				currentRow++
				currentRowStr = strconv.Itoa(currentRow)
				f.SetCellValue("Sheet2", "D"+currentRowStr, "Измерение")
				f.SetCellValue("Sheet2", "E"+currentRowStr, "1")
				currentRow++
				currentRowStr = strconv.Itoa(currentRow)
				f.SetCellValue("Sheet2", "D"+currentRowStr, "Норма/Авария")
				f.SetCellValue("Sheet2", "F"+currentRowStr, "1")
				currentRow++
				currentRowStr = strconv.Itoa(currentRow)
				f.SetCellValue("Sheet2", "C"+currentRowStr, "Протечка")
				f.SetCellValue("Sheet2", "D"+currentRowStr, "Норма/Авария")
				f.SetCellValue("Sheet2", "F"+currentRowStr, "1")
				//Modbus
				currentRow++
				currentRow++
				currentRow++
				currentRow++
				currentRowStr = strconv.Itoa(currentRow)
				f.SetCellValue("Sheet2", "C"+currentRowStr, "Счетчик воды 1")
				f.SetCellValue("Sheet2", "D"+currentRowStr, "Измерение")
				f.SetCellValue("Sheet2", "E"+currentRowStr, "1")
				currentRow++
				currentRowStr = strconv.Itoa(currentRow)
				f.SetCellValue("Sheet2", "C"+currentRowStr, "Счетчик воды 2")
				f.SetCellValue("Sheet2", "D"+currentRowStr, "Измерение")
				f.SetCellValue("Sheet2", "E"+currentRowStr, "1")
				currentRow++
				currentRowStr = strconv.Itoa(currentRow)
				f.SetCellValue("Sheet2", "C"+currentRowStr, "Датчик протечки воды 1")
				f.SetCellValue("Sheet2", "D"+currentRowStr, "Норма/Авария")
				f.SetCellValue("Sheet2", "F"+currentRowStr, "1")
				currentRow++
				currentRowStr = strconv.Itoa(currentRow)
				f.SetCellValue("Sheet2", "C"+currentRowStr, "Датчик протечки воды 2")
				f.SetCellValue("Sheet2", "D"+currentRowStr, "Норма/Авария")
				f.SetCellValue("Sheet2", "F"+currentRowStr, "1")
				currentRow++
				currentRowStr = strconv.Itoa(currentRow)
				f.SetCellValue("Sheet2", "C"+currentRowStr, "Давление с датчика давления 1")
				f.SetCellValue("Sheet2", "D"+currentRowStr, "Измерение")
				f.SetCellValue("Sheet2", "E"+currentRowStr, "1")
				currentRow++
				currentRowStr = strconv.Itoa(currentRow)
				f.SetCellValue("Sheet2", "C"+currentRowStr, "Давление с датчика давления 2")
				f.SetCellValue("Sheet2", "D"+currentRowStr, "Измерение")
				f.SetCellValue("Sheet2", "E"+currentRowStr, "1")
			case "р+":
				//Щит КФ

				currentRow++
				currentRowStr = strconv.Itoa(currentRow)
				f.SetCellValue("Sheet2", "C"+currentRowStr, "Состояние вводного рубильника")
				f.MergeCell("Sheet2", "C"+currentRowStr, "C"+strconv.Itoa(currentRow+1))
				f.SetCellValue("Sheet2", "D"+currentRowStr, "Вкл/выкл")
				f.SetCellValue("Sheet2", "F"+currentRowStr, "1")
				currentRow++
				currentRowStr = strconv.Itoa(currentRow)
				f.SetCellValue("Sheet2", "D"+currentRowStr, "Норма/Авария")
				f.SetCellValue("Sheet2", "F"+currentRowStr, "1")
				//

				currentRow++
				currentRowStr = strconv.Itoa(currentRow)
				f.SetCellValue("Sheet2", "C"+currentRowStr, "Состояние реле контроля фаз")
				f.MergeCell("Sheet2", "C"+currentRowStr, "C"+strconv.Itoa(currentRow+1))
				f.SetCellValue("Sheet2", "D"+currentRowStr, "Вкл/выкл")
				f.SetCellValue("Sheet2", "F"+currentRowStr, "1")
				currentRow++
				currentRowStr = strconv.Itoa(currentRow)
				f.SetCellValue("Sheet2", "D"+currentRowStr, "Норма/Авария")
				f.SetCellValue("Sheet2", "F"+currentRowStr, "1")
				//
				currentRow++
				currentRowStr = strconv.Itoa(currentRow)
				f.SetCellValue("Sheet2", "C"+currentRowStr, "Общий статус состояния групповых авт. выключателей")
				f.MergeCell("Sheet2", "C"+currentRowStr, "C"+strconv.Itoa(currentRow+1))
				f.SetCellValue("Sheet2", "D"+currentRowStr, "Вкл/выкл")
				f.SetCellValue("Sheet2", "F"+currentRowStr, "1")
				currentRow++
				currentRowStr = strconv.Itoa(currentRow)
				f.SetCellValue("Sheet2", "D"+currentRowStr, "Норма/Авария")
				f.SetCellValue("Sheet2", "F"+currentRowStr, "1")
				currentRow++
				currentRowStr = strconv.Itoa(currentRow)
				f.SetCellValue("Sheet2", "C"+currentRowStr, "Статус состояния контакторов управления авар. освещением")
				//f.MergeCell("Sheet2", "C"+currentRowStr, "C"+strconv.Itoa(currentRow+1))
				f.SetCellValue("Sheet2", "D"+currentRowStr, "Вкл/выкл")
				f.SetCellValue("Sheet2", "F"+currentRowStr, "1")
			case "агпт":
				//агпт
				currentRow++
				currentRowStr = strconv.Itoa(currentRow)
				f.SetCellValue("Sheet2", "C"+currentRowStr, "Давление в баллоне газового пожаротушения мультиплексорной ГТИ")
				f.SetCellValue("Sheet2", "D"+currentRowStr, "Норма/Авария")
				f.SetCellValue("Sheet2", "F"+currentRowStr, "1")
			case "сс":
				//Заполняем таблицу по типу 1
				currentRow++
				currentRowStr = strconv.Itoa(currentRow)
				f.SetCellValue("Sheet2", "C"+currentRowStr, "Датчик протечки LS1")
				f.SetCellValue("Sheet2", "D"+currentRowStr, "Норма/Авария")
				f.SetCellValue("Sheet2", "F"+currentRowStr, "1")
			case "б":
				//Щит Б
				currentRow++
				currentRowStr = strconv.Itoa(currentRow)
				f.SetCellValue("Sheet2", "C"+currentRowStr, "Рубильник Питание ИБП")
				f.MergeCell("Sheet2", "C"+currentRowStr, "C"+strconv.Itoa(currentRow+1))
				f.SetCellValue("Sheet2", "D"+currentRowStr, "Вкл/выкл")
				f.SetCellValue("Sheet2", "F"+currentRowStr, "1")

				currentRow++
				currentRowStr = strconv.Itoa(currentRow)
				f.SetCellValue("Sheet2", "D"+currentRowStr, "Норма/Авария")
				f.SetCellValue("Sheet2", "F"+currentRowStr, "1")
				currentRow++
				currentRowStr = strconv.Itoa(currentRow)
				f.SetCellValue("Sheet2", "C"+currentRowStr, "Байпасный рубильник Позиция I Питание от ИБП")
				f.SetCellValue("Sheet2", "D"+currentRowStr, "Вкл/выкл")
				f.SetCellValue("Sheet2", "F"+currentRowStr, "1")
				currentRow++
				currentRowStr = strconv.Itoa(currentRow)
				f.SetCellValue("Sheet2", "C"+currentRowStr, "Байпасный рубильник Позиция II Питание по байпасу")
				f.MergeCell("Sheet2", "C"+currentRowStr, "C"+strconv.Itoa(currentRow+1))
				f.SetCellValue("Sheet2", "D"+currentRowStr, "Вкл/выкл")
				f.SetCellValue("Sheet2", "F"+currentRowStr, "1")
				currentRow++
				currentRowStr = strconv.Itoa(currentRow)
				f.SetCellValue("Sheet2", "D"+currentRowStr, "Норма/Авария")
				f.SetCellValue("Sheet2", "F"+currentRowStr, "1")
				//
				//
				//
				//currentRow++
				//currentRowStr = strconv.Itoa(currentRow)
				//f.SetCellValue("Sheet2", "C"+currentRowStr, "ИБП в норме")
				//f.SetCellValue("Sheet2", "D"+currentRowStr, "Норма/Авария")
				//f.SetCellValue("Sheet2", "F"+currentRowStr, "1")
				//
				//
				//currentRow++
				//currentRowStr = strconv.Itoa(currentRow)
				//f.SetCellValue("Sheet2", "C"+currentRowStr, "Работа от инвертора")
				//f.SetCellValue("Sheet2", "D"+currentRowStr, "Норма/Авария")
				//f.SetCellValue("Sheet2", "F"+currentRowStr, "1")
				//currentRow++
				//currentRowStr = strconv.Itoa(currentRow)
				//f.SetCellValue("Sheet2", "C"+currentRowStr, "Низкий заряд батареи")
				//f.SetCellValue("Sheet2", "D"+currentRowStr, "Норма/Авария")
				//f.SetCellValue("Sheet2", "F"+currentRowStr, "1")
				//currentRow++
				//currentRowStr = strconv.Itoa(currentRow)
				//f.SetCellValue("Sheet2", "C"+currentRowStr, "Общая авария ИБП")
				//f.SetCellValue("Sheet2", "D"+currentRowStr, "Норма/Авария")
				//f.SetCellValue("Sheet2", "F"+currentRowStr, "1")
			case "ибп":
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
			case "ув":
				//Узел ввода УВ
				//
				currentRow++
				currentRowStr = strconv.Itoa(currentRow)
				f.SetCellValue("Sheet2", "C"+currentRowStr, "Температура охлажденной воды")
				f.MergeCell("Sheet2", "C"+currentRowStr, "C"+strconv.Itoa(currentRow+1))
				f.SetCellValue("Sheet2", "D"+currentRowStr, "Значение температуры")
				f.SetCellValue("Sheet2", "E"+currentRowStr, "1")
				currentRow++
				currentRowStr = strconv.Itoa(currentRow)
				f.SetCellValue("Sheet2", "D"+currentRowStr, "Норма/Авария")
				f.SetCellValue("Sheet2", "F"+currentRowStr, "1")
				//
				currentRow++
				currentRowStr = strconv.Itoa(currentRow)
				f.SetCellValue("Sheet2", "C"+currentRowStr, "Температура нагретой воды")
				f.MergeCell("Sheet2", "C"+currentRowStr, "C"+strconv.Itoa(currentRow+1))
				f.SetCellValue("Sheet2", "D"+currentRowStr, "Значение температуры")
				f.SetCellValue("Sheet2", "E"+currentRowStr, "1")
				currentRow++
				currentRowStr = strconv.Itoa(currentRow)
				f.SetCellValue("Sheet2", "D"+currentRowStr, "Норма/Авария")
				f.SetCellValue("Sheet2", "F"+currentRowStr, "1")

				currentRow++
				currentRowStr = strconv.Itoa(currentRow)
				f.SetCellValue("Sheet2", "C"+currentRowStr, "Давление охлажденной воды")
				f.MergeCell("Sheet2", "C"+currentRowStr, "C"+strconv.Itoa(currentRow+1))
				f.SetCellValue("Sheet2", "D"+currentRowStr, "Значение давления")
				f.SetCellValue("Sheet2", "E"+currentRowStr, "1")
				currentRow++
				currentRowStr = strconv.Itoa(currentRow)
				f.SetCellValue("Sheet2", "D"+currentRowStr, "Норма/Авария")
				f.SetCellValue("Sheet2", "F"+currentRowStr, "1")

				currentRow++
				currentRowStr = strconv.Itoa(currentRow)
				f.SetCellValue("Sheet2", "C"+currentRowStr, "Давление нагретой воды")
				f.MergeCell("Sheet2", "C"+currentRowStr, "C"+strconv.Itoa(currentRow+1))
				f.SetCellValue("Sheet2", "D"+currentRowStr, "Значение давления")
				f.SetCellValue("Sheet2", "E"+currentRowStr, "1")
				currentRow++
				currentRowStr = strconv.Itoa(currentRow)
				f.SetCellValue("Sheet2", "D"+currentRowStr, "Норма/Авария")
				f.SetCellValue("Sheet2", "F"+currentRowStr, "1")

			case "лист":
				//Listcontroller

				currentRow++
				currentRowStr = strconv.Itoa(currentRow)
				f.SetCellValue("Sheet2", "C"+currentRowStr, "Температура в кабельном лотке")
				f.SetCellValue("Sheet2", "D"+currentRowStr, "Значение температуры")
				f.SetCellValue("Sheet2", "E"+currentRowStr, "1")
				currentRow++
				currentRowStr = strconv.Itoa(currentRow)
				f.SetCellValue("Sheet2", "C"+currentRowStr, "Обнаружение очага возгорания кабельных трасс")
				f.SetCellValue("Sheet2", "D"+currentRowStr, "Норма/Авария")
				f.SetCellValue("Sheet2", "F"+currentRowStr, "1")
			case "мульт":
				//Заполняем таблицу по типу 1
				currentRow++
				currentRowStr = strconv.Itoa(currentRow)
				f.SetCellValue("Sheet2", "C"+currentRowStr, "Давление в балоне газового пожаротушения в мультиплексорной ГТИ")
				f.MergeCell("Sheet2", "C"+currentRowStr, "C"+strconv.Itoa(currentRow+1))
				f.SetCellValue("Sheet2", "D"+currentRowStr, "Значение давления")
				f.SetCellValue("Sheet2", "E"+currentRowStr, "1")
				currentRow++
				currentRowStr = strconv.Itoa(currentRow)
				f.SetCellValue("Sheet2", "D"+currentRowStr, "Норма/Авария")
				f.SetCellValue("Sheet2", "F"+currentRowStr, "1")
				//
			case "осу":
				//Заполняем таблицу по типу 1
				currentRow++
				currentRow++
				currentRow++
				currentRow++
				currentRowStr = strconv.Itoa(currentRow)
				f.SetCellValue("Sheet2", "C"+currentRowStr, "Температура окружающего воздуха")
				f.SetCellValue("Sheet2", "D"+currentRowStr, "Значение температуры")
				f.SetCellValue("Sheet2", "E"+currentRowStr, "1")
				currentRow++
				currentRowStr = strconv.Itoa(currentRow)
				f.SetCellValue("Sheet2", "C"+currentRowStr, "Влажность окружающего воздуха")
				f.SetCellValue("Sheet2", "D"+currentRowStr, "Значение влажности")
				f.SetCellValue("Sheet2", "F"+currentRowStr, "1")
				currentRow++
				currentRowStr = strconv.Itoa(currentRow)
				f.SetCellValue("Sheet2", "C"+currentRowStr, "Уставка влажности")
				f.SetCellValue("Sheet2", "D"+currentRowStr, "Значение влажности")
				f.SetCellValue("Sheet2", "F"+currentRowStr, "1")
				currentRow++
				currentRowStr = strconv.Itoa(currentRow)
				f.SetCellValue("Sheet2", "C"+currentRowStr, "Уставка температуры")
				f.SetCellValue("Sheet2", "D"+currentRowStr, "Значение температуры")
				f.SetCellValue("Sheet2", "F"+currentRowStr, "1")
				currentRow++
				currentRowStr = strconv.Itoa(currentRow)
				f.SetCellValue("Sheet2", "C"+currentRowStr, "Выход оповещения 1")
				f.SetCellValue("Sheet2", "D"+currentRowStr, "Вкл/выкл")
				f.SetCellValue("Sheet2", "F"+currentRowStr, "1")
				currentRow++
				currentRowStr = strconv.Itoa(currentRow)
				f.SetCellValue("Sheet2", "C"+currentRowStr, "Выход оповещения 2")
				f.SetCellValue("Sheet2", "D"+currentRowStr, "Вкл/выкл")
				f.SetCellValue("Sheet2", "F"+currentRowStr, "1")
				currentRow++
				currentRowStr = strconv.Itoa(currentRow)
				f.SetCellValue("Sheet2", "C"+currentRowStr, "Выход оповещения 1")
				f.SetCellValue("Sheet2", "D"+currentRowStr, "Вкл/выкл")
				f.SetCellValue("Sheet2", "F"+currentRowStr, "1")
				//
				currentRow++
				currentRowStr = strconv.Itoa(currentRow)
				f.SetCellValue("Sheet2", "C"+currentRowStr, "Состояние режима Fail start")
				f.SetCellValue("Sheet2", "D"+currentRowStr, "Норма/Авария")
				f.SetCellValue("Sheet2", "F"+currentRowStr, "1")
				currentRow++
				currentRowStr = strconv.Itoa(currentRow)
				f.SetCellValue("Sheet2", "C"+currentRowStr, "Состояние режима готовности")
				f.SetCellValue("Sheet2", "D"+currentRowStr, "Норма/Авария")
				f.SetCellValue("Sheet2", "F"+currentRowStr, "1")
				currentRow++
				currentRowStr = strconv.Itoa(currentRow)
				f.SetCellValue("Sheet2", "C"+currentRowStr, "Состояние осушения")
				f.SetCellValue("Sheet2", "D"+currentRowStr, "Норма/Авария")
				f.SetCellValue("Sheet2", "F"+currentRowStr, "1")
				currentRow++
				currentRowStr = strconv.Itoa(currentRow)
				f.SetCellValue("Sheet2", "C"+currentRowStr, "Состояние оттаивания")
				f.SetCellValue("Sheet2", "D"+currentRowStr, "Норма/Авария")
				f.SetCellValue("Sheet2", "F"+currentRowStr, "1")
				currentRow++
				currentRowStr = strconv.Itoa(currentRow)
				f.SetCellValue("Sheet2", "C"+currentRowStr, "Состояние режима Чрезмерное низкое давление")
				f.SetCellValue("Sheet2", "D"+currentRowStr, "Норма/Авария")
				f.SetCellValue("Sheet2", "F"+currentRowStr, "1")
				currentRow++
				currentRowStr = strconv.Itoa(currentRow)
				f.SetCellValue("Sheet2", "C"+currentRowStr, "Неисправность датчика")
				f.SetCellValue("Sheet2", "D"+currentRowStr, "Норма/Авария")
				f.SetCellValue("Sheet2", "F"+currentRowStr, "1")
				currentRow++
				currentRowStr = strconv.Itoa(currentRow)
				f.SetCellValue("Sheet2", "C"+currentRowStr, "Состояние режима Чрезмерное высокое давление")
				f.SetCellValue("Sheet2", "D"+currentRowStr, "Норма/Авария")
				f.SetCellValue("Sheet2", "F"+currentRowStr, "1")
				currentRow++
				currentRowStr = strconv.Itoa(currentRow)
				f.SetCellValue("Sheet2", "C"+currentRowStr, "Внешние условия вне рабочего диапазона")
				f.SetCellValue("Sheet2", "D"+currentRowStr, "Норма/Авария")
				f.SetCellValue("Sheet2", "F"+currentRowStr, "1")
				currentRow++
				currentRowStr = strconv.Itoa(currentRow)
				f.SetCellValue("Sheet2", "C"+currentRowStr, "Температура окружающей среды вне рабочего диапазона")
				f.SetCellValue("Sheet2", "D"+currentRowStr, "Норма/Авария")
				f.SetCellValue("Sheet2", "F"+currentRowStr, "1")
				currentRow++
				currentRowStr = strconv.Itoa(currentRow)
				f.SetCellValue("Sheet2", "C"+currentRowStr, "Влажность окружающей среды вне рабочего диапазона")
				f.SetCellValue("Sheet2", "D"+currentRowStr, "Норма/Авария")
				f.SetCellValue("Sheet2", "F"+currentRowStr, "1")
				//
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
	} else {
		//НАЧАЛО ИМПОРТА ЕТС
		f, err := excelize.OpenFile("Liis_office_1nd_floor.xlsx")
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
		cell, err := f.GetCellValue("Лист 1 - Liis_office_1nd_floor.", "A1")
		if err != nil {
			fmt.Println(err)
			return
		}
		fmt.Println(cell)

		// Получить все строки в Sheet1
		rows, err := f.GetRows("Лист 1 - Liis_office_1nd_floor.")
		if err != nil {
			fmt.Println(err)
			return
		}

		//rows[22][7]
		//3 - GA
		//2 - name
		//7 - DPT-1

		outputdata := ""

		/*
					{
			            "name": "Floor3_OpenSpace_Left_Light_Status_on/off",
			            "dpt": "1.001",
			            "addresses":["5/1/3"],
			            "control": 0
			        },
			        {
			            "name": "Floor3_OpenSpace_Left_Level_status_light",
			            "dpt": "5.001",
			            "addresses":["5/3/3"],
			            "control": 0
			        },
		*/

		for _, row := range rows {
			if len(row) > 3 {
				if row[2] != "" && strings.Contains(row[3], "-") == false {
					dptstr := ""
					if row[7] == "DPT-1" {
						dptstr = "1.001"
					}
					if row[7] == "DPT-5" {
						dptstr = "5.001"
					}
					if dptstr != "" {
						preparingString := fmt.Sprintf("{\n\"name\": \"%v\",\n\"dpt\": \"%v\",\n\"addresses\":[\"%v\"],\n\"control\": \"0\"\n},\n", row[2], dptstr, row[3])
						outputdata += preparingString
					}

				}
			}
		}
		fo, err := os.Create("sample.file")
		if err != nil {
			panic(err)
		}
		defer f.Close()

		_, err = fo.WriteString(outputdata)
		if err != nil {
			panic(err)
		}
	}
}
