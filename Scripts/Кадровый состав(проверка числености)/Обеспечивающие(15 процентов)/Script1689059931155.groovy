import static com.kms.katalon.core.checkpoint.CheckpointFactory.findCheckpoint
import com.kms.katalon.keyword.excel.ExcelKeywords as ExcelKeywords
import static com.kms.katalon.core.testcase.TestCaseFactory.findTestCase
import static com.kms.katalon.core.testdata.TestDataFactory.findTestData
import static com.kms.katalon.core.testobject.ObjectRepository.findTestObject
import static com.kms.katalon.core.testobject.ObjectRepository.findWindowsObject
import com.kms.katalon.core.checkpoint.Checkpoint as Checkpoint
import com.kms.katalon.core.cucumber.keyword.CucumberBuiltinKeywords as CucumberKW
import com.kms.katalon.core.mobile.keyword.MobileBuiltInKeywords as Mobile
import com.kms.katalon.core.model.FailureHandling as FailureHandling
import com.kms.katalon.core.testcase.TestCase as TestCase
import com.kms.katalon.core.testdata.TestData as TestData
import com.kms.katalon.core.testng.keyword.TestNGBuiltinKeywords as TestNGKW
import com.kms.katalon.core.testobject.TestObject as TestObject
import com.kms.katalon.core.webservice.keyword.WSBuiltInKeywords as WS
import com.kms.katalon.core.webui.keyword.WebUiBuiltInKeywords as WebUI
import com.kms.katalon.core.windows.keyword.WindowsBuiltinKeywords as Windows
import internal.GlobalVariable as GlobalVariable
import org.openqa.selenium.Keys as Keys
import java.nio.file.Path as Path
import java.nio.file.Paths as Paths

WebUI.openBrowser(findTestData('New Test Data').getValue(4, 1))

WebUI.setText(findTestObject('Page_/input__username'), findTestData('New Test Data').getValue(1, 1))

WebUI.setText(findTestObject('Page_/input__password'), findTestData('New Test Data').getValue(2, 1))

WebUI.click(findTestObject('Page_/button_'))

def resetFilters = ResetDataAndDzo()

WebUI.click(findTestObject('Кадровый состав/Фильтр ДЗО'))

WebUI.click(findTestObject('Кадровый состав/Обеспечивающие(раскрыть)'))

WebUI.click(findTestObject('Кадровый состав/НИЦ ЕЭС верхний уровень'))

WebUI.click(findTestObject('Кадровый состав/Применить в фильтре ДЗО'))

def data2021And2022 = Data()

resetFilters = ResetDataAndDzo()

WebUI.click(findTestObject('Кадровый состав/Фильтр ДЗО'))

WebUI.click(findTestObject('Кадровый состав/Снять выделение в фильтре ДЗО'))

WebUI.click(findTestObject('Кадровый состав/АО Россети Цифра'))

WebUI.click(findTestObject('Кадровый состав/Применить в фильтре ДЗО'))

data2021And2022 = Data()

resetFilters = ResetDataAndDzo()

WebUI.click(findTestObject('Кадровый состав/Фильтр ДЗО'))

WebUI.click(findTestObject('Кадровый состав/Снять выделение в фильтре ДЗО'))

WebUI.click(findTestObject('Кадровый состав/АО Центр технического заказчика'))

WebUI.click(findTestObject('Кадровый состав/Применить в фильтре ДЗО'))

data2021And2022 = Data()

resetFilters = ResetDataAndDzo()

WebUI.click(findTestObject('Кадровый состав/Фильтр ДЗО'))

WebUI.click(findTestObject('Кадровый состав/Снять выделение в фильтре ДЗО'))

WebUI.click(findTestObject('Кадровый состав/ДЗО ПАО ФИЦ'))

WebUI.click(findTestObject('Кадровый состав/Применить в фильтре ДЗО'))

data2021And2022 = Data()

resetFilters = ResetDataAndDzo()

WebUI.click(findTestObject('Кадровый состав/Фильтр ДЗО'))

WebUI.click(findTestObject('Кадровый состав/Снять выделение в фильтре ДЗО'))

WebUI.click(findTestObject('Кадровый состав/ПАО СЗЭУК'))

WebUI.click(findTestObject('Кадровый состав/Применить в фильтре ДЗО'))

data2021And2022 = Data()

WebUI.closeBrowser()

static def Srav(int numPerson2021, int numPerson2022) {
    int year21 = numPerson2021

    int year22 = numPerson2022

    println(year21)

    println(year22)

    if (year21 > year22) {
        int god1 = ((year21 * 100) / year22) - 100

        println(god1 + '%')

        if (god1 > 15) {
            def write = WriteToExcel(god1)

            println('Запуск функции для записи в ексель файл')
        } else if (god1 < 15) {
            println('Не превышает 15%')
        }
    } else if (year22 > year21) {
        int god1 = ((year22 * 100) / year21) - 100

        println(god1 + '%')

        if (god1 > 15) {
            def write = WriteToExcel(god1)

            println('Запуск функции для записи в ексель файл')
        } else if (god1 < 15) {
            println('Не превышает 15%')
        }
    }
}

static def WriteToExcel(def god) {
    String sheetName = 'List1'

    def data = findTestData('New Test Data')

    int n = data.getRowNumbers() + 1

    String dZO = WebUI.getText(findTestObject('Кадровый состав/Фильтр ДЗО'))

    String year = '2021-2022 год'

    println(year)

    String dashboardName = 'Кадровый состав'

    String percent = 'Разница больше 15%'

    String percent1 = ('Разница на ' + god) + '%'

    System.err.println('err')

    WebUI.verifyTextPresent('ОШИБКА ВЫВОДИТСЯ В ТОМ СЛУЧАЕ, ЕСЛИ РАЗНИЦА БОЛЬШЕ 15%', false)

    def workbook01 = ExcelKeywords.getWorkbook(GlobalVariable.excelFilePath)

    def sheet01 = ExcelKeywords.getExcelSheet(workbook01, sheetName)

    ExcelKeywords.setValueToCellByIndex(sheet01, n, 0, dashboardName)

    ExcelKeywords.setValueToCellByIndex(sheet01, n, 1, dZO)

    ExcelKeywords.setValueToCellByIndex(sheet01, n, 2, year)

    ExcelKeywords.setValueToCellByIndex(sheet01, n, 3, percent)

    ExcelKeywords.setValueToCellByIndex(sheet01, n, 4, percent1)

    n = (n + 1)

    ExcelKeywords.saveWorkbook(GlobalVariable.excelFilePath, workbook01)
}

static def ResetDataAndDzo() {
    WebUI.click(findTestObject('Кадровый состав/Фильтр Дата'))

    WebUI.click(findTestObject('Кадровый состав/Снять выделение в фильтре Дата'))

    WebUI.click(findTestObject('Кадровый состав/2021 год'))

    WebUI.click(findTestObject('Кадровый состав/Применить в фильтре Дата'))

    WebUI.click(findTestObject('Кадровый состав/Фильтр ДЗО'))

    WebUI.click(findTestObject('Кадровый состав/Снять выделение в фильтре ДЗО'))

    WebUI.click(findTestObject('Кадровый состав/Применить в фильтре ДЗО'))
}

static def Data() {
    WebUI.click(findTestObject('Кадровый состав/Фильтр Дата'))

    WebUI.click(findTestObject('Кадровый состав/Снять выделение в фильтре Дата'))

    WebUI.click(findTestObject('Кадровый состав/2021 год'))

    WebUI.click(findTestObject('Кадровый состав/Применить в фильтре Дата'))

    String strPerson2021 = WebUI.getText(findTestObject('Кадровый состав/Виджет общей численности персонала')).replaceAll(
        ' ', '')

    if (WebUI.getText(findTestObject('Кадровый состав/Виджет общей численности персонала')) == 'нет данных') {
        def writeNoData = WriteNoData(def noData = 'Нет данных за 2021 год')
    }
    
    int numPerson2021 = Integer.parseInt(strPerson2021)

    println(numPerson2021 + ' Данные за 2021 год')

    WebUI.click(findTestObject('Кадровый состав/Фильтр Дата'))

    WebUI.click(findTestObject('Кадровый состав/Снять выделение в фильтре Дата'))

    WebUI.click(findTestObject('Кадровый состав/2022 год'))

    WebUI.click(findTestObject('Кадровый состав/Применить в фильтре Дата'))

    String strPerson2022 = WebUI.getText(findTestObject('Кадровый состав/Виджет общей численности персонала')).replaceAll(
        ' ', '')

    if (WebUI.getText(findTestObject('Кадровый состав/Виджет общей численности персонала')) == 'нет данных') {
        def writeNoData = WriteNoData(def noData = 'Нет данных за 2022 год')
    } else {
        int numPerson2022 = Integer.parseInt(strPerson2022)

        println(numPerson2022 + ' Данные за 2022 год')

        def srv = Srav(numPerson2021, numPerson2022)
    }
}

static def WriteNoData(def noData) {
    String sheetName = 'List1'

    def data = findTestData('New Test Data')

    int n = data.getRowNumbers() + 1

    String dZO = WebUI.getText(findTestObject('Кадровый состав/Фильтр ДЗО'))

    String year = WebUI.getText(findTestObject('Кадровый состав/Фильтр Дата'))

    println(year)

    String dashboardName = 'Кадровый состав'

    String typeErr = noData

    WebUI.verifyTextPresent('ОШИБКА ВЫВОДИТСЯ В ТОМ СЛУЧАЕ, ЕСЛИ НЕТ ДАННЫХ В ВИДЖЕТЕ', false)

    def workbook01 = ExcelKeywords.getWorkbook(GlobalVariable.excelFilePath)

    def sheet01 = ExcelKeywords.getExcelSheet(workbook01, sheetName)

    ExcelKeywords.setValueToCellByIndex(sheet01, n, 0, dashboardName)

    ExcelKeywords.setValueToCellByIndex(sheet01, n, 1, dZO)

    ExcelKeywords.setValueToCellByIndex(sheet01, n, 2, year)

    ExcelKeywords.setValueToCellByIndex(sheet01, n, 3, typeErr)

    n = (n + 1)

    ExcelKeywords.saveWorkbook(GlobalVariable.excelFilePath, workbook01)
}

