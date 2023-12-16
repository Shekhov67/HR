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

WebUI.openBrowser(findTestData('New Test Data').getValue(3, 5))

WebUI.setText(findTestObject('Page_/input__username'), findTestData('New Test Data').getValue(1, 1))

WebUI.setText(findTestObject('Page_/input__password'), findTestData('New Test Data').getValue(2, 1))

WebUI.click(findTestObject('Page_/button_'))

WebUI.click(findTestObject('Расходы на персонал/Фильтр Дата'))

WebUI.click(findTestObject('Расходы на персонал/Снять выделение в фильтре Дата'))

WebUI.click(findTestObject('Расходы на персонал/div-2020'))

WebUI.scrollToElement(findTestObject('Расходы на персонал/2020-1'), 30)

WebUI.scrollToElement(findTestObject('Расходы на персонал/a_DZO'), 30)

WebUI.click(findTestObject('Расходы на персонал/2020-1'))

WebUI.click(findTestObject('Расходы на персонал/Применить в фильтре Дата'))

WebUI.click(findTestObject('Расходы на персонал/фильтр Деятельность ДЗО'))

WebUI.click(findTestObject('Расходы на персонал/деятельность КАПИТАЛЬНОЕ СТРОИТЕЛЬСТРВО'))

WebUI.click(findTestObject('Расходы на персонал/применить в фильтре деятельность'))

WebUI.click(findTestObject('Расходы на персонал/Фильтр ДЗО'))

WebUI.click(findTestObject('Расходы на персонал/Снять выделение в фильтре ДЗО'))

WebUI.click(findTestObject('Расходы на персонал/Обеспечивающие'))

WebUI.click(findTestObject('Расходы на персонал/Применить в фильтре ДЗО'))

def write = WriteToExcel()

WebUI.click(findTestObject('Расходы на персонал/Фильтр ДЗО'))

WebUI.click(findTestObject('Расходы на персонал/Снять выделение в фильтре ДЗО'))

WebUI.click(findTestObject('Расходы на персонал/Развернуть Обеспечивающие'))

WebUI.click(findTestObject('Расходы на персонал/НИЦ ЕЭС верхний уровень'))

WebUI.click(findTestObject('Расходы на персонал/Применить в фильтре ДЗО'))

write = WriteToExcel()

WebUI.click(findTestObject('Расходы на персонал/Фильтр ДЗО'))

WebUI.click(findTestObject('Расходы на персонал/Снять выделение в фильтре ДЗО'))

WebUI.click(findTestObject('Общие элементы/Развернуть НИЦ ЕЭС'))

WebUI.click(findTestObject('Расходы на персонал/НИЦ ЕЭС нижний уровень'))

WebUI.click(findTestObject('Расходы на персонал/Применить в фильтре ДЗО'))

write = WriteToExcel()

WebUI.click(findTestObject('Расходы на персонал/Фильтр ДЗО'))

WebUI.click(findTestObject('Расходы на персонал/Снять выделение в фильтре ДЗО'))

WebUI.click(findTestObject('Расходы на персонал/Центр технического заказчика'))

WebUI.click(findTestObject('Расходы на персонал/Применить в фильтре ДЗО'))

write = WriteToExcel()

WebUI.click(findTestObject('Расходы на персонал/Фильтр ДЗО'))

WebUI.click(findTestObject('Расходы на персонал/Снять выделение в фильтре ДЗО'))

WebUI.click(findTestObject('Расходы на персонал/Развернуть Центр технического заказчика'))

WebUI.click(findTestObject('Расходы на персонал/Центр технического заказчика нижний уровень'))

WebUI.click(findTestObject('Расходы на персонал/Применить в фильтре ДЗО'))

write = WriteToExcel()

WebUI.click(findTestObject('Расходы на персонал/Фильтр ДЗО'))

WebUI.click(findTestObject('Расходы на персонал/Снять выделение в фильтре ДЗО'))

WebUI.click(findTestObject('Расходы на персонал/ДЗО ПАО ФИЦ'))

WebUI.click(findTestObject('Расходы на персонал/Применить в фильтре ДЗО'))

write = WriteToExcel()

WebUI.click(findTestObject('Расходы на персонал/Фильтр ДЗО'))

WebUI.click(findTestObject('Расходы на персонал/Снять выделение в фильтре ДЗО'))

WebUI.scrollToElement(findTestObject('Расходы на персонал/ДЗО ПАО ФИЦ'), 30)

WebUI.scrollToElement(findTestObject('Расходы на персонал/a_DZO'), 30)

WebUI.click(findTestObject('Общие элементы/Развернуть ДЗО ПАО ФИЦ'))

WebUI.click(findTestObject('Общие элементы/Оператор АСТУ ДЗО ПАО ФИЦ'))

WebUI.scrollToElement(findTestObject('Расходы на персонал/a_DZO'), 30)

WebUI.click(findTestObject('Расходы на персонал/Применить в фильтре ДЗО'))

write = WriteToExcel()

WebUI.click(findTestObject('Расходы на персонал/Фильтр ДЗО'))

WebUI.click(findTestObject('Расходы на персонал/Снять выделение в фильтре ДЗО'))

WebUI.scrollToElement(findTestObject('Расходы на персонал/Развернуть ПАО СЗЭУК2'), 30)

WebUI.scrollToElement(findTestObject('Расходы на персонал/a_DZO'), 30)

WebUI.click(findTestObject('Расходы на персонал/выбрать верхний уровень ПАО СЗЭУК'))

WebUI.click(findTestObject('Расходы на персонал/Применить в фильтре ДЗО'))

write = WriteToExcel()

WebUI.click(findTestObject('Расходы на персонал/Фильтр ДЗО'))

WebUI.click(findTestObject('Расходы на персонал/Снять выделение в фильтре ДЗО'))

WebUI.scrollToElement(findTestObject('Расходы на персонал/Развернуть ПАО СЗЭУК2'), 30)

WebUI.scrollToElement(findTestObject('Расходы на персонал/a_DZO'), 30)

WebUI.click(findTestObject('Расходы на персонал/Развернуть ПАО СЗЭУК2'))

WebUI.click(findTestObject('Расходы на персонал/выбрать ПАО СЗЭУК'))

WebUI.scrollToElement(findTestObject('Расходы на персонал/a_DZO'), 30)

WebUI.click(findTestObject('Расходы на персонал/Применить в фильтре ДЗО'))

write = WriteToExcel()

WebUI.click(findTestObject('Расходы на персонал/Фильтр ДЗО'))

WebUI.click(findTestObject('Расходы на персонал/Снять выделение в фильтре ДЗО'))

WebUI.scrollToElement(findTestObject('Общие элементы/Развернуть Управление ВОЛС-ВЛ'), 30)

WebUI.scrollToElement(findTestObject('Расходы на персонал/a_DZO'), 30)

WebUI.click(findTestObject('Общие элементы/Развернуть Управление ВОЛС-ВЛ'))

WebUI.click(findTestObject('Общие элементы/Управление ВОЛС-ВЛ нижний уровень'))

WebUI.scrollToElement(findTestObject('Расходы на персонал/a_DZO'), 30)

WebUI.click(findTestObject('Расходы на персонал/Применить в фильтре ДЗО'))

write = WriteToExcel()

WebUI.click(findTestObject('Расходы на персонал/Фильтр ДЗО'))

WebUI.click(findTestObject('Расходы на персонал/Снять выделение в фильтре ДЗО'))

WebUI.scrollToElement(findTestObject('Общие элементы/Развернуть Управление ВОЛС-ВЛ'), 30)

WebUI.scrollToElement(findTestObject('Расходы на персонал/a_DZO'), 30)

WebUI.click(findTestObject('Расходы на персонал/Верхний уровень Управление ВОЛС-ВЛ'))

write = WriteToExcel()

WebUI.closeBrowser( //Необходимо указать название листа
    )

static def WriteToExcel() {
    String excelFilePath = 'C:\\Users\\shehs\\OneDrive\\Desktop\\HR\\Data Files\\HR.xlsx'

    String sheetName = 'List1'

    def data = findTestData('New Test Data')

    int n = data.getRowNumbers() + 1

    String dZO = WebUI.getText(findTestObject('Расходы на персонал/Фильтр ДЗО'))

    String deitel = WebUI.getText(findTestObject('Расходы на персонал/фильтр Деятельность ДЗО'))

    String date = WebUI.getText(findTestObject('Расходы на персонал/Фильтр Дата'))

    println(date)

    date = date.replaceAll('[a-zA-Zа-яА-Я]*', '')

    date = date.replaceAll('\\s*', '')

    date = date.replaceAll('\\pP', '')

    println(date)

    String y1 = date.charAt(0)

    String y2 = date.charAt(1)

    String y3 = date.charAt(2)

    String y4 = date.charAt(3)

    String year = ((y1 + y2) + y3) + y4

    int y = year.toInteger()

    println(y)

    String q1

    if (date.length() == 5) {
        q1 = date.charAt(4)

        int q = q1.toInteger()

        println(q)

        String dashboardName = 'Расходы на персонал'

        def workbook01 = ExcelKeywords.getWorkbook(excelFilePath)

        def sheet01 = ExcelKeywords.getExcelSheet(workbook01, sheetName)

        if (WebUI.verifyTextNotPresent('нет данных', false) == false) {
            ExcelKeywords.setValueToCellByIndex(sheet01, n, 3, dashboardName)

            ExcelKeywords.setValueToCellByIndex(sheet01, n, 4, dZO)

            ExcelKeywords.setValueToCellByIndex(sheet01, n, 5, y)

            ExcelKeywords.setValueToCellByIndex(sheet01, n, 6, q)

            ExcelKeywords.setValueToCellByIndex(sheet01, n, 7, deitel)

            n = (n + 1)

            ExcelKeywords.saveWorkbook(excelFilePath, workbook01)
        } else {
            if (WebUI.verifyTextNotPresent('Ошибка запроса данных', false) == false) {
                ExcelKeywords.setValueToCellByIndex(sheet01, n, 3, dashboardName)

                ExcelKeywords.setValueToCellByIndex(sheet01, n, 4, dZO)

                ExcelKeywords.setValueToCellByIndex(sheet01, n, 5, y)

                ExcelKeywords.setValueToCellByIndex(sheet01, n, 6, q)

                ExcelKeywords.setValueToCellByIndex(sheet01, n, 7, deitel)

                n = (n + 1)

                ExcelKeywords.saveWorkbook(excelFilePath, workbook01)
            } else {
                if (WebUI.verifyTextNotPresent('Произошла ошибка при выполнении пользовательского кода', false) == false) {
                    ExcelKeywords.setValueToCellByIndex(sheet01, n, 3, dashboardName)

                    ExcelKeywords.setValueToCellByIndex(sheet01, n, 4, dZO)

                    ExcelKeywords.setValueToCellByIndex(sheet01, n, 5, y)

                    ExcelKeywords.setValueToCellByIndex(sheet01, n, 6, q)

                    ExcelKeywords.setValueToCellByIndex(sheet01, n, 7, deitel)

                    n = (n + 1)

                    ExcelKeywords.saveWorkbook(excelFilePath, workbook01)
                } else {
                    if (WebUI.verifyTextNotPresent('У виджета нет данных', false) == false) {
                        ExcelKeywords.setValueToCellByIndex(sheet01, n, 3, dashboardName)

                        ExcelKeywords.setValueToCellByIndex(sheet01, n, 4, dZO)

                        ExcelKeywords.setValueToCellByIndex(sheet01, n, 5, y)

                        ExcelKeywords.setValueToCellByIndex(sheet01, n, 6, q)

                        ExcelKeywords.setValueToCellByIndex(sheet01, n, 7, deitel)

                        n = (n + 1)

                        ExcelKeywords.saveWorkbook(excelFilePath, workbook01)
                    } else {
                        if (WebUI.verifyTextNotPresent('Некорректные фильтры', false) == false) {
                            ExcelKeywords.setValueToCellByIndex(sheet01, n, 3, dashboardName)

                            ExcelKeywords.setValueToCellByIndex(sheet01, n, 4, dZO)

                            ExcelKeywords.setValueToCellByIndex(sheet01, n, 5, y)

                            ExcelKeywords.setValueToCellByIndex(sheet01, n, 6, q)

                            ExcelKeywords.setValueToCellByIndex(sheet01, n, 7, deitel)

                            n = (n + 1)

                            ExcelKeywords.saveWorkbook(excelFilePath, workbook01)
                        }
                    }
                }
            }
        }
    } else {
        String q = ''

        println(q)

        String dashboardName = 'Расходы на персонал'

        def workbook01 = ExcelKeywords.getWorkbook(excelFilePath)

        def sheet01 = ExcelKeywords.getExcelSheet(workbook01, sheetName)

        if (WebUI.verifyTextNotPresent('нет данных', false) == false) {
            ExcelKeywords.setValueToCellByIndex(sheet01, n, 3, dashboardName)

            ExcelKeywords.setValueToCellByIndex(sheet01, n, 4, dZO)

            ExcelKeywords.setValueToCellByIndex(sheet01, n, 5, y)

            ExcelKeywords.setValueToCellByIndex(sheet01, n, 6, q)

            ExcelKeywords.setValueToCellByIndex(sheet01, n, 7, deitel)

            n = (n + 1)

            ExcelKeywords.saveWorkbook(excelFilePath, workbook01)
        } else {
            if (WebUI.verifyTextNotPresent('Ошибка запроса данных', false) == false) {
                ExcelKeywords.setValueToCellByIndex(sheet01, n, 3, dashboardName)

                ExcelKeywords.setValueToCellByIndex(sheet01, n, 4, dZO)

                ExcelKeywords.setValueToCellByIndex(sheet01, n, 5, y)

                ExcelKeywords.setValueToCellByIndex(sheet01, n, 6, q)

                ExcelKeywords.setValueToCellByIndex(sheet01, n, 7, deitel)

                n = (n + 1)

                ExcelKeywords.saveWorkbook(excelFilePath, workbook01)
            } else {
                if (WebUI.verifyTextNotPresent('Произошла ошибка при выполнении пользовательского кода', false) == false) {
                    ExcelKeywords.setValueToCellByIndex(sheet01, n, 3, dashboardName)

                    ExcelKeywords.setValueToCellByIndex(sheet01, n, 4, dZO)

                    ExcelKeywords.setValueToCellByIndex(sheet01, n, 5, y)

                    ExcelKeywords.setValueToCellByIndex(sheet01, n, 6, q)

                    ExcelKeywords.setValueToCellByIndex(sheet01, n, 7, deitel)

                    n = (n + 1)

                    ExcelKeywords.saveWorkbook(excelFilePath, workbook01)
                } else {
                    if (WebUI.verifyTextNotPresent('У виджета нет данных', false) == false) {
                        ExcelKeywords.setValueToCellByIndex(sheet01, n, 3, dashboardName)

                        ExcelKeywords.setValueToCellByIndex(sheet01, n, 4, dZO)

                        ExcelKeywords.setValueToCellByIndex(sheet01, n, 5, y)

                        ExcelKeywords.setValueToCellByIndex(sheet01, n, 6, q)

                        ExcelKeywords.setValueToCellByIndex(sheet01, n, 7, deitel)

                        n = (n + 1)

                        ExcelKeywords.saveWorkbook(excelFilePath, workbook01)
                    } else {
                        if (WebUI.verifyTextNotPresent('Некорректные фильтры', false) == false) {
                            ExcelKeywords.setValueToCellByIndex(sheet01, n, 3, dashboardName)

                            ExcelKeywords.setValueToCellByIndex(sheet01, n, 4, dZO)

                            ExcelKeywords.setValueToCellByIndex(sheet01, n, 5, y)

                            ExcelKeywords.setValueToCellByIndex(sheet01, n, 6, q)

                            ExcelKeywords.setValueToCellByIndex(sheet01, n, 7, deitel)

                            n = (n + 1)

                            ExcelKeywords.saveWorkbook(excelFilePath, workbook01)
                        }
                    }
                }
            }
        }
    }
}

