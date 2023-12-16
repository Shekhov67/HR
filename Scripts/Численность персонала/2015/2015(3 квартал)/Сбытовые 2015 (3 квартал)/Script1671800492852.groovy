import static com.kms.katalon.core.checkpoint.CheckpointFactory.findCheckpoint
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
import com.kms.katalon.keyword.excel.ExcelKeywords as ExcelKeywords

WebUI.openBrowser(findTestData('New Test Data').getValue(3, 4))

WebUI.setText(findTestObject('Page_/input__username'), findTestData('New Test Data').getValue(1, 1))

WebUI.setText(findTestObject('Page_/input__password'), findTestData('New Test Data').getValue(2, 1))

WebUI.click(findTestObject('Page_/button_'))

WebUI.click(findTestObject('Численность персонала/Фильтр Дата'))

WebUI.click(findTestObject('Численность персонала/Снять выделение в фильтре Дата'))

WebUI.click(findTestObject('Численность персонала/div_2015'))

WebUI.click(findTestObject('Численность персонала/2015-3'))

WebUI.click(findTestObject('Численность персонала/Применить в фильтре Дата'))

WebUI.click(findTestObject('Численность персонала/Фильтр ДЗО'))

WebUI.click(findTestObject('Численность персонала/Снять выделение в фильтре ДЗО'))

WebUI.click(findTestObject('Численность персонала/Сбытовые'))

WebUI.click(findTestObject('Численность персонала/Применить в фильтре ДЗО'))

def write = WriteToExcel()

//def percentCheck = Percents()
WebUI.click(findTestObject('Численность персонала/Фильтр ДЗО'))

WebUI.click(findTestObject('Численность персонала/Снять выделение в фильтре ДЗО'))

WebUI.click(findTestObject('Численность персонала/Развернуть Сбытовые'))

WebUI.click(findTestObject('Численность персонала/Калмэнергосбыт верхний уровень'))

WebUI.click(findTestObject('Численность персонала/Применить в фильтре ДЗО'))

write = WriteToExcel()

//percentCheck = Percents()
WebUI.click(findTestObject('Численность персонала/Фильтр ДЗО'))

WebUI.click(findTestObject('Численность персонала/Снять выделение в фильтре ДЗО'))

WebUI.click(findTestObject('Общие элементы/Развернуть Калмэнергосбыт'))

WebUI.click(findTestObject('Общие элементы/Калмэнергосбыт нижний уровень'))

WebUI.click(findTestObject('Численность персонала/Применить в фильтре ДЗО'))

write = WriteToExcel()

//percentCheck = Percents()
WebUI.click(findTestObject('Численность персонала/Фильтр ДЗО'))

WebUI.click(findTestObject('Численность персонала/Снять выделение в фильтре ДЗО'))

WebUI.click(findTestObject('Численность персонала/АО Карачаево-Черкесскэнерго верхний уровень'))

WebUI.click(findTestObject('Численность персонала/Применить в фильтре ДЗО'))

write = WriteToExcel()

//percentCheck = Percents()
WebUI.click(findTestObject('Численность персонала/Фильтр ДЗО'))

WebUI.click(findTestObject('Численность персонала/Снять выделение в фильтре ДЗО'))

WebUI.click(findTestObject('Численность персонала/Развернуть АО Карачаево-Черкесскэнерго'))

WebUI.click(findTestObject('Общие элементы/АО Карачаево-Черкесскэнерго нижний уровень'))

WebUI.click(findTestObject('Численность персонала/Применить в фильтре ДЗО'))

write = WriteToExcel()

//percentCheck = Percents()
WebUI.click(findTestObject('Численность персонала/Фильтр ДЗО'))

WebUI.click(findTestObject('Численность персонала/Снять выделение в фильтре ДЗО'))

WebUI.scrollToElement(findTestObject('Численность персонала/АО Карачаево-Черкесскэнерго верхний уровень'), 30)

WebUI.scrollToElement(findTestObject('Численность персонала/a_DZO'), 30)

WebUI.click(findTestObject('Численность персонала/АО Тываэнергосбыт верхний уровень'))

WebUI.click(findTestObject('Численность персонала/Применить в фильтре ДЗО'))

write = WriteToExcel()

//percentCheck = Percents()
WebUI.click(findTestObject('Численность персонала/Фильтр ДЗО'))

WebUI.click(findTestObject('Численность персонала/Снять выделение в фильтре ДЗО'))

WebUI.scrollToElement(findTestObject('Численность персонала/АО Тываэнергосбыт верхний уровень'), 30)

WebUI.scrollToElement(findTestObject('Численность персонала/a_DZO'), 30)

WebUI.click(findTestObject('Общие элементы/Развернуть АО Тываэнергосбыт'))

WebUI.click(findTestObject('Общие элементы/АО Тываэнергосбыт нижний уровень'))

WebUI.click(findTestObject('Численность персонала/Применить в фильтре ДЗО'))

write = WriteToExcel()

//percentCheck = Percents()
WebUI.click(findTestObject('Численность персонала/Фильтр ДЗО'))

WebUI.click(findTestObject('Численность персонала/Снять выделение в фильтре ДЗО'))

WebUI.scrollToElement(findTestObject('Численность персонала/АО Тываэнергосбыт верхний уровень'), 30)

WebUI.scrollToElement(findTestObject('Численность персонала/a_DZO'), 30)

WebUI.click(findTestObject('Численность персонала/АО Севкавказэнерго верхний уровень'))

WebUI.click(findTestObject('Численность персонала/Применить в фильтре ДЗО'))

write = WriteToExcel()

//percentCheck = Percents()
WebUI.click(findTestObject('Численность персонала/Фильтр ДЗО'))

WebUI.click(findTestObject('Численность персонала/Снять выделение в фильтре ДЗО'))

WebUI.scrollToElement(findTestObject('Численность персонала/АО Тываэнергосбыт верхний уровень'), 30)

WebUI.scrollToElement(findTestObject('Численность персонала/a_DZO'), 30)

WebUI.click(findTestObject('Численность персонала/Развернуть АО Севкавказэнерго'))

WebUI.click(findTestObject('Численность персонала/АО Севкавказэнерго нижний уровень'))

WebUI.click(findTestObject('Численность персонала/Применить в фильтре ДЗО'))

write = WriteToExcel()

//percentCheck = Percents()
WebUI.click(findTestObject('Численность персонала/Фильтр ДЗО'))

WebUI.click(findTestObject('Численность персонала/Снять выделение в фильтре ДЗО'))

WebUI.scrollToElement(findTestObject('Численность персонала/АО Севкавказэнерго верхний уровень'), 30)

WebUI.scrollToElement(findTestObject('Численность персонала/a_DZO'), 30)

WebUI.click(findTestObject('Численность персонала/Кабардино-Балкарское АО верхний уровень'))

WebUI.click(findTestObject('Численность персонала/Применить в фильтре ДЗО'))

write = WriteToExcel()

//percentCheck = Percents()
WebUI.click(findTestObject('Численность персонала/Фильтр ДЗО'))

WebUI.click(findTestObject('Численность персонала/Снять выделение в фильтре ДЗО'))

WebUI.scrollToElement(findTestObject('Численность персонала/АО Севкавказэнерго верхний уровень'), 30)

WebUI.scrollToElement(findTestObject('Численность персонала/a_DZO'), 30)

WebUI.click(findTestObject('Общие элементы/Развернуть Кабардино-Балкарское АО'))

WebUI.click(findTestObject('Общие элементы/ПАО Каббалкэнерго'))

WebUI.click(findTestObject('Численность персонала/Применить в фильтре ДЗО'))

write = WriteToExcel()

//percentCheck = Percents()
WebUI.click(findTestObject('Численность персонала/Фильтр ДЗО'))

WebUI.click(findTestObject('Численность персонала/Снять выделение в фильтре ДЗО'))

WebUI.scrollToElement(findTestObject('Численность персонала/Кабардино-Балкарское АО верхний уровень'), 30)

WebUI.scrollToElement(findTestObject('Численность персонала/a_DZO'), 30)

WebUI.click(findTestObject('Численность персонала/ОАО Екатеринбургэнергосбыт верхний уровень'))

WebUI.click(findTestObject('Численность персонала/Применить в фильтре ДЗО'))

write = WriteToExcel()

//percentCheck = Percents()
WebUI.click(findTestObject('Численность персонала/Фильтр ДЗО'))

WebUI.click(findTestObject('Численность персонала/Снять выделение в фильтре ДЗО'))

WebUI.scrollToElement(findTestObject('Численность персонала/Кабардино-Балкарское АО верхний уровень'), 30)

WebUI.scrollToElement(findTestObject('Численность персонала/a_DZO'), 30)

WebUI.click(findTestObject('Численность персонала/Развернуть ОАО Екатеринбургэнергосбыт'))

WebUI.click(findTestObject('Численность персонала/ОАО Екатеринбургэнергосбыт нижний уровень'))

WebUI.click(findTestObject('Численность персонала/Применить в фильтре ДЗО'))

write = WriteToExcel()

//percentCheck = Percents()
WebUI.click(findTestObject('Численность персонала/Фильтр ДЗО'))

WebUI.click(findTestObject('Численность персонала/Снять выделение в фильтре ДЗО'))

WebUI.scrollToElement(findTestObject('Численность персонала/ОАО Екатеринбургэнергосбыт верхний уровень'), 30)

WebUI.scrollToElement(findTestObject('Численность персонала/a_DZO'), 30)

WebUI.click(findTestObject('Численность персонала/Янтарьэнергосбыт верхний уровень'))

WebUI.click(findTestObject('Численность персонала/Применить в фильтре ДЗО'))

write = WriteToExcel()

//percentCheck = Percents()
WebUI.click(findTestObject('Численность персонала/Фильтр ДЗО'))

WebUI.click(findTestObject('Численность персонала/Снять выделение в фильтре ДЗО'))

WebUI.scrollToElement(findTestObject('Численность персонала/ОАО Екатеринбургэнергосбыт верхний уровень'), 30)

WebUI.scrollToElement(findTestObject('Численность персонала/a_DZO'), 30)

WebUI.click(findTestObject('Общие элементы/Развернуть Янтарьэнергосбыт'))

WebUI.click(findTestObject('Численность персонала/Янтарьэнергосбыт нижний уровень'))

WebUI.click(findTestObject('Численность персонала/Применить в фильтре ДЗО'))

write = WriteToExcel()

//percentCheck = Percents()
WebUI.click(findTestObject('Численность персонала/Фильтр ДЗО'))

WebUI.click(findTestObject('Численность персонала/Снять выделение в фильтре ДЗО'))

WebUI.scrollToElement(findTestObject('Численность персонала/Янтарьэнергосбыт верхний уровень'), 30)

WebUI.scrollToElement(findTestObject('Численность персонала/a_DZO'), 30)

WebUI.click(findTestObject('Численность персонала/Псковэнергосбыт верхний уровень'))

WebUI.click(findTestObject('Численность персонала/Применить в фильтре ДЗО'))

write = WriteToExcel()

//percentCheck = Percents()
WebUI.click(findTestObject('Численность персонала/Фильтр ДЗО'))

WebUI.click(findTestObject('Численность персонала/Снять выделение в фильтре ДЗО'))

WebUI.scrollToElement(findTestObject('Численность персонала/Янтарьэнергосбыт верхний уровень'), 30)

WebUI.scrollToElement(findTestObject('Численность персонала/a_DZO'), 30)

WebUI.click(findTestObject('Общие элементы/Развернуть Псковэнергосбыт'))

WebUI.click(findTestObject('Численность персонала/Псковэнергосбыт нижний уровень'))

WebUI.click(findTestObject('Численность персонала/Применить в фильтре ДЗО'))

write = WriteToExcel()

//percentCheck = Percents()
WebUI.click(findTestObject('Численность персонала/Фильтр ДЗО'))

WebUI.click(findTestObject('Численность персонала/Снять выделение в фильтре ДЗО'))

WebUI.scrollToElement(findTestObject('Численность персонала/Янтарьэнергосбыт верхний уровень'), 30)

WebUI.scrollToElement(findTestObject('Численность персонала/a_DZO'), 30)

WebUI.click(findTestObject('Численность персонала/ПАО Дагестанская энергосбытовая компания верхний уровень'))

WebUI.click(findTestObject('Численность персонала/Применить в фильтре ДЗО'))

write = WriteToExcel()

//percentCheck = Percents()
WebUI.click(findTestObject('Численность персонала/Фильтр ДЗО'))

WebUI.click(findTestObject('Численность персонала/Снять выделение в фильтре ДЗО'))

WebUI.scrollToElement(findTestObject('Численность персонала/Янтарьэнергосбыт верхний уровень'), 30)

WebUI.scrollToElement(findTestObject('Численность персонала/a_DZO'), 30)

WebUI.click(findTestObject('Общие элементы/Развернуть ПАО Дагестанская энергосбытовая компания'))

WebUI.click(findTestObject('Общие элементы/ПАО Дагестанская энергосбытовая компания нижний уровень'))

WebUI.click(findTestObject('Численность персонала/Применить в фильтре ДЗО'))

write = WriteToExcel()

WebUI.closeBrowser( //Необходимо указать название листа
    //WebUI.getText(findTestObject('headcount'))
    ) //Необходимо указать корректный путь до файла c Test Data
//Необходимо указать название листа
//Название файла Test Data
//Номер строки = количество уже записанных строк +1
//Название ДЗО берётся из фильтра
//Нужно указать квартал, выбираемый в тесте
	static def WriteToExcel() {
		String excelFilePath = 'C:\\Users\\shehs\\OneDrive\\Desktop\\HR\\Data Files\\HR.xlsx'
	
		String sheetName = 'List1'
	
		def data = findTestData('New Test Data')
	
		int n = data.getRowNumbers() + 1
	
		String dZO = WebUI.getText(findTestObject('Численность персонала/Фильтр ДЗО'))
	
		String date = WebUI.getText(findTestObject('Численность персонала/Фильтр Дата'))
	
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
	
			String dashboardName = WebUI.getText(findTestObject('Численность персонала/dashboardName'))
	
			def workbook01 = ExcelKeywords.getWorkbook(excelFilePath)
	
			def sheet01 = ExcelKeywords.getExcelSheet(workbook01, sheetName)
	
			if (WebUI.verifyTextNotPresent('нет данных', false) == false) {
				ExcelKeywords.setValueToCellByIndex(sheet01, n, 3, dashboardName)
	
				ExcelKeywords.setValueToCellByIndex(sheet01, n, 4, dZO)
	
				ExcelKeywords.setValueToCellByIndex(sheet01, n, 5, y)
	
				ExcelKeywords.setValueToCellByIndex(sheet01, n, 6, q)
	
				n = (n + 1)
	
				ExcelKeywords.saveWorkbook(excelFilePath, workbook01)
			} else {
				if (WebUI.verifyTextNotPresent('Ошибка запроса данных', false) == false) {
					ExcelKeywords.setValueToCellByIndex(sheet01, n, 3, dashboardName)
	
					ExcelKeywords.setValueToCellByIndex(sheet01, n, 4, dZO)
	
					ExcelKeywords.setValueToCellByIndex(sheet01, n, 5, y)
	
					ExcelKeywords.setValueToCellByIndex(sheet01, n, 6, q)
	
					n = (n + 1)
	
					ExcelKeywords.saveWorkbook(excelFilePath, workbook01)
				} else {
					if (WebUI.verifyTextNotPresent('Произошла ошибка при выполнении пользовательского кода', false) == false) {
						ExcelKeywords.setValueToCellByIndex(sheet01, n, 3, dashboardName)
	
						ExcelKeywords.setValueToCellByIndex(sheet01, n, 4, dZO)
	
						ExcelKeywords.setValueToCellByIndex(sheet01, n, 5, y)
	
						ExcelKeywords.setValueToCellByIndex(sheet01, n, 6, q)
	
						n = (n + 1)
	
						ExcelKeywords.saveWorkbook(excelFilePath, workbook01)
					} else {
						if (WebUI.verifyTextNotPresent('У виджета нет данных', false) == false) {
							ExcelKeywords.setValueToCellByIndex(sheet01, n, 3, dashboardName)
	
							ExcelKeywords.setValueToCellByIndex(sheet01, n, 4, dZO)
	
							ExcelKeywords.setValueToCellByIndex(sheet01, n, 5, y)
	
							ExcelKeywords.setValueToCellByIndex(sheet01, n, 6, q)
	
							n = (n + 1)
	
							ExcelKeywords.saveWorkbook(excelFilePath, workbook01)
						}
					}
				}
			}
		} else {
			String q = ''
	
			println(q)
	
			String dashboardName = WebUI.getText(findTestObject('Численность персонала/dashboardName'))
	
			def workbook01 = ExcelKeywords.getWorkbook(excelFilePath)
	
			def sheet01 = ExcelKeywords.getExcelSheet(workbook01, sheetName)
	
			if (WebUI.verifyTextNotPresent('нет данных', false) == false) {
				ExcelKeywords.setValueToCellByIndex(sheet01, n, 3, dashboardName)
	
				ExcelKeywords.setValueToCellByIndex(sheet01, n, 4, dZO)
	
				ExcelKeywords.setValueToCellByIndex(sheet01, n, 5, y)
	
				ExcelKeywords.setValueToCellByIndex(sheet01, n, 6, q)
	
				n = (n + 1)
	
				ExcelKeywords.saveWorkbook(excelFilePath, workbook01)
			} else {
				if (WebUI.verifyTextNotPresent('Ошибка запроса данных', false) == false) {
					ExcelKeywords.setValueToCellByIndex(sheet01, n, 3, dashboardName)
	
					ExcelKeywords.setValueToCellByIndex(sheet01, n, 4, dZO)
	
					ExcelKeywords.setValueToCellByIndex(sheet01, n, 5, y)
	
					ExcelKeywords.setValueToCellByIndex(sheet01, n, 6, q)
	
					n = (n + 1)
	
					ExcelKeywords.saveWorkbook(excelFilePath, workbook01)
				} else {
					if (WebUI.verifyTextNotPresent('Произошла ошибка при выполнении пользовательского кода', false) == false) {
						ExcelKeywords.setValueToCellByIndex(sheet01, n, 3, dashboardName)
	
						ExcelKeywords.setValueToCellByIndex(sheet01, n, 4, dZO)
	
						ExcelKeywords.setValueToCellByIndex(sheet01, n, 5, y)
	
						ExcelKeywords.setValueToCellByIndex(sheet01, n, 6, q)
	
						n = (n + 1)
	
						ExcelKeywords.saveWorkbook(excelFilePath, workbook01)
					} else {
						if (WebUI.verifyTextNotPresent('У виджета нет данных', false) == false) {
							ExcelKeywords.setValueToCellByIndex(sheet01, n, 3, dashboardName)
	
							ExcelKeywords.setValueToCellByIndex(sheet01, n, 4, dZO)
	
							ExcelKeywords.setValueToCellByIndex(sheet01, n, 5, y)
	
							ExcelKeywords.setValueToCellByIndex(sheet01, n, 6, q)
	
							n = (n + 1)
	
							ExcelKeywords.saveWorkbook(excelFilePath, workbook01)
						}
					}
				}
			}
		}
	}
	
	
