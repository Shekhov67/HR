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

WebUI.openBrowser(findTestData('New Test Data').getValue(3, 5))

WebUI.setText(findTestObject('Page_/input__username'), findTestData('New Test Data').getValue(1, 1))

WebUI.setText(findTestObject('Page_/input__password'), findTestData('New Test Data').getValue(2, 1))

WebUI.click(findTestObject('Page_/button_'))

WebUI.click(findTestObject('Расходы на персонал/фильтр Деятельность ДЗО'))

WebUI.click(findTestObject('Расходы на персонал/деятельность КАПИТАЛЬНОЕ СТРОИТЕЛЬСТРВО'))

WebUI.click(findTestObject('Расходы на персонал/применить в фильтре деятельность'))

WebUI.click(findTestObject('Расходы на персонал/Фильтр ДЗО'))

WebUI.click(findTestObject('Расходы на персонал/Снять выделение в фильтре ДЗО'))

WebUI.click(findTestObject('Расходы на персонал/Развернуть Сетевые'))

WebUI.scrollToElement(findTestObject('Общие элементы/Развернуть Россети Северо-Запад'), 30)

WebUI.scrollToElement(findTestObject('Расходы на персонал/a_DZO'), 30)

WebUI.click(findTestObject('Расходы на персонал/Россети Северо-Запад верхний уровень'))

WebUI.click(findTestObject('Расходы на персонал/Применить в фильтре ДЗО'))

def write = WriteToExcel()

//def percentCheck = Percents()
WebUI.click(findTestObject('Расходы на персонал/Фильтр ДЗО'))

WebUI.click(findTestObject('Расходы на персонал/Снять выделение в фильтре ДЗО'))

WebUI.scrollToElement(findTestObject('Общие элементы/Развернуть Россети Северо-Запад'), 30)

WebUI.scrollToElement(findTestObject('Расходы на персонал/a_DZO'), 30)

WebUI.click(findTestObject('Общие элементы/Развернуть Россети Северо-Запад'))

WebUI.click(findTestObject('Общие элементы/Архангельский филиал'))

WebUI.click(findTestObject('Расходы на персонал/Применить в фильтре ДЗО'))

write = WriteToExcel()

//percentCheck = Percents()
WebUI.click(findTestObject('Расходы на персонал/Фильтр ДЗО'))

WebUI.click(findTestObject('Расходы на персонал/Снять выделение в фильтре ДЗО'))

WebUI.scrollToElement(findTestObject('Общие элементы/Развернуть Россети Северо-Запад'), 30)

WebUI.scrollToElement(findTestObject('Расходы на персонал/a_DZO'), 30)

WebUI.click(findTestObject('Общие элементы/Вологодский филиал'))

WebUI.click(findTestObject('Расходы на персонал/Применить в фильтре ДЗО'))

write = WriteToExcel()

//percentCheck = Percents()
WebUI.click(findTestObject('Расходы на персонал/Фильтр ДЗО'))

WebUI.click(findTestObject('Расходы на персонал/Снять выделение в фильтре ДЗО'))

WebUI.scrollToElement(findTestObject('Общие элементы/Развернуть Россети Северо-Запад'), 30)

WebUI.scrollToElement(findTestObject('Расходы на персонал/a_DZO'), 30)

WebUI.click(findTestObject('Расходы на персонал/Исполнительный аппарат Северо-Запад'))

WebUI.click(findTestObject('Расходы на персонал/Применить в фильтре ДЗО'))

write = WriteToExcel()

//percentCheck = Percents()
WebUI.click(findTestObject('Расходы на персонал/Фильтр ДЗО'))

WebUI.click(findTestObject('Расходы на персонал/Снять выделение в фильтре ДЗО'))

WebUI.scrollToElement(findTestObject('Общие элементы/Развернуть Россети Северо-Запад'), 30)

WebUI.scrollToElement(findTestObject('Расходы на персонал/a_DZO'), 30)

WebUI.click(findTestObject('Общие элементы/Карельский филиал'))

WebUI.click(findTestObject('Расходы на персонал/Применить в фильтре ДЗО'))

write = WriteToExcel()

//percentCheck = Percents()
WebUI.click(findTestObject('Расходы на персонал/Фильтр ДЗО'))

WebUI.click(findTestObject('Расходы на персонал/Снять выделение в фильтре ДЗО'))

WebUI.scrollToElement(findTestObject('Общие элементы/Развернуть Россети Северо-Запад'), 30)

WebUI.scrollToElement(findTestObject('Расходы на персонал/a_DZO'), 30)

WebUI.click(findTestObject('Общие элементы/Колэнерго'))

WebUI.click(findTestObject('Расходы на персонал/Применить в фильтре ДЗО'))

write = WriteToExcel()

//percentCheck = Percents()
WebUI.click(findTestObject('Расходы на персонал/Фильтр ДЗО'))

WebUI.click(findTestObject('Расходы на персонал/Снять выделение в фильтре ДЗО'))

WebUI.scrollToElement(findTestObject('Общие элементы/Колэнерго'), 30)

WebUI.scrollToElement(findTestObject('Расходы на персонал/a_DZO'), 30)

WebUI.click(findTestObject('Общие элементы/Новгородский филиал'))

WebUI.click(findTestObject('Расходы на персонал/Применить в фильтре ДЗО'))

write = WriteToExcel()

//percentCheck = Percents()
WebUI.click(findTestObject('Расходы на персонал/Фильтр ДЗО'))

WebUI.click(findTestObject('Расходы на персонал/Снять выделение в фильтре ДЗО'))

WebUI.scrollToElement(findTestObject('Общие элементы/Колэнерго'), 30)

WebUI.scrollToElement(findTestObject('Расходы на персонал/a_DZO'), 30)

WebUI.click(findTestObject('Общие элементы/Псковский филиал'))

WebUI.click(findTestObject('Расходы на персонал/Применить в фильтре ДЗО'))

write = WriteToExcel()

//percentCheck = Percents()
WebUI.click(findTestObject('Расходы на персонал/Фильтр ДЗО'))

WebUI.click(findTestObject('Расходы на персонал/Снять выделение в фильтре ДЗО'))

WebUI.scrollToElement(findTestObject('Общие элементы/Колэнерго'), 30)

WebUI.scrollToElement(findTestObject('Расходы на персонал/a_DZO'), 30)

WebUI.click(findTestObject('Общие элементы/филиал в Республике Коми'))

WebUI.click(findTestObject('Расходы на персонал/Применить в фильтре ДЗО'))

write = WriteToExcel()

//percentCheck = Percents()
WebUI.closeBrowser()

static def WriteToExcel() {
	String dZO = WebUI.getText(findTestObject('Расходы на персонал/Фильтр ДЗО'))
		
	println(dZO)
	
	
	
	String[] mass = dZO.split("-")
	
	println(mass.length)
	
	if(mass.length==1) {
		String mass1 = mass[0]
		String mass2 = 'Все ДЗО'
		String mass3 = 'Все филиалы'
		
		String excelFilePath = 'C:\\Users\\shehs\\OneDrive\\Desktop\\HR\\Data Files\\HR.xlsx'
		
			String sheetName = 'List1'
		
			def data = findTestData('New Test Data')
		
			int n = data.getRowNumbers() + 1
			
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
			
						ExcelKeywords.setValueToCellByIndex(sheet01, n, 4, mass1)
						
						ExcelKeywords.setValueToCellByIndex(sheet01, n, 5, mass2)
						
						ExcelKeywords.setValueToCellByIndex(sheet01, n, 6, mass3)
						
						ExcelKeywords.setValueToCellByIndex(sheet01, n, 7, y)
			
						ExcelKeywords.setValueToCellByIndex(sheet01, n, 8, q)
			
						ExcelKeywords.setValueToCellByIndex(sheet01, n, 9, deitel)
			
						n = (n + 1)
			
						ExcelKeywords.saveWorkbook(excelFilePath, workbook01)
						println(mass1 + mass2 + mass3)
					} else {
						if (WebUI.verifyTextNotPresent('Ошибка запроса данных', false) == false) {
							ExcelKeywords.setValueToCellByIndex(sheet01, n, 3, dashboardName)
			
							ExcelKeywords.setValueToCellByIndex(sheet01, n, 4, mass1)
						
							ExcelKeywords.setValueToCellByIndex(sheet01, n, 5, mass2)
						
							ExcelKeywords.setValueToCellByIndex(sheet01, n, 6, mass3)
						
							ExcelKeywords.setValueToCellByIndex(sheet01, n, 7, y)
			
							ExcelKeywords.setValueToCellByIndex(sheet01, n, 8, q)
			
							ExcelKeywords.setValueToCellByIndex(sheet01, n, 9, deitel)
			
							n = (n + 1)
			
							ExcelKeywords.saveWorkbook(excelFilePath, workbook01)
							println(mass1 + mass2 + mass3)
						} else {
							if (WebUI.verifyTextNotPresent('Произошла ошибка при выполнении пользовательского кода', false) == false) {
								ExcelKeywords.setValueToCellByIndex(sheet01, n, 3, dashboardName)
			
								ExcelKeywords.setValueToCellByIndex(sheet01, n, 4, mass1)
						
								ExcelKeywords.setValueToCellByIndex(sheet01, n, 5, mass2)
						
								ExcelKeywords.setValueToCellByIndex(sheet01, n, 6, mass3)
						
								ExcelKeywords.setValueToCellByIndex(sheet01, n, 7, y)
								
								ExcelKeywords.setValueToCellByIndex(sheet01, n, 8, q)
			
								ExcelKeywords.setValueToCellByIndex(sheet01, n, 9, deitel)
			
								n = (n + 1)
			
								ExcelKeywords.saveWorkbook(excelFilePath, workbook01)
								println(mass1 + mass2 + mass3)
							} else {
								if (WebUI.verifyTextNotPresent('У виджета нет данных', false) == false) {
									ExcelKeywords.setValueToCellByIndex(sheet01, n, 3, dashboardName)
			
									ExcelKeywords.setValueToCellByIndex(sheet01, n, 4, mass1)
						
									ExcelKeywords.setValueToCellByIndex(sheet01, n, 5, mass2)
							
									ExcelKeywords.setValueToCellByIndex(sheet01, n, 6, mass3)
						
									ExcelKeywords.setValueToCellByIndex(sheet01, n, 7, y)
							
									ExcelKeywords.setValueToCellByIndex(sheet01, n, 8, q)
			
									ExcelKeywords.setValueToCellByIndex(sheet01, n, 9, deitel)
			
									n = (n + 1)
			
									ExcelKeywords.saveWorkbook(excelFilePath, workbook01)
									println(mass1 + mass2 + mass3)
								} else {
									if (WebUI.verifyTextNotPresent('Некорректные фильтры', false) == false) {
										ExcelKeywords.setValueToCellByIndex(sheet01, n, 3, dashboardName)
			
										ExcelKeywords.setValueToCellByIndex(sheet01, n, 4, mass1)
						
										ExcelKeywords.setValueToCellByIndex(sheet01, n, 5, mass2)
						
										ExcelKeywords.setValueToCellByIndex(sheet01, n, 6, mass3)
						
										ExcelKeywords.setValueToCellByIndex(sheet01, n, 7, y)
			
										ExcelKeywords.setValueToCellByIndex(sheet01, n, 8, q)
			
										ExcelKeywords.setValueToCellByIndex(sheet01, n, 9, deitel)
										n = (n + 1)
			
										ExcelKeywords.saveWorkbook(excelFilePath, workbook01)
										println(mass1 + mass2 + mass3)
									}
								}
							}
						}
					}
				} else {
					String q = '1-4 кварталы'
			
					println(q)
			
					String dashboardName = 'Расходы на персонал'
			
					def workbook01 = ExcelKeywords.getWorkbook(excelFilePath)
			
					def sheet01 = ExcelKeywords.getExcelSheet(workbook01, sheetName)
			
					if (WebUI.verifyTextNotPresent('нет данных', false) == false) {
							ExcelKeywords.setValueToCellByIndex(sheet01, n, 3, dashboardName)
						
							ExcelKeywords.setValueToCellByIndex(sheet01, n, 4, mass1)
							
							ExcelKeywords.setValueToCellByIndex(sheet01, n, 5, mass2)
						
							ExcelKeywords.setValueToCellByIndex(sheet01, n, 6, mass3)
						
							ExcelKeywords.setValueToCellByIndex(sheet01, n, 7, y)
			
							ExcelKeywords.setValueToCellByIndex(sheet01, n, 8, q)
			
							ExcelKeywords.setValueToCellByIndex(sheet01, n, 9, deitel)
						n = (n + 1)
			
						ExcelKeywords.saveWorkbook(excelFilePath, workbook01)
						println(mass1 +' good '+ mass2 )
					} else {
						if (WebUI.verifyTextNotPresent('Ошибка запроса данных', false) == false) {
							ExcelKeywords.setValueToCellByIndex(sheet01, n, 3, dashboardName)
			
							ExcelKeywords.setValueToCellByIndex(sheet01, n, 4, mass1)
							
							ExcelKeywords.setValueToCellByIndex(sheet01, n, 5, mass2)
						
							ExcelKeywords.setValueToCellByIndex(sheet01, n, 6, mass3)
						
							ExcelKeywords.setValueToCellByIndex(sheet01, n, 7, y)
			
							ExcelKeywords.setValueToCellByIndex(sheet01, n, 8, q)
			
							ExcelKeywords.setValueToCellByIndex(sheet01, n, 9, deitel)
			
							n = (n + 1)
			
							ExcelKeywords.saveWorkbook(excelFilePath, workbook01)
							println(mass1 +' good '+ mass2 )
						} else {
							if (WebUI.verifyTextNotPresent('Произошла ошибка при выполнении пользовательского кода', false) == false) {
								ExcelKeywords.setValueToCellByIndex(sheet01, n, 3, dashboardName)
			
								ExcelKeywords.setValueToCellByIndex(sheet01, n, 4, mass1)
						
								ExcelKeywords.setValueToCellByIndex(sheet01, n, 5, mass2)
						
								ExcelKeywords.setValueToCellByIndex(sheet01, n, 6, mass3)
						
								ExcelKeywords.setValueToCellByIndex(sheet01, n, 7, y)
			
								ExcelKeywords.setValueToCellByIndex(sheet01, n, 8, q)
			
								ExcelKeywords.setValueToCellByIndex(sheet01, n, 9, deitel)
			
								n = (n + 1)
			
								ExcelKeywords.saveWorkbook(excelFilePath, workbook01)
								println(mass1 +' good '+ mass2 )
							} else {
								if (WebUI.verifyTextNotPresent('У виджета нет данных', false) == false) {
									ExcelKeywords.setValueToCellByIndex(sheet01, n, 3, dashboardName)
			
									ExcelKeywords.setValueToCellByIndex(sheet01, n, 4, mass1)
						
									ExcelKeywords.setValueToCellByIndex(sheet01, n, 5, mass2)
						
									ExcelKeywords.setValueToCellByIndex(sheet01, n, 6, mass3)
						
									ExcelKeywords.setValueToCellByIndex(sheet01, n, 7, y)
			
									ExcelKeywords.setValueToCellByIndex(sheet01, n, 8, q)
			
									ExcelKeywords.setValueToCellByIndex(sheet01, n, 9, deitel)
									n = (n + 1)
			
									ExcelKeywords.saveWorkbook(excelFilePath, workbook01)
									println(mass1 +' good '+ mass2 )
								} else {
									if (WebUI.verifyTextNotPresent('Некорректные фильтры', false) == false) {
										ExcelKeywords.setValueToCellByIndex(sheet01, n, 3, dashboardName)
										
										ExcelKeywords.setValueToCellByIndex(sheet01, n, 4, mass1)
													
										ExcelKeywords.setValueToCellByIndex(sheet01, n, 5, mass2)
													
										ExcelKeywords.setValueToCellByIndex(sheet01, n, 6, mass3)
												
										ExcelKeywords.setValueToCellByIndex(sheet01, n, 7, y)
										
										ExcelKeywords.setValueToCellByIndex(sheet01, n, 8, q)
										
										ExcelKeywords.setValueToCellByIndex(sheet01, n, 9, deitel)
			
										n = (n + 1)
			
										ExcelKeywords.saveWorkbook(excelFilePath, workbook01)
										println(mass1 +' good '+ mass2 )
									}
								}
							}
						}
					}
				}
				
		//println(mass1)
	}else if(mass.length==2){
		String mass1 = mass[0]
		String mass2 = mass[1]
		String mass3 = 'Все филиалы'
		
		String excelFilePath = 'C:\\Users\\shehs\\OneDrive\\Desktop\\HR\\Data Files\\HR.xlsx'
		
			String sheetName = 'List1'
		
			def data = findTestData('New Test Data')
		
			int n = data.getRowNumbers() + 1
			
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
			
						ExcelKeywords.setValueToCellByIndex(sheet01, n, 4, mass1)
						
						ExcelKeywords.setValueToCellByIndex(sheet01, n, 5, mass2)
						
						ExcelKeywords.setValueToCellByIndex(sheet01, n, 6, mass3)
						
						ExcelKeywords.setValueToCellByIndex(sheet01, n, 7, y)
			
						ExcelKeywords.setValueToCellByIndex(sheet01, n, 8, q)
			
						ExcelKeywords.setValueToCellByIndex(sheet01, n, 9, deitel)
			
						n = (n + 1)
			
						ExcelKeywords.saveWorkbook(excelFilePath, workbook01)
						println(mass1 + mass2 + mass3)
					} else {
						if (WebUI.verifyTextNotPresent('Ошибка запроса данных', false) == false) {
							ExcelKeywords.setValueToCellByIndex(sheet01, n, 3, dashboardName)
			
							ExcelKeywords.setValueToCellByIndex(sheet01, n, 4, mass1)
						
							ExcelKeywords.setValueToCellByIndex(sheet01, n, 5, mass2)
						
							ExcelKeywords.setValueToCellByIndex(sheet01, n, 6, mass3)
						
							ExcelKeywords.setValueToCellByIndex(sheet01, n, 7, y)
			
							ExcelKeywords.setValueToCellByIndex(sheet01, n, 8, q)
			
							ExcelKeywords.setValueToCellByIndex(sheet01, n, 9, deitel)
			
							n = (n + 1)
			
							ExcelKeywords.saveWorkbook(excelFilePath, workbook01)
							println(mass1 + mass2 + mass3)
						} else {
							if (WebUI.verifyTextNotPresent('Произошла ошибка при выполнении пользовательского кода', false) == false) {
								ExcelKeywords.setValueToCellByIndex(sheet01, n, 3, dashboardName)
			
								ExcelKeywords.setValueToCellByIndex(sheet01, n, 4, mass1)
						
								ExcelKeywords.setValueToCellByIndex(sheet01, n, 5, mass2)
						
								ExcelKeywords.setValueToCellByIndex(sheet01, n, 6, mass3)
						
								ExcelKeywords.setValueToCellByIndex(sheet01, n, 7, y)
								
								ExcelKeywords.setValueToCellByIndex(sheet01, n, 8, q)
			
								ExcelKeywords.setValueToCellByIndex(sheet01, n, 9, deitel)
			
								n = (n + 1)
			
								ExcelKeywords.saveWorkbook(excelFilePath, workbook01)
								println(mass1 + mass2 + mass3)
							} else {
								if (WebUI.verifyTextNotPresent('У виджета нет данных', false) == false) {
									ExcelKeywords.setValueToCellByIndex(sheet01, n, 3, dashboardName)
			
									ExcelKeywords.setValueToCellByIndex(sheet01, n, 4, mass1)
						
									ExcelKeywords.setValueToCellByIndex(sheet01, n, 5, mass2)
							
									ExcelKeywords.setValueToCellByIndex(sheet01, n, 6, mass3)
						
									ExcelKeywords.setValueToCellByIndex(sheet01, n, 7, y)
							
									ExcelKeywords.setValueToCellByIndex(sheet01, n, 8, q)
			
									ExcelKeywords.setValueToCellByIndex(sheet01, n, 9, deitel)
			
									n = (n + 1)
			
									ExcelKeywords.saveWorkbook(excelFilePath, workbook01)
									println(mass1 + mass2 + mass3)
								} else {
									if (WebUI.verifyTextNotPresent('Некорректные фильтры', false) == false) {
										ExcelKeywords.setValueToCellByIndex(sheet01, n, 3, dashboardName)
			
										ExcelKeywords.setValueToCellByIndex(sheet01, n, 4, mass1)
						
										ExcelKeywords.setValueToCellByIndex(sheet01, n, 5, mass2)
						
										ExcelKeywords.setValueToCellByIndex(sheet01, n, 6, mass3)
						
										ExcelKeywords.setValueToCellByIndex(sheet01, n, 7, y)
			
										ExcelKeywords.setValueToCellByIndex(sheet01, n, 8, q)
			
										ExcelKeywords.setValueToCellByIndex(sheet01, n, 9, deitel)
										n = (n + 1)
			
										ExcelKeywords.saveWorkbook(excelFilePath, workbook01)
										println(mass1 + mass2 + mass3)
									}
								}
							}
						}
					}
				} else {
					String q = '1-4 кварталы'
			
					println(q)
			
					String dashboardName = 'Расходы на персонал'
			
					def workbook01 = ExcelKeywords.getWorkbook(excelFilePath)
			
					def sheet01 = ExcelKeywords.getExcelSheet(workbook01, sheetName)
			
					if (WebUI.verifyTextNotPresent('нет данных', false) == false) {
							ExcelKeywords.setValueToCellByIndex(sheet01, n, 3, dashboardName)
						
							ExcelKeywords.setValueToCellByIndex(sheet01, n, 4, mass1)
							
							ExcelKeywords.setValueToCellByIndex(sheet01, n, 5, mass2)
						
							ExcelKeywords.setValueToCellByIndex(sheet01, n, 6, mass3)
						
							ExcelKeywords.setValueToCellByIndex(sheet01, n, 7, y)
			
							ExcelKeywords.setValueToCellByIndex(sheet01, n, 8, q)
			
							ExcelKeywords.setValueToCellByIndex(sheet01, n, 9, deitel)
						n = (n + 1)
			
						ExcelKeywords.saveWorkbook(excelFilePath, workbook01)
						println(mass1 +' good '+ mass2 )
					} else {
						if (WebUI.verifyTextNotPresent('Ошибка запроса данных', false) == false) {
							ExcelKeywords.setValueToCellByIndex(sheet01, n, 3, dashboardName)
			
							ExcelKeywords.setValueToCellByIndex(sheet01, n, 4, mass1)
							
							ExcelKeywords.setValueToCellByIndex(sheet01, n, 5, mass2)
						
							ExcelKeywords.setValueToCellByIndex(sheet01, n, 6, mass3)
						
							ExcelKeywords.setValueToCellByIndex(sheet01, n, 7, y)
			
							ExcelKeywords.setValueToCellByIndex(sheet01, n, 8, q)
			
							ExcelKeywords.setValueToCellByIndex(sheet01, n, 9, deitel)
			
							n = (n + 1)
			
							ExcelKeywords.saveWorkbook(excelFilePath, workbook01)
							println(mass1 +' good '+ mass2 )
						} else {
							if (WebUI.verifyTextNotPresent('Произошла ошибка при выполнении пользовательского кода', false) == false) {
								ExcelKeywords.setValueToCellByIndex(sheet01, n, 3, dashboardName)
			
								ExcelKeywords.setValueToCellByIndex(sheet01, n, 4, mass1)
						
								ExcelKeywords.setValueToCellByIndex(sheet01, n, 5, mass2)
						
								ExcelKeywords.setValueToCellByIndex(sheet01, n, 6, mass3)
						
								ExcelKeywords.setValueToCellByIndex(sheet01, n, 7, y)
			
								ExcelKeywords.setValueToCellByIndex(sheet01, n, 8, q)
			
								ExcelKeywords.setValueToCellByIndex(sheet01, n, 9, deitel)
			
								n = (n + 1)
			
								ExcelKeywords.saveWorkbook(excelFilePath, workbook01)
								println(mass1 +' good '+ mass2 )
							} else {
								if (WebUI.verifyTextNotPresent('У виджета нет данных', false) == false) {
									ExcelKeywords.setValueToCellByIndex(sheet01, n, 3, dashboardName)
			
									ExcelKeywords.setValueToCellByIndex(sheet01, n, 4, mass1)
						
									ExcelKeywords.setValueToCellByIndex(sheet01, n, 5, mass2)
						
									ExcelKeywords.setValueToCellByIndex(sheet01, n, 6, mass3)
						
									ExcelKeywords.setValueToCellByIndex(sheet01, n, 7, y)
			
									ExcelKeywords.setValueToCellByIndex(sheet01, n, 8, q)
			
									ExcelKeywords.setValueToCellByIndex(sheet01, n, 9, deitel)
									n = (n + 1)
			
									ExcelKeywords.saveWorkbook(excelFilePath, workbook01)
									println(mass1 +' good '+ mass2 )
								} else {
									if (WebUI.verifyTextNotPresent('Некорректные фильтры', false) == false) {
										ExcelKeywords.setValueToCellByIndex(sheet01, n, 3, dashboardName)
										
										ExcelKeywords.setValueToCellByIndex(sheet01, n, 4, mass1)
													
										ExcelKeywords.setValueToCellByIndex(sheet01, n, 5, mass2)
													
										ExcelKeywords.setValueToCellByIndex(sheet01, n, 6, mass3)
												
										ExcelKeywords.setValueToCellByIndex(sheet01, n, 7, y)
										
										ExcelKeywords.setValueToCellByIndex(sheet01, n, 8, q)
										
										ExcelKeywords.setValueToCellByIndex(sheet01, n, 9, deitel)
			
										n = (n + 1)
			
										ExcelKeywords.saveWorkbook(excelFilePath, workbook01)
										println(mass1 +' good '+ mass2 )
									}
								}
							}
						}
					}
				}
				println(mass1 +' good '+ mass2 )
		
		//println(mass1)
		//println(mass2)
	}else if(mass.length==3){
		String mass1 = mass[0]
		String mass2 = mass[1]
		String mass3 = mass[2]
		
		String excelFilePath = 'C:\\Users\\shehs\\OneDrive\\Desktop\\HR\\Data Files\\HR.xlsx'
		
			String sheetName = 'List1'
		
			def data = findTestData('New Test Data')
		
			int n = data.getRowNumbers() + 1
			
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
			
						ExcelKeywords.setValueToCellByIndex(sheet01, n, 4, mass1)
						
						ExcelKeywords.setValueToCellByIndex(sheet01, n, 5, mass2)
						
						ExcelKeywords.setValueToCellByIndex(sheet01, n, 6, mass3)
						
						ExcelKeywords.setValueToCellByIndex(sheet01, n, 7, y)
			
						ExcelKeywords.setValueToCellByIndex(sheet01, n, 8, q)
			
						ExcelKeywords.setValueToCellByIndex(sheet01, n, 9, deitel)
			
						n = (n + 1)
			
						ExcelKeywords.saveWorkbook(excelFilePath, workbook01)
						println(mass1 + mass2 + mass3)
					} else {
						if (WebUI.verifyTextNotPresent('Ошибка запроса данных', false) == false) {
							ExcelKeywords.setValueToCellByIndex(sheet01, n, 3, dashboardName)
			
							ExcelKeywords.setValueToCellByIndex(sheet01, n, 4, mass1)
						
							ExcelKeywords.setValueToCellByIndex(sheet01, n, 5, mass2)
						
							ExcelKeywords.setValueToCellByIndex(sheet01, n, 6, mass3)
						
							ExcelKeywords.setValueToCellByIndex(sheet01, n, 7, y)
			
							ExcelKeywords.setValueToCellByIndex(sheet01, n, 8, q)
			
							ExcelKeywords.setValueToCellByIndex(sheet01, n, 9, deitel)
			
							n = (n + 1)
			
							ExcelKeywords.saveWorkbook(excelFilePath, workbook01)
							println(mass1 + mass2 + mass3)
						} else {
							if (WebUI.verifyTextNotPresent('Произошла ошибка при выполнении пользовательского кода', false) == false) {
								ExcelKeywords.setValueToCellByIndex(sheet01, n, 3, dashboardName)
			
								ExcelKeywords.setValueToCellByIndex(sheet01, n, 4, mass1)
						
								ExcelKeywords.setValueToCellByIndex(sheet01, n, 5, mass2)
						
								ExcelKeywords.setValueToCellByIndex(sheet01, n, 6, mass3)
						
								ExcelKeywords.setValueToCellByIndex(sheet01, n, 7, y)
								
								ExcelKeywords.setValueToCellByIndex(sheet01, n, 8, q)
			
								ExcelKeywords.setValueToCellByIndex(sheet01, n, 9, deitel)
			
								n = (n + 1)
			
								ExcelKeywords.saveWorkbook(excelFilePath, workbook01)
								println(mass1 + mass2 + mass3)
							} else {
								if (WebUI.verifyTextNotPresent('У виджета нет данных', false) == false) {
								ExcelKeywords.setValueToCellByIndex(sheet01, n, 3, dashboardName)
			
								ExcelKeywords.setValueToCellByIndex(sheet01, n, 4, mass1)
						
								ExcelKeywords.setValueToCellByIndex(sheet01, n, 5, mass2)
							
								ExcelKeywords.setValueToCellByIndex(sheet01, n, 6, mass3)
						
								ExcelKeywords.setValueToCellByIndex(sheet01, n, 7, y)
							
								ExcelKeywords.setValueToCellByIndex(sheet01, n, 8, q)
			
								ExcelKeywords.setValueToCellByIndex(sheet01, n, 9, deitel)
			
								n = (n + 1)
			
								ExcelKeywords.saveWorkbook(excelFilePath, workbook01)
								println(mass1 + mass2 + mass3)
								} else {
									if (WebUI.verifyTextNotPresent('Некорректные фильтры', false) == false) {
									ExcelKeywords.setValueToCellByIndex(sheet01, n, 3, dashboardName)
			
									ExcelKeywords.setValueToCellByIndex(sheet01, n, 4, mass1)
						
									ExcelKeywords.setValueToCellByIndex(sheet01, n, 5, mass2)
						
									ExcelKeywords.setValueToCellByIndex(sheet01, n, 6, mass3)
						
									ExcelKeywords.setValueToCellByIndex(sheet01, n, 7, y)
			
									ExcelKeywords.setValueToCellByIndex(sheet01, n, 8, q)
			
									ExcelKeywords.setValueToCellByIndex(sheet01, n, 9, deitel)
									n = (n + 1)
		
									ExcelKeywords.saveWorkbook(excelFilePath, workbook01)
									println(mass1 + mass2 + mass3)
									}
								}
							}
						}
					}
				} else {
					String q = '1-4 кварталы'
			
					println(q)
			
					String dashboardName = 'Расходы на персонал'
			
					def workbook01 = ExcelKeywords.getWorkbook(excelFilePath)
			
					def sheet01 = ExcelKeywords.getExcelSheet(workbook01, sheetName)
			
					if (WebUI.verifyTextNotPresent('нет данных', false) == false) {
						ExcelKeywords.setValueToCellByIndex(sheet01, n, 3, dashboardName)
						
						ExcelKeywords.setValueToCellByIndex(sheet01, n, 4, mass1)
							
						ExcelKeywords.setValueToCellByIndex(sheet01, n, 5, mass2)
						
						ExcelKeywords.setValueToCellByIndex(sheet01, n, 6, mass3)
					
						ExcelKeywords.setValueToCellByIndex(sheet01, n, 7, y)
			
						ExcelKeywords.setValueToCellByIndex(sheet01, n, 8, q)
			
						ExcelKeywords.setValueToCellByIndex(sheet01, n, 9, deitel)
						n = (n + 1)
			
						ExcelKeywords.saveWorkbook(excelFilePath, workbook01)
						println(mass1 +' good '+ mass2 )
					} else {
						if (WebUI.verifyTextNotPresent('Ошибка запроса данных', false) == false) {
						ExcelKeywords.setValueToCellByIndex(sheet01, n, 3, dashboardName)
			
						ExcelKeywords.setValueToCellByIndex(sheet01, n, 4, mass1)
							
						ExcelKeywords.setValueToCellByIndex(sheet01, n, 5, mass2)
					
						ExcelKeywords.setValueToCellByIndex(sheet01, n, 6, mass3)
						
						ExcelKeywords.setValueToCellByIndex(sheet01, n, 7, y)
			
						ExcelKeywords.setValueToCellByIndex(sheet01, n, 8, q)
			
						ExcelKeywords.setValueToCellByIndex(sheet01, n, 9, deitel)
			
							n = (n + 1)
			
							ExcelKeywords.saveWorkbook(excelFilePath, workbook01)
							println(mass1 +' good '+ mass2 )
						} else {
							if (WebUI.verifyTextNotPresent('Произошла ошибка при выполнении пользовательского кода', false) == false) {
							ExcelKeywords.setValueToCellByIndex(sheet01, n, 3, dashboardName)
			
							ExcelKeywords.setValueToCellByIndex(sheet01, n, 4, mass1)
						
							ExcelKeywords.setValueToCellByIndex(sheet01, n, 5, mass2)
						
							ExcelKeywords.setValueToCellByIndex(sheet01, n, 6, mass3)
						
							ExcelKeywords.setValueToCellByIndex(sheet01, n, 7, y)
			
							ExcelKeywords.setValueToCellByIndex(sheet01, n, 8, q)
			
							ExcelKeywords.setValueToCellByIndex(sheet01, n, 9, deitel)
			
								n = (n + 1)
			
								ExcelKeywords.saveWorkbook(excelFilePath, workbook01)
								println(mass1 +' good '+ mass2 )
							} else {
								if (WebUI.verifyTextNotPresent('У виджета нет данных', false) == false) {
								ExcelKeywords.setValueToCellByIndex(sheet01, n, 3, dashboardName)
			
								ExcelKeywords.setValueToCellByIndex(sheet01, n, 4, mass1)
						
								ExcelKeywords.setValueToCellByIndex(sheet01, n, 5, mass2)
						
								ExcelKeywords.setValueToCellByIndex(sheet01, n, 6, mass3)
						
								ExcelKeywords.setValueToCellByIndex(sheet01, n, 7, y)
			
								ExcelKeywords.setValueToCellByIndex(sheet01, n, 8, q)
			
								ExcelKeywords.setValueToCellByIndex(sheet01, n, 9, deitel)
								n = (n + 1)
			
									ExcelKeywords.saveWorkbook(excelFilePath, workbook01)
									println(mass1 +' good '+ mass2 )
								} else {
									if (WebUI.verifyTextNotPresent('Некорректные фильтры', false) == false) {
									ExcelKeywords.setValueToCellByIndex(sheet01, n, 3, dashboardName)
										
									ExcelKeywords.setValueToCellByIndex(sheet01, n, 4, mass1)
													
									ExcelKeywords.setValueToCellByIndex(sheet01, n, 5, mass2)
													
									ExcelKeywords.setValueToCellByIndex(sheet01, n, 6, mass3)
												
									ExcelKeywords.setValueToCellByIndex(sheet01, n, 7, y)
										
									ExcelKeywords.setValueToCellByIndex(sheet01, n, 8, q)
										
									ExcelKeywords.setValueToCellByIndex(sheet01, n, 9, deitel)
			
									n = (n + 1)
			
									ExcelKeywords.saveWorkbook(excelFilePath, workbook01)
										println(mass1 +' good '+ mass2 )
									}
								}
							}
						}
					}
				}
		
		println(mass1 +' '+ mass2 + ' ' + mass3)
		//println(mass2)
		//println(mass3)
	}
	
}
