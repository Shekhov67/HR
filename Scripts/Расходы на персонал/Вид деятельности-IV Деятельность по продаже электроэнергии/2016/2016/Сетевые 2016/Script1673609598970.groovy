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

WebUI.click(findTestObject('Расходы на персонал/Фильтр Дата'))

WebUI.click(findTestObject('Расходы на персонал/Снять выделение в фильтре Дата'))

WebUI.click(findTestObject('Расходы на персонал/2016(1-4q)'))

WebUI.click(findTestObject('Расходы на персонал/Применить в фильтре Дата'))

WebUI.click(findTestObject('Расходы на персонал/фильтр Деятельность ДЗО'))

WebUI.click(findTestObject('Расходы на персонал/ДЕЯТЕЛЬНОСТЬ ПО ПРОДАЖЕ ЭЛЕКТРОЭНЕРГИИ'))

WebUI.click(findTestObject('Расходы на персонал/применить в фильтре деятельность'))

//percentCheck = Percents()
WebUI.click(findTestObject('Расходы на персонал/Фильтр ДЗО'))

WebUI.click(findTestObject('Расходы на персонал/Снять выделение в фильтре ДЗО'))

WebUI.click(findTestObject('Расходы на персонал/Сетевые'))

WebUI.click(findTestObject('Расходы на персонал/Применить в фильтре ДЗО'))

def write = WriteToExcel()

//def percentCheck = Percents()
WebUI.click(findTestObject('Расходы на персонал/Фильтр ДЗО'))

WebUI.click(findTestObject('Расходы на персонал/Снять выделение в фильтре ДЗО'))

WebUI.click(findTestObject('Расходы на персонал/Развернуть Сетевые'))

WebUI.click(findTestObject('Расходы на персонал/Тываэнерго верхний уровень'))

WebUI.click(findTestObject('Расходы на персонал/Применить в фильтре ДЗО'))

write = WriteToExcel()

WebUI.click(findTestObject('Расходы на персонал/Фильтр ДЗО'))

WebUI.click(findTestObject('Расходы на персонал/Снять выделение в фильтре ДЗО'))

WebUI.click(findTestObject('Расходы на персонал/развернуть Тываэнерго'))

WebUI.click(findTestObject('Расходы на персонал/Тываэнерго нижний уровень'))

WebUI.click(findTestObject('Расходы на персонал/Применить в фильтре ДЗО'))

write = WriteToExcel()

//percentCheck = Percents()
WebUI.click(findTestObject('Расходы на персонал/Фильтр ДЗО'))

WebUI.click(findTestObject('Расходы на персонал/Снять выделение в фильтре ДЗО'))

WebUI.click(findTestObject('Расходы на персонал/Чеченэнерго верхний уровень'))

WebUI.click(findTestObject('Расходы на персонал/Применить в фильтре ДЗО'))

write = WriteToExcel()

//percentCheck = Percents()
WebUI.click(findTestObject('Расходы на персонал/Фильтр ДЗО'))

WebUI.click(findTestObject('Расходы на персонал/Снять выделение в фильтре ДЗО'))

WebUI.scrollToElement(findTestObject('Общие элементы/Развернуть Чеченэнерго'), 30)

WebUI.scrollToElement(findTestObject('Расходы на персонал/a_DZO'), 30)

WebUI.click(findTestObject('Общие элементы/Развернуть Чеченэнерго'))

WebUI.click(findTestObject('Расходы на персонал/Чеченэнерго нижний уровень'))

WebUI.click(findTestObject('Расходы на персонал/Применить в фильтре ДЗО'))

write = WriteToExcel()

//percentCheck = Percents()
WebUI.closeBrowser()

