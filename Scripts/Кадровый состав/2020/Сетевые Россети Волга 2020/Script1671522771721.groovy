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

WebUI.openBrowser(findTestData('New Test Data').getValue(3, 1))

WebUI.setText(findTestObject('Page_/input__username'), findTestData('New Test Data').getValue(1, 1))

WebUI.setText(findTestObject('Page_/input__password'), findTestData('New Test Data').getValue(2, 1))

WebUI.click(findTestObject('Page_/button_'))

WebUI.click(findTestObject('Кадровый состав/Фильтр Дата'))

WebUI.click(findTestObject('Кадровый состав/Снять выделение в фильтре Дата'))

WebUI.click(findTestObject('Кадровый состав/div_2020'))

WebUI.click(findTestObject('Кадровый состав/Применить в фильтре Дата'))

WebUI.click(findTestObject('Кадровый состав/Фильтр ДЗО'))

WebUI.click(findTestObject('Кадровый состав/Снять выделение в фильтре ДЗО'))

WebUI.click(findTestObject('Кадровый состав/Развернуть Сетевые'))

WebUI.scrollToElement(findTestObject('Общие элементы/Развернуть Чеченэнерго'), 30)

WebUI.click(findTestObject('Кадровый состав/Россети Волга'))

WebUI.click(findTestObject('Кадровый состав/Применить в фильтре ДЗО'))

WebUI.click(findTestObject('Кадровый состав/Фильтр ДЗО'))

WebUI.click(findTestObject('Кадровый состав/Снять выделение в фильтре ДЗО'))

WebUI.scrollToElement(findTestObject('Общие элементы/Развернуть Чеченэнерго'), 30)

WebUI.click(findTestObject('Общие элементы/Развернуть Россети Волга'))

WebUI.click(findTestObject('Кадровый состав/Исполнительный аппарат Россети Волга'))

WebUI.click(findTestObject('Кадровый состав/Применить в фильтре ДЗО'))

WebUI.click(findTestObject('Кадровый состав/Фильтр ДЗО'))

WebUI.click(findTestObject('Кадровый состав/Снять выделение в фильтре ДЗО'))

WebUI.scrollToElement(findTestObject('Общие элементы/Развернуть Россети Волга'), 30)

WebUI.click(findTestObject('Общие элементы/Мордовэнерго'))

WebUI.click(findTestObject('Кадровый состав/Применить в фильтре ДЗО'))

WebUI.click(findTestObject('Кадровый состав/Фильтр ДЗО'))

WebUI.click(findTestObject('Кадровый состав/Снять выделение в фильтре ДЗО'))

WebUI.scrollToElement(findTestObject('Общие элементы/Развернуть Россети Волга'), 30)

WebUI.click(findTestObject('Общие элементы/Оренбургэнерго'))

WebUI.click(findTestObject('Кадровый состав/Применить в фильтре ДЗО'))

WebUI.click(findTestObject('Кадровый состав/Фильтр ДЗО'))

WebUI.click(findTestObject('Кадровый состав/Снять выделение в фильтре ДЗО'))

WebUI.scrollToElement(findTestObject('Общие элементы/Развернуть Россети Волга'), 30)

WebUI.click(findTestObject('Общие элементы/Пензаэнерго'))

WebUI.click(findTestObject('Кадровый состав/Применить в фильтре ДЗО'))

WebUI.click(findTestObject('Кадровый состав/Фильтр ДЗО'))

WebUI.click(findTestObject('Кадровый состав/Снять выделение в фильтре ДЗО'))

WebUI.scrollToElement(findTestObject('Общие элементы/Пензаэнерго'), 30)

WebUI.click(findTestObject('Общие элементы/Самарские РС'))

WebUI.click(findTestObject('Кадровый состав/Применить в фильтре ДЗО'))

WebUI.click(findTestObject('Кадровый состав/Фильтр ДЗО'))

WebUI.click(findTestObject('Кадровый состав/Снять выделение в фильтре ДЗО'))

WebUI.scrollToElement(findTestObject('Общие элементы/Пензаэнерго'), 30)

WebUI.click(findTestObject('Общие элементы/Саратовские РС'))

WebUI.click(findTestObject('Кадровый состав/Применить в фильтре ДЗО'))

WebUI.click(findTestObject('Кадровый состав/Фильтр ДЗО'))

WebUI.click(findTestObject('Кадровый состав/Снять выделение в фильтре ДЗО'))

WebUI.scrollToElement(findTestObject('Общие элементы/Пензаэнерго'), 30)

WebUI.click(findTestObject('Общие элементы/Ульяновские РС'))

WebUI.click(findTestObject('Кадровый состав/Применить в фильтре ДЗО'))

WebUI.click(findTestObject('Кадровый состав/Фильтр ДЗО'))

WebUI.click(findTestObject('Кадровый состав/Снять выделение в фильтре ДЗО'))

WebUI.scrollToElement(findTestObject('Общие элементы/Пензаэнерго'), 30)

WebUI.click(findTestObject('Общие элементы/Чувашэнерго'))

WebUI.click(findTestObject('Кадровый состав/Применить в фильтре ДЗО'))

