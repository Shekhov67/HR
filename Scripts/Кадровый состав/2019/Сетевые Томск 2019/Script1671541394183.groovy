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

WebUI.click(findTestObject('Кадровый состав/div_2019'))

WebUI.click(findTestObject('Кадровый состав/Применить в фильтре Дата'))

WebUI.click(findTestObject('Кадровый состав/Фильтр ДЗО'))

WebUI.click(findTestObject('Кадровый состав/Снять выделение в фильтре ДЗО'))

WebUI.click(findTestObject('Кадровый состав/Развернуть Сетевые'))

WebUI.scrollToElement(findTestObject('Общие элементы/Развернуть Россети Томск'), 30)

WebUI.click(findTestObject('Кадровый состав/Россети Томск'))

WebUI.click(findTestObject('Кадровый состав/Применить в фильтре ДЗО'))

WebUI.verifyTextNotPresent('нет данных', false)

WebUI.verifyTextNotPresent('Ошибка запроса данных', false)

def percentCheck = Percents()

WebUI.click(findTestObject('Кадровый состав/Фильтр ДЗО'))

WebUI.click(findTestObject('Кадровый состав/Снять выделение в фильтре ДЗО'))

WebUI.scrollToElement(findTestObject('Общие элементы/Развернуть Россети Томск'), 30)

WebUI.click(findTestObject('Общие элементы/Развернуть Россети Томск'))

WebUI.click(findTestObject('Общие элементы/Томская распределительная компания'))

WebUI.click(findTestObject('Кадровый состав/Применить в фильтре ДЗО'))

WebUI.verifyTextNotPresent('нет данных', false)

WebUI.verifyTextNotPresent('Ошибка запроса данных', false)

percentCheck = Percents()

WebUI.closeBrowser()

static def Percents() {
    if (WebUI.verifyTextNotPresent('Ошибка запроса данных', false) == true) {
        String percent = WebUI.getText(findTestObject('Кадровый состав/percentOfWomen'))

        int indexP = percent.indexOf('%')

        WebUI.verifyGreaterThan(indexP, -1)

        String stringP = percent.replaceAll('\\s*', '')

        stringP = stringP.replaceAll('\\pP', '')

        println(stringP)

        String regex = '\\d+'

        WebUI.verifyMatch(stringP, regex, true)

        percent = WebUI.getText(findTestObject('Кадровый состав/percentOfMen'))

        indexP = percent.indexOf('%')

        WebUI.verifyGreaterThan(indexP, -1)

        stringP = percent.replaceAll('\\s*', '')

        stringP = stringP.replaceAll('\\pP', '')

        println(stringP)

        regex = '\\d+'

        WebUI.verifyMatch(stringP, regex, true)

        percent = WebUI.getText(findTestObject('Кадровый состав/percentExperience1'))

        if (percent != null) {
            indexP = percent.indexOf('%')

            WebUI.verifyGreaterThan(indexP, -1)

            stringP = percent.replaceAll('\\s*', '')

            stringP = stringP.replaceAll('\\pP', '')

            println(stringP)

            regex = '\\d+'

            WebUI.verifyMatch(stringP, regex, true)
        } else {
            percent = WebUI.getText(findTestObject('Кадровый состав/percentExperience3'))

            if (percent != null) {
                indexP = percent.indexOf('%')

                WebUI.verifyGreaterThan(indexP, -1)

                stringP = percent.replaceAll('\\s*', '')

                stringP = stringP.replaceAll('\\pP', '')

                println(stringP)

                regex = '\\d+'

                WebUI.verifyMatch(stringP, regex, true)
            } else {
                percent = WebUI.getText(findTestObject('Кадровый состав/percentExperience10'))

                if (percent != null) {
                    indexP = percent.indexOf('%')

                    WebUI.verifyGreaterThan(indexP, -1)

                    stringP = percent.replaceAll('\\s*', '')

                    stringP = stringP.replaceAll('\\pP', '')

                    println(stringP)

                    regex = '\\d+'

                    WebUI.verifyMatch(stringP, regex, true)
                } else {
                    percent = WebUI.getText(findTestObject('Кадровый состав/percentExperience20'))

                    if (percent != null) {
                        indexP = percent.indexOf('%')

                        WebUI.verifyGreaterThan(indexP, -1)

                        stringP = percent.replaceAll('\\s*', '')

                        stringP = stringP.replaceAll('\\pP', '')

                        println(stringP)

                        regex = '\\d+'

                        WebUI.verifyMatch(stringP, regex, true)
                    } else {
                        percent = WebUI.getText(findTestObject('Кадровый состав/percentExperience30'))

                        if (percent != null) {
                            indexP = percent.indexOf('%')

                            WebUI.verifyGreaterThan(indexP, -1)

                            stringP = percent.replaceAll('\\s*', '')

                            stringP = stringP.replaceAll('\\pP', '')

                            println(stringP)

                            regex = '\\d+'

                            WebUI.verifyMatch(stringP, regex, true)
                        } else {
                            percent = WebUI.getText(findTestObject('Кадровый состав/percentExperience40'))

                            if (percent != null) {
                                indexP = percent.indexOf('%')

                                WebUI.verifyGreaterThan(indexP, -1)

                                stringP = percent.replaceAll('\\s*', '')

                                stringP = stringP.replaceAll('\\pP', '')

                                println(stringP)

                                regex = '\\d+'

                                WebUI.verifyMatch(stringP, regex, true)
                            } else {
                                percent = WebUI.getText(findTestObject('Кадровый состав/percentExperienceMoreThan40'))

                                if (percent != null) {
                                    indexP = percent.indexOf('%')

                                    WebUI.verifyGreaterThan(indexP, -1)

                                    stringP = percent.replaceAll('\\s*', '')

                                    stringP = stringP.replaceAll('\\pP', '')

                                    println(stringP)

                                    regex = '\\d+'

                                    WebUI.verifyMatch(stringP, regex, true)
                                } else {
                                    percent = WebUI.getText(findTestObject('Кадровый состав/percentEducationHigher'))

                                    if (percent != null) {
                                        indexP = percent.indexOf('%')

                                        WebUI.verifyGreaterThan(indexP, -1)

                                        stringP = percent.replaceAll('\\s*', '')

                                        stringP = stringP.replaceAll('\\pP', '')

                                        println(stringP)

                                        regex = '\\d+'

                                        WebUI.verifyMatch(stringP, regex, true)
                                    } else {
                                        percent = WebUI.getText(findTestObject('Кадровый состав/percentEducationLower'))

                                        if (percent != null) {
                                            indexP = percent.indexOf('%')

                                            WebUI.verifyGreaterThan(indexP, -1)

                                            stringP = percent.replaceAll('\\s*', '')

                                            stringP = stringP.replaceAll('\\pP', '')

                                            println(stringP)

                                            regex = '\\d+'

                                            WebUI.verifyMatch(stringP, regex, true)
                                        } else {
                                            percent = WebUI.getText(findTestObject('Кадровый состав/percentEducationMiddle'))

                                            if (percent != null) {
                                                indexP = percent.indexOf('%')

                                                WebUI.verifyGreaterThan(indexP, -1)

                                                stringP = percent.replaceAll('\\s*', '')

                                                stringP = stringP.replaceAll('\\pP', '')

                                                println(stringP)

                                                regex = '\\d+'

                                                WebUI.verifyMatch(stringP, regex, true)
                                            } else {
                                                println('Ошибка отображения данных в графике "Стаж"')
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
        
        percent = WebUI.getText(findTestObject('Кадровый состав/percentEducationLower'))

        if (percent != null) {
            indexP = percent.indexOf('%')

            WebUI.verifyGreaterThan(indexP, -1)

            stringP = percent.replaceAll('\\s*', '')

            stringP = stringP.replaceAll('\\pP', '')

            println(stringP)

            regex = '\\d+'

            WebUI.verifyMatch(stringP, regex, true)
        } else {
            percent = WebUI.getText(findTestObject('Кадровый состав/percentEducationMiddle'))

            if (percent != null) {
                indexP = percent.indexOf('%')

                WebUI.verifyGreaterThan(indexP, -1)

                stringP = percent.replaceAll('\\s*', '')

                stringP = stringP.replaceAll('\\pP', '')

                println(stringP)

                regex = '\\d+'

                WebUI.verifyMatch(stringP, regex, true)
            } else {
                percent = WebUI.getText(findTestObject('Кадровый состав/percentEducationHigher'))

                if (percent != null) {
                    indexP = percent.indexOf('%')

                    WebUI.verifyGreaterThan(indexP, -1)

                    stringP = percent.replaceAll('\\s*', '')

                    stringP = stringP.replaceAll('\\pP', '')

                    println(stringP)

                    regex = '\\d+'

                    WebUI.verifyMatch(stringP, regex, true)
                } else {
                    println('Ошибка отображения данных в графике "Образование"')
                }
            }
        }
        
        percent = WebUI.getText(findTestObject('Кадровый состав/percentAgeBelow35'))

        if (percent != null) {
            indexP = percent.indexOf('%')

            WebUI.verifyGreaterThan(indexP, -1)

            stringP = percent.replaceAll('\\s*', '')

            stringP = stringP.replaceAll('\\pP', '')

            println(stringP)

            regex = '\\d+'

            WebUI.verifyMatch(stringP, regex, true)
        } else {
            percent = WebUI.getText(findTestObject('Кадровый состав/percentAgeAbove35'))

            if (percent != null) {
                indexP = percent.indexOf('%')

                WebUI.verifyGreaterThan(indexP, -1)

                stringP = percent.replaceAll('\\s*', '')

                stringP = stringP.replaceAll('\\pP', '')

                println(stringP)

                regex = '\\d+'

                WebUI.verifyMatch(stringP, regex, true)
            } else {
                percent = WebUI.getText(findTestObject('percentAgeBelow50'))

                if (percent != null) {
                    indexP = percent.indexOf('%')

                    WebUI.verifyGreaterThan(indexP, -1)

                    stringP = percent.replaceAll('\\s*', '')

                    stringP = stringP.replaceAll('\\pP', '')

                    println(stringP)

                    regex = '\\d+'

                    WebUI.verifyMatch(stringP, regex, true)
                } else {
                    println('Ошибка отображения данных в графике "Образование"')
                }
            }
        }
    }
}

