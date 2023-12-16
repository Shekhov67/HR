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

WebUI.openBrowser(findTestData('New Test Data').getValue(3, 4))

WebUI.setText(findTestObject('Page_/input__username'), findTestData('New Test Data').getValue(1, 1))

WebUI.setText(findTestObject('Page_/input__password'), findTestData('New Test Data').getValue(2, 1))

WebUI.click(findTestObject('Page_/button_'))

WebUI.click(findTestObject('Численность персонала/Фильтр Дата'))

WebUI.click(findTestObject('Численность персонала/Снять выделение в фильтре Дата'))

WebUI.click(findTestObject('Численность персонала/2015'))

WebUI.click(findTestObject('Численность персонала/Применить в фильтре Дата'))

WebUI.click(findTestObject('Численность персонала/Фильтр ДЗО'))

WebUI.click(findTestObject('Численность персонала/Снять выделение в фильтре ДЗО'))

WebUI.click(findTestObject('Численность персонала/Обеспечивающие'))

WebUI.click(findTestObject('Численность персонала/Применить в фильтре ДЗО'))

WebUI.verifyTextNotPresent('нет данных', false)

WebUI.verifyTextNotPresent('Ошибка запроса данных', false)

//def percentCheck = Percents()
WebUI.click(findTestObject('Численность персонала/Фильтр ДЗО'))

WebUI.click(findTestObject('Численность персонала/Снять выделение в фильтре ДЗО'))

WebUI.click(findTestObject('Численность персонала/Развернуть Обеспечивающие'))

WebUI.click(findTestObject('Численность персонала/НИЦ ЕЭС верхний уровень'))

WebUI.click(findTestObject('Численность персонала/Применить в фильтре ДЗО'))

WebUI.verifyTextNotPresent('нет данных', false)

WebUI.verifyTextNotPresent('Ошибка запроса данных', false)

//percentCheck = Percents()
WebUI.click(findTestObject('Численность персонала/Фильтр ДЗО'))

WebUI.click(findTestObject('Численность персонала/Снять выделение в фильтре ДЗО'))

WebUI.click(findTestObject('Общие элементы/Развернуть НИЦ ЕЭС'))

WebUI.click(findTestObject('Численность персонала/НИЦ ЕЭС нижний уровень'))

WebUI.click(findTestObject('Численность персонала/Применить в фильтре ДЗО'))

WebUI.verifyTextNotPresent('нет данных', false)

WebUI.verifyTextNotPresent('Ошибка запроса данных', false)

WebUI.click(findTestObject('Численность персонала/Фильтр ДЗО'))

WebUI.click(findTestObject('Численность персонала/НИЦ ЕЭС нижний уровень'))

WebUI.click(findTestObject('Численность персонала/НИЦ ЕЭС_2'))

WebUI.click(findTestObject('Численность персонала/Применить в фильтре ДЗО'))

WebUI.verifyTextNotPresent('нет данных', false)

WebUI.verifyTextNotPresent('Ошибка запроса данных', false)

//percentCheck = Percents()
WebUI.click(findTestObject('Численность персонала/Фильтр ДЗО'))

WebUI.click(findTestObject('Численность персонала/Снять выделение в фильтре ДЗО'))

WebUI.click(findTestObject('Численность персонала/Центр технического заказчика'))

WebUI.click(findTestObject('Численность персонала/Применить в фильтре ДЗО'))

WebUI.verifyTextNotPresent('нет данных', false)

WebUI.verifyTextNotPresent('Ошибка запроса данных', false)

//percentCheck = Percents()
WebUI.click(findTestObject('Численность персонала/Фильтр ДЗО'))

WebUI.click(findTestObject('Численность персонала/Снять выделение в фильтре ДЗО'))

WebUI.click(findTestObject('Численность персонала/Развернуть Центр технического заказчика'))

WebUI.click(findTestObject('Численность персонала/Центр технического заказчика нижний уровень'))

WebUI.click(findTestObject('Численность персонала/Применить в фильтре ДЗО'))

WebUI.verifyTextNotPresent('нет данных', false)

WebUI.verifyTextNotPresent('Ошибка запроса данных', false)

//percentCheck = Percents()
WebUI.click(findTestObject('Численность персонала/Фильтр ДЗО'))

WebUI.scrollToElement(findTestObject('Численность персонала/ДЗО ПАО ФИЦ'), 30)

WebUI.scrollToElement(findTestObject('Численность персонала/a_DZO'), 30)

WebUI.click(findTestObject('Численность персонала/Снять выделение в фильтре ДЗО'))

WebUI.click(findTestObject('Общие элементы/Развернуть ДЗО ПАО ФИЦ'))

WebUI.click(findTestObject('Общие элементы/Оператор АСТУ ДЗО ПАО ФИЦ'))

WebUI.click(findTestObject('Численность персонала/Применить в фильтре ДЗО'))

WebUI.verifyTextNotPresent('нет данных', false)

WebUI.verifyTextNotPresent('Ошибка запроса данных', false)

//percentCheck = Percents()
WebUI.click(findTestObject('Численность персонала/Фильтр ДЗО'))

WebUI.click(findTestObject('Численность персонала/Снять выделение в фильтре ДЗО'))

WebUI.scrollToElement(findTestObject('Численность персонала/скролл сетевые'), 30)

WebUI.scrollToElement(findTestObject('Численность персонала/a_DZO'), 30)

WebUI.click(findTestObject('Численность персонала/Снять выделение в фильтре ДЗО'))

WebUI.click(findTestObject('Общие элементы/Развернуть Управление ВОЛС-ВЛ'), FailureHandling.CONTINUE_ON_FAILURE)

WebUI.click(findTestObject('Общие элементы/Управление ВОЛС-ВЛ нижний уровень'))

WebUI.click(findTestObject('Численность персонала/Применить в фильтре ДЗО'))

WebUI.verifyTextNotPresent('нет данных', false)

WebUI.verifyTextNotPresent('Ошибка запроса данных', false)

//percentCheck = Percents()
WebUI.closeBrowser()

/*static def Percents() {
    if (WebUI.verifyTextNotPresent('Ошибка запроса данных', false) == true) {
        String percent = WebUI.getText(findTestObject('Численность персонала/precentSumPerson93'))

        int indexP = percent.indexOf('%')

        WebUI.verifyGreaterThan(indexP, -1)

        String stringP = percent.replaceAll('\\s*', '')

        stringP = stringP.replaceAll('\\pP', '')

        println(stringP)

        String regex = '\\d+'

        WebUI.verifyMatch(stringP, regex, true)

        percent = WebUI.getText(findTestObject('Численность персонала/complect_PP94'))

        indexP = percent.indexOf('%')

        WebUI.verifyGreaterThan(indexP, -1)

        stringP = percent.replaceAll('\\s*', '')

        stringP = stringP.replaceAll('\\pP', '')

        println(stringP)

        regex = '\\d+'

        WebUI.verifyMatch(stringP, regex, true)

        percent = WebUI.getText(findTestObject('Численность персонала/totalNumberPercent24'))

        if (percent != null) {
            indexP = percent.indexOf('%')

            WebUI.verifyGreaterThan(indexP, -1)

            stringP = percent.replaceAll('\\s*', '')

            stringP = stringP.replaceAll('\\pP', '')

            println(stringP)

            regex = '\\d+'

            WebUI.verifyMatch(stringP, regex, true)
        } else {
            percent = WebUI.getText(findTestObject('Численность персонала/total_list'))

            if (percent != null) {
                indexP = percent.indexOf('%')

                WebUI.verifyGreaterThan(indexP, -1)

                stringP = percent.replaceAll('\\s*', '')

                stringP = stringP.replaceAll('\\pP', '')

                println(stringP)

                regex = '\\d+'

                WebUI.verifyMatch(stringP, regex, true)
            } else {
                percent = WebUI.getText(findTestObject('Численность персонала/totalNoList'))

                if (percent != null) {
                    indexP = percent.indexOf('%')

                    WebUI.verifyGreaterThan(indexP, -1)

                    stringP = percent.replaceAll('\\s*', '')

                    stringP = stringP.replaceAll('\\pP', '')

                    println(stringP)

                    regex = '\\d+'

                    WebUI.verifyMatch(stringP, regex, true)
                } else {
                    percent = WebUI.getText(findTestObject('Численность персонала/percentExperience20'))

                    if (percent != null) {
                        indexP = percent.indexOf('%')

                        WebUI.verifyGreaterThan(indexP, -1)

                        stringP = percent.replaceAll('\\s*', '')

                        stringP = stringP.replaceAll('\\pP', '')

                        println(stringP)

                        regex = '\\d+'

                        WebUI.verifyMatch(stringP, regex, true)
                    } else {
                        percent = WebUI.getText(findTestObject('Численность персонала/percentExperience30'))

                        if (percent != null) {
                            indexP = percent.indexOf('%')

                            WebUI.verifyGreaterThan(indexP, -1)

                            stringP = percent.replaceAll('\\s*', '')

                            stringP = stringP.replaceAll('\\pP', '')

                            println(stringP)

                            regex = '\\d+'

                            WebUI.verifyMatch(stringP, regex, true)
                        } else {
                            percent = WebUI.getText(findTestObject('Численность персонала/percentExperience40'))

                            if (percent != null) {
                                indexP = percent.indexOf('%')

                                WebUI.verifyGreaterThan(indexP, -1)

                                stringP = percent.replaceAll('\\s*', '')

                                stringP = stringP.replaceAll('\\pP', '')

                                println(stringP)

                                regex = '\\d+'

                                WebUI.verifyMatch(stringP, regex, true)
                            } else {
                                percent = WebUI.getText(findTestObject('Численность персонала/percentExperienceMoreThan40'))

                                if (percent != null) {
                                    indexP = percent.indexOf('%')

                                    WebUI.verifyGreaterThan(indexP, -1)

                                    stringP = percent.replaceAll('\\s*', '')

                                    stringP = stringP.replaceAll('\\pP', '')

                                    println(stringP)

                                    regex = '\\d+'

                                    WebUI.verifyMatch(stringP, regex, true)
                                } else {
                                    percent = WebUI.getText(findTestObject('Численность персонала/percentEducationHigher'))

                                    if (percent != null) {
                                        indexP = percent.indexOf('%')

                                        WebUI.verifyGreaterThan(indexP, -1)

                                        stringP = percent.replaceAll('\\s*', '')

                                        stringP = stringP.replaceAll('\\pP', '')

                                        println(stringP)

                                        regex = '\\d+'

                                        WebUI.verifyMatch(stringP, regex, true)
                                    } else {
                                        percent = WebUI.getText(findTestObject('Численность персонала/percentEducationLower'))

                                        if (percent != null) {
                                            indexP = percent.indexOf('%')

                                            WebUI.verifyGreaterThan(indexP, -1)

                                            stringP = percent.replaceAll('\\s*', '')

                                            stringP = stringP.replaceAll('\\pP', '')

                                            println(stringP)

                                            regex = '\\d+'

                                            WebUI.verifyMatch(stringP, regex, true)
                                        } else {
                                            percent = WebUI.getText(findTestObject('Численность персонала/percentEducationMiddle'))

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
        
        percent = WebUI.getText(findTestObject('Численность персонала/percentEducationLower'))

        if (percent != null) {
            indexP = percent.indexOf('%')

            WebUI.verifyGreaterThan(indexP, -1)

            stringP = percent.replaceAll('\\s*', '')

            stringP = stringP.replaceAll('\\pP', '')

            println(stringP)

            regex = '\\d+'

            WebUI.verifyMatch(stringP, regex, true)
        } else {
            percent = WebUI.getText(findTestObject('Численность персонала/percentEducationMiddle'))

            if (percent != null) {
                indexP = percent.indexOf('%')

                WebUI.verifyGreaterThan(indexP, -1)

                stringP = percent.replaceAll('\\s*', '')

                stringP = stringP.replaceAll('\\pP', '')

                println(stringP)

                regex = '\\d+'

                WebUI.verifyMatch(stringP, regex, true)
            } else {
                percent = WebUI.getText(findTestObject('Численность персонала/percentEducationHigher'))

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
        
        percent = WebUI.getText(findTestObject('Численность персонала/percentAgeBelow35'))

        if (percent != null) {
            indexP = percent.indexOf('%')

            WebUI.verifyGreaterThan(indexP, -1)

            stringP = percent.replaceAll('\\s*', '')

            stringP = stringP.replaceAll('\\pP', '')

            println(stringP)

            regex = '\\d+'

            WebUI.verifyMatch(stringP, regex, true)
        } else {
            percent = WebUI.getText(findTestObject('Численность персонала/percentAgeAbove35'))

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
*/
