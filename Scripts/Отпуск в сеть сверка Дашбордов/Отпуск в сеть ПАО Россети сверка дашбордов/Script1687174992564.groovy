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
import org.apache.commons.lang3.StringUtils as StringUtils

WebUI.openBrowser('')

'БЛОК РУКОВОДИТЕЛЕЙ'
WebUI.navigateToUrl(findTestData('Test Data').getValue(7, 5))

WebUI.setText(findTestObject('Общие/input__username'), findTestData('Test Data').getValue(5, 1))

WebUI.setText(findTestObject('Общие/input__password'), findTestData('Test Data').getValue(6, 1))

WebUI.click(findTestObject('Общие в сеть/button_'))

WebUI.click(findTestObject('Общие в сеть/Фильтр ДЗО'))

WebUI.click(findTestObject('Общие в сеть/Снять выделения в фильтре ДЗО'))

WebUI.click(findTestObject('Общие в сеть/Применить в фильтре ДЗО'))

def scan = ScanErrors()

String a = WebUI.getText(findTestObject('Отпуск в сеть сверка/Данные со страницы Отпуск в сеть/Page_Visiology Platform/a1'))

println(a)

String a0 = a.replaceAll('\\s+', '')

println(a0)

println(a0.length())

int numA = a0.length()

String a02 = WebUI.getText(findTestObject('Отпуск в сеть сверка/Данные со страницы Отпуск в сеть/Page_Visiology Platform/a2'))

println(a02)

String a2 = a02.replaceAll('\\s+', '')

println(a2)

println(a2.length())

int numA2 = a2.length()

int numA02 = a2.length() * 2

println(numA02)

String a03 = WebUI.getText(findTestObject('Отпуск в сеть сверка/Данные со страницы Отпуск в сеть/Page_Visiology Platform/a3'))

println(a03)

String a3 = a03.replaceAll('\\s+', '')

println(a3)

println(a3.length())

int numA3 = a3.length()

int numA03 = a3.length() * 4

println(numA03)

println(a0)

println(a2)

println(a3)

'БАЛАНСЫ'
WebUI.navigateToUrl(findTestData('Test Data').getValue(7, 6))

WebUI.delay(10)

String b = WebUI.getText(findTestObject('Отпуск в сеть сверка/Данные со страницы Отпуск в сеть/данные с виджета в блоке Балансы'))

println(b)

String b1 = b.replaceAll('\\D+', '')

println(b1)

String numberB = StringUtils.substring(b1, 0, numA)

println(numberB)

if (WebUI.verifyEqual(a0, numberB)) {
    println('GOOD')
} else {
    String v = 'Виджет отпуск в сеть'

    def write = WriteToExcel(a0, numberB, v)
}

String b02 = WebUI.getText(findTestObject('Отпуск в сеть сверка/Данные со страницы Отпуск в сеть/данные с виджета в блоке Балансы'))

println(b02)

String b2 = b02.replaceAll('\\D+', '')

println(b2)

String numberB2 = b2.substring(numA02).substring(0, numA2)

println(numberB2)

if (WebUI.verifyEqual(a2, numberB2)) {
    println('GOOD')
} else {
    String v = 'Виджет отпуск в сеть'

    def write = WriteToExcel(a2, numberB2, v)
}

String b03 = WebUI.getText(findTestObject('Отпуск в сеть сверка/Данные со страницы Отпуск в сеть/данные с виджета в блоке Балансы'))

println(b03)

String b3 = b03.replaceAll('\\D+', '')

println(b3)

String numberB3 = b3.substring(numA03).substring(0, numA3)

println(numberB3)

if (WebUI.verifyEqual(a3, numberB3)) {
    println('GOOD')
} else {
    String v = 'Виджет отпуск в сеть'

    def write = WriteToExcel(a3, numberB3, v)
}

WebUI.closeBrowser()

static def WriteToExcel(def a, def b, def v) {
    String sheetName = 'List1'

    def data = findTestData('Test Data')

    int n = data.getRowNumbers() + 1

    String dZO = WebUI.getText(findTestObject('Отпуск в сеть сверка/фильтр ДЗО в Блоке Балансы'))

    String year = WebUI.getText(findTestObject('Отпуск в сеть сверка/фильтр Дата в блоке Балансы'))

    println(year)

    String dashboardName = 'Отпуск в сеть'

    String page = 'Данные не  совпадают'

    def workbook01 = ExcelKeywords.getWorkbook(GlobalVariable.excelFilePath)

    def sheet01 = ExcelKeywords.getExcelSheet(workbook01, sheetName)

    ExcelKeywords.setValueToCellByIndex(sheet01, n, 0, dashboardName)

    ExcelKeywords.setValueToCellByIndex(sheet01, n, 1, dZO)

    ExcelKeywords.setValueToCellByIndex(sheet01, n, 2, year)

    ExcelKeywords.setValueToCellByIndex(sheet01, n, 3, a)

    ExcelKeywords.setValueToCellByIndex(sheet01, n, 4, b)

    ExcelKeywords.setValueToCellByIndex(sheet01, n, 5, v)

    ExcelKeywords.setValueToCellByIndex(sheet01, n, 6, page)

    n = (n + 1)

    ExcelKeywords.saveWorkbook(GlobalVariable.excelFilePath, workbook01)
}

static def ScanErrors() {
    if (WebUI.verifyTextNotPresent('нет данных', false) == false) {
        def write = WriteToExcel2(def page = 'У виджета нет данных')
    } else if (WebUI.verifyTextNotPresent('Ошибка запроса данных', false) == false) {
        def write = WriteToExcel2(def page = 'Ошибка запроса данных')
    } else if (WebUI.verifyTextNotPresent('Произошла ошибка при выполнении пользовательского кода', false) == false) {
        def write = WriteToExcel2(def page = 'Произошла ошибка при выполнении пользовательского кода')
    } else if (WebUI.verifyTextNotPresent('У виджета нет данных', false) == false) {
        def write = WriteToExcel2(def page = 'У виджета нет данных')
    } else if (WebUI.verifyTextNotPresent('Некорректные фильтры', false) == false) {
        def write = WriteToExcel2(def page = 'Некорректные фильтры')
    }
}

static def WriteToExcel2(def page) {
    String sheetName = 'List1'

    def data = findTestData('Test Data')

    int n = data.getRowNumbers() + 1

    String dZO = WebUI.getText(findTestObject('Отпуск в сеть(виджеты)/фильтр ДЗО'))

    String year = WebUI.getText(findTestObject('Отпуск в сеть(виджеты)/фильтр Дата'))

    println(year)

    String dashboardName = 'Отпуск в сеть'

    def workbook01 = ExcelKeywords.getWorkbook(GlobalVariable.excelFilePath)

    def sheet01 = ExcelKeywords.getExcelSheet(workbook01, sheetName)

    if (WebUI.verifyTextNotPresent('нет данных', false) == false) {
        ExcelKeywords.setValueToCellByIndex(sheet01, n, 0, dashboardName)

        ExcelKeywords.setValueToCellByIndex(sheet01, n, 1, dZO)

        ExcelKeywords.setValueToCellByIndex(sheet01, n, 2, year)

        ExcelKeywords.setValueToCellByIndex(sheet01, n, 3, page)

        n = (n + 1)

        ExcelKeywords.saveWorkbook(GlobalVariable.excelFilePath, workbook01)

        WebUI.closeBrowser()
    } else {
        if (WebUI.verifyTextNotPresent('Ошибка запроса данных', false) == false) {
            ExcelKeywords.setValueToCellByIndex(sheet01, n, 0, dashboardName)

            ExcelKeywords.setValueToCellByIndex(sheet01, n, 1, dZO)

            ExcelKeywords.setValueToCellByIndex(sheet01, n, 2, year)

            ExcelKeywords.setValueToCellByIndex(sheet01, n, 3, page)

            n = (n + 1)

            ExcelKeywords.saveWorkbook(GlobalVariable.excelFilePath, workbook01)
        } else {
            if (WebUI.verifyTextNotPresent('Произошла ошибка при выполнении пользовательского кода', false) == false) {
                ExcelKeywords.setValueToCellByIndex(sheet01, n, 0, dashboardName)

                ExcelKeywords.setValueToCellByIndex(sheet01, n, 1, dZO)

                ExcelKeywords.setValueToCellByIndex(sheet01, n, 2, year)

                ExcelKeywords.setValueToCellByIndex(sheet01, n, 3, page)

                n = (n + 1)

                ExcelKeywords.saveWorkbook(GlobalVariable.excelFilePath, workbook01)
            } else {
                if (WebUI.verifyTextNotPresent('У виджета нет данных', false) == false) {
                    ExcelKeywords.setValueToCellByIndex(sheet01, n, 0, dashboardName)

                    ExcelKeywords.setValueToCellByIndex(sheet01, n, 1, dZO)

                    ExcelKeywords.setValueToCellByIndex(sheet01, n, 2, year)

                    ExcelKeywords.setValueToCellByIndex(sheet01, n, 3, page)

                    n = (n + 1)

                    ExcelKeywords.saveWorkbook(GlobalVariable.excelFilePath, workbook01)
                } else {
                    if (WebUI.verifyTextNotPresent('Некорректные фильтры', false) == false) {
                        ExcelKeywords.setValueToCellByIndex(sheet01, n, 0, dashboardName)

                        ExcelKeywords.setValueToCellByIndex(sheet01, n, 1, dZO)

                        ExcelKeywords.setValueToCellByIndex(sheet01, n, 2, year)

                        ExcelKeywords.setValueToCellByIndex(sheet01, n, 3, page)

                        n = (n + 1)

                        ExcelKeywords.saveWorkbook(GlobalVariable.excelFilePath, workbook01)
                    }
                }
            }
        }
    }
}

