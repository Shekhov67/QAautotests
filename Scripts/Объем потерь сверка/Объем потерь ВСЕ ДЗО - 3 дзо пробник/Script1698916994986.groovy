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
import java.util.Date as Date
import java.text.SimpleDateFormat as SimpleDateFormat

'0'
def start = OpenBrowser()

String obemPoter = findTestData('Test Data').getValue(2, 87)

println(obemPoter)

String percentPoter = findTestData('Test Data').getValue(3, 87)

println(percentPoter)

def raspred = Raspred()

def selectDate = SelectDate()

def write2 = WriteToExcel2(def err = 'Сверка не прошла')

path = 'Объем потерь сверка/Данные со страницы Объем потерь/Объем потерь АО Тываэнерго'

pageDataString = WebUI.getText(findTestObject(path))

println(pageDataString)

fileDataString = obemPoter

println(fileDataString)

check = Check(pageString = pageDataString, fileString = fileDataString, path)

path = 'Объем потерь сверка/Данные со страницы Объем потерь/Уровень потерь АО Тываэнерго'

pageDataString = WebUI.getText(findTestObject(path))

fileDataString = percentPoter

checkPercents = CheckPercents(pageString = pageDataString, fileString = fileDataString, path)

def scanErr = ScannErrors(path)

WebUI.deleteAllCookies()

WebUI.closeBrowser()

'9'
start = OpenBrowser()

obemPoter = findTestData('Test Data').getValue(2, 75)

println(obemPoter)

percentPoter = findTestData('Test Data').getValue(3, 75)

println(percentPoter)

filterDzo = SelectDzo()

WebUI.scrollToElement(findTestObject('Объем потерь сверка/Россети Центр'), 30)

WebUI.scrollToElement(findTestObject('Общие/скролл до фильтра дата'), 30)

WebUI.click(findTestObject('Объем потерь сверка/выбрать Россети Сибирь(ГК)'))

WebUI.click(findTestObject('Общие/Применить в фильтре ДЗО'))

selectDate = SelectDate()

write2 = WriteToExcel2(err = 'Сверка не прошла')

path = 'Объем потерь сверка/Данные со страницы Объем потерь/Объем потерь АО Тываэнерго'

pageDataString = WebUI.getText(findTestObject(path))

println(pageDataString)

fileDataString = obemPoter

println(fileDataString)

check = Check(pageString = pageDataString, fileString = fileDataString, path)

println(check)

path = 'Объем потерь сверка/Данные со страницы Объем потерь/Уровень потерь АО Тываэнерго'

pageDataString = WebUI.getText(findTestObject(path))

fileDataString = percentPoter

checkPercents = CheckPercents(pageString = pageDataString, fileString = fileDataString, path)

println(checkPercents)

def test1 = Test1(check, checkPercents)

scanErr = ScannErrors(path)

WebUI.deleteAllCookies()

WebUI.closeBrowser()

'90'
start = OpenBrowser()

obemPoter = findTestData('Test Data').getValue(2, 71)

println(obemPoter)

percentPoter = findTestData('Test Data').getValue(3, 71)

println(percentPoter)

filterDzo = SelectDzo()

WebUI.scrollToElement(findTestObject('Объем потерь сверка/Россети Центр'), 30)

WebUI.scrollToElement(findTestObject('Общие/скролл до фильтра дата'), 30)

WebUI.click(findTestObject('Объем потерь сверка/выбрать Россети Центр'))

WebUI.click(findTestObject('Общие/Применить в фильтре ДЗО'))

selectDate = SelectDate()

write2 = WriteToExcel2(err = 'Сверка не прошла')

path = 'Объем потерь сверка/Данные со страницы Объем потерь/Объем потерь АО Тываэнерго'

pageDataString = WebUI.getText(findTestObject(path))

println(pageDataString)

fileDataString = obemPoter

println(fileDataString)

check = Check(pageString = pageDataString, fileString = fileDataString, path)

println(check)

path = 'Объем потерь сверка/Данные со страницы Объем потерь/Уровень потерь АО Тываэнерго'

pageDataString = WebUI.getText(findTestObject(path))

fileDataString = percentPoter

checkPercents = CheckPercents(pageString = pageDataString, fileString = fileDataString, path)

println(checkPercents)

if (check == false) {
    println('Start filial')

    WebUI.callTestCase(findTestCase('Объем потерь сверка/Сверка по филиалам/Объем потерь Филиалы Россети Центр'), [:], FailureHandling.CONTINUE_ON_FAILURE)
} else if (checkPercents == false) {
    println('Start filial')

    WebUI.callTestCase(findTestCase('Объем потерь сверка/Сверка по филиалам/Объем потерь Филиалы Россети Центр'), [:], FailureHandling.CONTINUE_ON_FAILURE)
} else {
    println('End case DZO')
}

scanErr = ScannErrors(path)

WebUI.deleteAllCookies()

WebUI.closeBrowser()

static def OpenBrowser() {
    WebUI.openBrowser('')

    WebUI.navigateToUrl(findTestData('Test Data').getValue(7, 3))

    WebUI.setText(findTestObject('Общие/input__username'), findTestData('Test Data').getValue(5, 1))

    WebUI.setText(findTestObject('Общие/input__password'), findTestData('Test Data').getValue(6, 1))

    WebUI.click(findTestObject('Общие/button_'))

    WebUI.delay(10)
}

static def SelectDzo() {
    WebUI.click(findTestObject('Общие/Фильтр ДЗО'))

    WebUI.click(findTestObject('Общие/Снять выделения в фильтре ДЗО'))

    WebUI.click(findTestObject('Общие/Применить в фильтре ДЗО'))

    WebUI.click(findTestObject('Общие/Фильтр ДЗО'))

    WebUI.click(findTestObject('Объем потерь сверка/ПАО Росссети'))

    WebUI.click(findTestObject('Объем потерь сверка/РаспредКомплекс'))
}

static def SelectDate() {
    WebUI.click(findTestObject('Общие/Фильтр Дата'))

    WebUI.click(findTestObject('Общие/Снять выделения в фильтре Дата'))

    WebUI.click(findTestObject('Общие/Применить в фильтре Дата'))

    WebUI.click(findTestObject('Общие/Фильтр Дата'))

    WebUI.scrollToElement(findTestObject('Объем потерь (Данные в виджетах)/2023 год'), 30)

    WebUI.scrollToElement(findTestObject('Объем потерь (Данные в виджетах)/скролл'), 30)

    WebUI.click(findTestObject('Объем потерь (Данные в виджетах)/2023 год'), FailureHandling.CONTINUE_ON_FAILURE)

    WebUI.scrollToElement(findTestObject('Объем потерь (Данные в виджетах)/4 квартал 2023'), 30)

    WebUI.scrollToElement(findTestObject('Объем потерь (Данные в виджетах)/скролл'), 30)

    WebUI.click(findTestObject('Объем потерь сверка/Выбрать 1 квартал 2023'), FailureHandling.CONTINUE_ON_FAILURE)

    WebUI.click(findTestObject('Объем потерь сверка/выбрать 2 квартал 2023'), FailureHandling.CONTINUE_ON_FAILURE)

    WebUI.click(findTestObject('Общие в сеть/Объем потерь сверка/раскрыть 3 квартал 2023'), FailureHandling.CONTINUE_ON_FAILURE)

    WebUI.click(findTestObject('Общие в сеть/Объем потерь сверка/Июль'), FailureHandling.CONTINUE_ON_FAILURE)

    WebUI.scrollToElement(findTestObject('Объем потерь (Данные в виджетах)/Август'), 30)

    WebUI.scrollToElement(findTestObject('Объем потерь (Данные в виджетах)/скролл'), 30)

    WebUI.click(findTestObject('Объем потерь (Данные в виджетах)/Август'), FailureHandling.CONTINUE_ON_FAILURE)

    WebUI.click(findTestObject('Общие/Применить в фильтре Дата'))
}

static def ScannErrors(def path) {
    if (WebUI.verifyTextNotPresent('нет данных', false) == false) {
        def write = WriteToExcel(def file = '', def page = 'нет данных', path)
    } else if (WebUI.verifyTextNotPresent('Ошибка запроса данных', false) == false) {
        def write = WriteToExcel(def file = '', def page = 'Ошибка запроса данных', path)
    } else if (WebUI.verifyTextNotPresent('Произошла ошибка при выполнении пользовательского кода', false) == false) {
        def write = WriteToExcel(def file = '', def page = 'Произошла ошибка при выполнении пользовательского кода', path)
    } else if (WebUI.verifyTextNotPresent('У виджета нет данных', false) == false) {
        def write = WriteToExcel(def file = '', def page = 'У виджета нет данных', path)
    } else if (WebUI.verifyTextNotPresent('Некорректные фильтры', false) == false) {
        def write = WriteToExcel(def file = '', def page = 'Некорректные фильтры', path)
    }
}

static def Check(def pageString, def fileString, def path) {
    pageString = pageString.replaceAll('\\s+', '')

    int page = pageString.toInteger()

    int page1

    int file = fileString.toInteger()

    if (page == file) {
        page1 = page
    } else if (page > file) {
        page1 = (page - 1)
    } else if (page < file) {
        page1 = (page + 1)
    }
    
    if (WebUI.verifyEqual(page1, file) == true) {
        return true
    } else {
        def write = WriteToExcel(file, page, path)

        return false
    }
}

static def CheckPercents(def pageString, def fileString, def path) {
    println(pageString)

    int i = pageString.length()

    println(i)

    if (i > 4) {
        String first = pageString.charAt(0)

        String second = pageString.charAt(1)

        String third = pageString.charAt(2)

        String fourth = pageString.charAt(3)

        String fifth = pageString.charAt(4)

        pageString = ((((first + second) + third) + fourth) + fifth)

        pageString = pageString.trim()

        println(pageString)
    } else {
        pageString = pageString.trim()

        println(pageString)
    }
    
    double page = pageString.toDouble()

    println(page)

    String ii = fileString

    double file = Double.parseDouble(ii.replace(',', '.'))

    println(file)

    if (WebUI.verifyEqual(page, file) == true) {
        return true
    } else {
        def write = WriteToExcel(file, page, path)

        return false
    }
}

static def WriteToExcel(def file, def page, def path) {
    String sheetName = 'List1'

    def data = findTestData('Test Data')

    int n = data.getRowNumbers() + 1

    String dashboardName = 'Объем потерь'

    def workbook01 = ExcelKeywords.getWorkbook(GlobalVariable.excelFilePath)

    def sheet01 = ExcelKeywords.getExcelSheet(workbook01, sheetName)

    println(n)

    println(path)

    path = path.replaceAll('Объем потерь сверка/Данные со страницы Объем потерь/', '')

    String dZO = WebUI.getText(findTestObject('Общие/Фильтр ДЗО'))

    println(dZO)

    Date d = new Date()

    SimpleDateFormat format1

    format1 = new SimpleDateFormat('dd.MM.yyyy')

    String date = format1.format(d)

    println(date)

    String year = WebUI.getText(findTestObject('Объем потерь (Данные в виджетах)/фильтр Дата'))

    ExcelKeywords.setValueToCellByIndex(sheet01, n, 0, dashboardName)

    ExcelKeywords.setValueToCellByIndex(sheet01, n, 1, dZO)

    ExcelKeywords.setValueToCellByIndex(sheet01, n, 2, file)

    ExcelKeywords.setValueToCellByIndex(sheet01, n, 3, page)

    ExcelKeywords.setValueToCellByIndex(sheet01, n, 4, year)

    ExcelKeywords.setValueToCellByIndex(sheet01, n, 5, date)

    n = (n + 1)

    ExcelKeywords.saveWorkbook(GlobalVariable.excelFilePath, workbook01)
}

static def WriteToExcel2(def err) {
    String sheetName = 'List1'

    def data = findTestData('Test Data')

    int n = data.getRowNumbers() + 1

    String dZO = WebUI.getText(findTestObject('Общие/Фильтр ДЗО'))

    String year = WebUI.getText(findTestObject('Объем потерь (Данные в виджетах)/фильтр Дата'))

    println(year)

    String dashboardName = 'Объем потерь'

    def workbook01 = ExcelKeywords.getWorkbook(GlobalVariable.excelFilePath)

    def sheet01 = ExcelKeywords.getExcelSheet(workbook01, sheetName)

    if (WebUI.verifyTextNotPresent('нет данных', false) == false) {
        ExcelKeywords.setValueToCellByIndex(sheet01, n, 0, dashboardName)

        ExcelKeywords.setValueToCellByIndex(sheet01, n, 1, dZO)

        ExcelKeywords.setValueToCellByIndex(sheet01, n, 2, year)

        ExcelKeywords.setValueToCellByIndex(sheet01, n, 3, err)

        n = (n + 1)

        ExcelKeywords.saveWorkbook(GlobalVariable.excelFilePath, workbook01)

        WebUI.closeBrowser()
    } else {
        if (WebUI.verifyTextNotPresent('Ошибка запроса данных', false) == false) {
            ExcelKeywords.setValueToCellByIndex(sheet01, n, 0, dashboardName)

            ExcelKeywords.setValueToCellByIndex(sheet01, n, 1, dZO)

            ExcelKeywords.setValueToCellByIndex(sheet01, n, 2, year)

            ExcelKeywords.setValueToCellByIndex(sheet01, n, 3, err)

            n = (n + 1)

            ExcelKeywords.saveWorkbook(GlobalVariable.excelFilePath, workbook01)
        } else {
            if (WebUI.verifyTextNotPresent('Произошла ошибка при выполнении пользовательского кода', false) == false) {
                ExcelKeywords.setValueToCellByIndex(sheet01, n, 0, dashboardName)

                ExcelKeywords.setValueToCellByIndex(sheet01, n, 1, dZO)

                ExcelKeywords.setValueToCellByIndex(sheet01, n, 2, year)

                ExcelKeywords.setValueToCellByIndex(sheet01, n, 3, err)

                n = (n + 1)

                ExcelKeywords.saveWorkbook(GlobalVariable.excelFilePath, workbook01)
            } else {
                if (WebUI.verifyTextNotPresent('У виджета нет данных', false) == false) {
                    ExcelKeywords.setValueToCellByIndex(sheet01, n, 0, dashboardName)

                    ExcelKeywords.setValueToCellByIndex(sheet01, n, 1, dZO)

                    ExcelKeywords.setValueToCellByIndex(sheet01, n, 2, year)

                    ExcelKeywords.setValueToCellByIndex(sheet01, n, 3, err)

                    n = (n + 1)

                    ExcelKeywords.saveWorkbook(GlobalVariable.excelFilePath, workbook01)
                } else {
                    if (WebUI.verifyTextNotPresent('Некорректные фильтры', false) == false) {
                        ExcelKeywords.setValueToCellByIndex(sheet01, n, 0, dashboardName)

                        ExcelKeywords.setValueToCellByIndex(sheet01, n, 1, dZO)

                        ExcelKeywords.setValueToCellByIndex(sheet01, n, 2, year)

                        ExcelKeywords.setValueToCellByIndex(sheet01, n, 3, err)

                        n = (n + 1)

                        ExcelKeywords.saveWorkbook(GlobalVariable.excelFilePath, workbook01)
                    }
                }
            }
        }
    }
}

static def Raspred() {
    WebUI.click(findTestObject('Общие/Фильтр ДЗО'))

    WebUI.click(findTestObject('Общие/Снять выделения в фильтре ДЗО'))

    WebUI.click(findTestObject('Общие/Применить в фильтре ДЗО'))

    WebUI.click(findTestObject('Общие/Фильтр ДЗО'))

    WebUI.click(findTestObject('Объем потерь сверка/ПАО Росссети'))

    WebUI.click(findTestObject('Объем потерь сверка/Выбрать РаспердКомплекс'))

    WebUI.click(findTestObject('Общие/Применить в фильтре ДЗО'))
}

static def Test1(check, checkPercents) {
	
	if (check == false) {
		println('Start filial')
	
		WebUI.callTestCase(findTestCase('Объем потерь сверка/Сверка по филиалам/Объем потерь Филиалы Россети Сибирь(ГК)'), [:],
			FailureHandling.CONTINUE_ON_FAILURE)
	} else if (checkPercents == false) {
		println('Start filial')
	
		WebUI.callTestCase(findTestCase('Объем потерь сверка/Сверка по филиалам/Объем потерь Филиалы Россети Сибирь(ГК)'), [:],
			FailureHandling.CONTINUE_ON_FAILURE)
	} else {
		println('End case DZO')
	}
	
}
