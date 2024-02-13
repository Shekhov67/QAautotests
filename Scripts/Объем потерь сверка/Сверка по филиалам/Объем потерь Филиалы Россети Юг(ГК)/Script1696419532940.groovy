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

'1'
start = OpenBrowser()

obemPoter = findTestData('Test Data').getValue(3, 52)

println(obemPoter)

percentPoter = findTestData('Test Data').getValue(4, 52)

println(percentPoter)

filterDzo = SelectDzo()

WebUI.scrollToElement(findTestObject('Объем потерь (Данные в виджетах)/Россети Юг(ГК)'), 30)

WebUI.scrollToElement(findTestObject('Объем потерь (Данные в виджетах)/скролл'), 30)

WebUI.click(findTestObject('Объем потерь (Данные в виджетах)/Россети Юг(ГК)'))

WebUI.scrollToElement(findTestObject('Объем потерь (Данные в виджетах)/Астраханьэнерго'), 30)

WebUI.scrollToElement(findTestObject('Объем потерь (Данные в виджетах)/скролл'), 30)

WebUI.click(findTestObject('Объем потерь (Данные в виджетах)/Астраханьэнерго'), FailureHandling.CONTINUE_ON_FAILURE)

WebUI.click(findTestObject('Общие/Применить в фильтре ДЗО'))

WebUI.callTestCase(findTestCase('Объем потерь сверка/Выбор даты для сверки филиалов/Дата для сверки'), [:], FailureHandling.CONTINUE_ON_FAILURE)

scanErr = ScannErrors(path = 'Нет данных')

if (WebUI.verifyTextNotPresent('нет данных', true) == true) {
    print('Сверка с ПБЭ')

    path = 'Объем потерь сверка/Данные со страницы Объем потерь/Объем потерь АО Тываэнерго'

    pageDataString = WebUI.getText(findTestObject(path)).replace('Объем потерь', '').replaceAll('\\s+', '').substring(0, 
        obemPoter.length())

    println(pageDataString)

    fileDataString = obemPoter

    println(fileDataString)

    check = Check(pageString = pageDataString, fileString = fileDataString, path)

    path = 'Объем потерь сверка/Данные со страницы Объем потерь/Уровень потерь АО Тываэнерго'

    pageDataString = WebUI.getText(findTestObject(path)).replace('Уровень потерь', '').replaceAll('\\s+', '').substring(
        0, percentPoter.length())

    fileDataString = percentPoter

    checkPercents = CheckPercents(pageString = pageDataString, fileString = fileDataString, path)
} else {
    print('Данные не отображются')

    WebUI.closeBrowser()
}

WebUI.deleteAllCookies()

WebUI.closeBrowser()

'2'
start = OpenBrowser()

obemPoter = findTestData('Test Data').getValue(3, 53)

println(obemPoter)

percentPoter = findTestData('Test Data').getValue(4, 53)

println(percentPoter)

filterDzo = SelectDzo()

WebUI.scrollToElement(findTestObject('Объем потерь (Данные в виджетах)/Россети Юг(ГК)'), 30)

WebUI.scrollToElement(findTestObject('Объем потерь (Данные в виджетах)/скролл'), 30)

WebUI.click(findTestObject('Объем потерь (Данные в виджетах)/Россети Юг(ГК)'))

WebUI.scrollToElement(findTestObject('Объем потерь (Данные в виджетах)/Волгоградэнерго'), 30)

WebUI.scrollToElement(findTestObject('Объем потерь (Данные в виджетах)/скролл'), 30)

WebUI.click(findTestObject('Объем потерь (Данные в виджетах)/Волгоградэнерго'), FailureHandling.CONTINUE_ON_FAILURE)

WebUI.click(findTestObject('Общие/Применить в фильтре ДЗО'))

WebUI.callTestCase(findTestCase('Объем потерь сверка/Выбор даты для сверки филиалов/Дата для сверки'), [:], FailureHandling.CONTINUE_ON_FAILURE)

scanErr = ScannErrors(path = 'Нет данных')

if (WebUI.verifyTextNotPresent('нет данных', true) == true) {
    print('Сверка с ПБЭ')

    path = 'Объем потерь сверка/Данные со страницы Объем потерь/Объем потерь АО Тываэнерго'

    pageDataString = WebUI.getText(findTestObject(path)).replace('Объем потерь', '').replaceAll('\\s+', '').substring(0, 
        obemPoter.length())

    println(pageDataString)

    fileDataString = obemPoter

    println(fileDataString)

    check = Check(pageString = pageDataString, fileString = fileDataString, path)

    path = 'Объем потерь сверка/Данные со страницы Объем потерь/Уровень потерь АО Тываэнерго'

    pageDataString = WebUI.getText(findTestObject(path)).replace('Уровень потерь', '').replaceAll('\\s+', '').substring(
        0, percentPoter.length())

    fileDataString = percentPoter

    checkPercents = CheckPercents(pageString = pageDataString, fileString = fileDataString, path)
} else {
    print('Данные не отображются')

    WebUI.closeBrowser()
}

WebUI.deleteAllCookies()

WebUI.closeBrowser()

'3'
start = OpenBrowser()

obemPoter = findTestData('Test Data').getValue(3, 54)

println(obemPoter)

percentPoter = findTestData('Test Data').getValue(4, 54)

println(percentPoter)

filterDzo = SelectDzo()

WebUI.scrollToElement(findTestObject('Объем потерь (Данные в виджетах)/Россети Юг(ГК)'), 30)

WebUI.scrollToElement(findTestObject('Объем потерь (Данные в виджетах)/скролл'), 30)

WebUI.click(findTestObject('Объем потерь (Данные в виджетах)/Россети Юг(ГК)'))

WebUI.scrollToElement(findTestObject('Объем потерь (Данные в виджетах)/Калмэнерго'), 30)

WebUI.scrollToElement(findTestObject('Объем потерь (Данные в виджетах)/скролл'), 30)

WebUI.click(findTestObject('Объем потерь (Данные в виджетах)/Калмэнерго'), FailureHandling.CONTINUE_ON_FAILURE)

WebUI.click(findTestObject('Общие/Применить в фильтре ДЗО'))

WebUI.callTestCase(findTestCase('Объем потерь сверка/Выбор даты для сверки филиалов/Дата для сверки'), [:], FailureHandling.CONTINUE_ON_FAILURE)

scanErr = ScannErrors(path = 'Нет данных')

if (WebUI.verifyTextNotPresent('нет данных', true) == true) {
    print('Сверка с ПБЭ')

    path = 'Объем потерь сверка/Данные со страницы Объем потерь/Объем потерь АО Тываэнерго'

    pageDataString = WebUI.getText(findTestObject(path)).replace('Объем потерь', '').replaceAll('\\s+', '').substring(0, 
        obemPoter.length())

    println(pageDataString)

    fileDataString = obemPoter

    println(fileDataString)

    check = Check(pageString = pageDataString, fileString = fileDataString, path)

    path = 'Объем потерь сверка/Данные со страницы Объем потерь/Уровень потерь АО Тываэнерго'

    pageDataString = WebUI.getText(findTestObject(path)).replace('Уровень потерь', '').replaceAll('\\s+', '').substring(
        0, percentPoter.length())

    fileDataString = percentPoter

    checkPercents = CheckPercents(pageString = pageDataString, fileString = fileDataString, path)
} else {
    print('Данные не отображются')

    WebUI.closeBrowser()
}

WebUI.deleteAllCookies()

WebUI.closeBrowser()

'4\r\n'
start = OpenBrowser()

obemPoter = findTestData('Test Data').getValue(3, 55)

println(obemPoter)

percentPoter = findTestData('Test Data').getValue(4, 55)

println(percentPoter)

filterDzo = SelectDzo()

WebUI.scrollToElement(findTestObject('Объем потерь (Данные в виджетах)/Россети Юг(ГК)'), 30)

WebUI.scrollToElement(findTestObject('Объем потерь (Данные в виджетах)/скролл'), 30)

WebUI.click(findTestObject('Объем потерь (Данные в виджетах)/Россети Юг(ГК)'))

WebUI.scrollToElement(findTestObject('Объем потерь (Данные в виджетах)/Ростовэнерго'), 30)

WebUI.scrollToElement(findTestObject('Объем потерь (Данные в виджетах)/скролл'), 30)

WebUI.click(findTestObject('Объем потерь (Данные в виджетах)/Ростовэнерго'), FailureHandling.CONTINUE_ON_FAILURE)

WebUI.click(findTestObject('Общие/Применить в фильтре ДЗО'))

WebUI.callTestCase(findTestCase('Объем потерь сверка/Выбор даты для сверки филиалов/Дата для сверки'), [:], FailureHandling.CONTINUE_ON_FAILURE)

scanErr = ScannErrors(path = 'Нет данных')

if (WebUI.verifyTextNotPresent('нет данных', true) == true) {
    print('Сверка с ПБЭ')

    path = 'Объем потерь сверка/Данные со страницы Объем потерь/Объем потерь АО Тываэнерго'

    pageDataString = WebUI.getText(findTestObject(path)).replace('Объем потерь', '').replaceAll('\\s+', '').substring(0, 
        obemPoter.length())

    println(pageDataString)

    fileDataString = obemPoter

    println(fileDataString)

    check = Check(pageString = pageDataString, fileString = fileDataString, path)

    path = 'Объем потерь сверка/Данные со страницы Объем потерь/Уровень потерь АО Тываэнерго'

    pageDataString = WebUI.getText(findTestObject(path)).replace('Уровень потерь', '').replaceAll('\\s+', '').substring(
        0, percentPoter.length())

    fileDataString = percentPoter

    checkPercents = CheckPercents(pageString = pageDataString, fileString = fileDataString, path)
} else {
    print('Данные не отображются')

    WebUI.closeBrowser()
}

WebUI.deleteAllCookies()

WebUI.closeBrowser()

static def OpenBrowser() {
    WebUI.openBrowser('')

    WebUI.refresh()

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

static def ScannErrors(def path) {
    if (WebUI.verifyTextNotPresent('нет данных', false) == false) {
        def write = WriteToExcel2(def page = 'нет данных', path)
    } else if (WebUI.verifyTextNotPresent('Ошибка запроса данных', false) == false) {
        def write = WriteToExcel2(def page = 'Ошибка запроса данных', path)
    } else if (WebUI.verifyTextNotPresent('Произошла ошибка при выполнении пользовательского кода', false) == false) {
        def write = WriteToExcel2(def page = 'Произошла ошибка при выполнении пользовательского кода', path)
    } else if (WebUI.verifyTextNotPresent('У виджета нет данных', false) == false) {
        def write = WriteToExcel2(def page = 'У виджета нет данных', path)
    } else if (WebUI.verifyTextNotPresent('Некорректные фильтры', false) == false) {
        def write = WriteToExcel2(def page = 'Некорректные фильтры', path)
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
    } else {
        def write = WriteToExcel(file, page, path)
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
    } else {
        def write = WriteToExcel(file, page, path)
    }
}

static def WriteToExcel(def file, def page, def path) {
    String sheetName = 'List1'

    def data = findTestData('Test Data')

    int n = data.getRowNumbers() + 1

    String dashboardName = 'Объем потерь'

    def workbook01 = ExcelKeywords.getWorkbook(GlobalVariable.excelFilePathFilials)

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

    ExcelKeywords.saveWorkbook(GlobalVariable.excelFilePathFilials, workbook01)
}

static def WriteToExcel2(def err, def page) {
    String sheetName = 'List1'

    def data = findTestData('Test Data')

    int n = data.getRowNumbers() + 1

    String dZO = WebUI.getText(findTestObject('Общие/Фильтр ДЗО'))

    String year = WebUI.getText(findTestObject('Объем потерь (Данные в виджетах)/фильтр Дата'))

    println(year)

    String dashboardName = 'Объем потерь'

    def workbook01 = ExcelKeywords.getWorkbook(GlobalVariable.excelFilePathFilials)

    def sheet01 = ExcelKeywords.getExcelSheet(workbook01, sheetName)

    if (WebUI.verifyTextNotPresent('нет данных', false) == false) {
        ExcelKeywords.setValueToCellByIndex(sheet01, n, 0, dashboardName)

        ExcelKeywords.setValueToCellByIndex(sheet01, n, 1, dZO)

        ExcelKeywords.setValueToCellByIndex(sheet01, n, 2, year)

        ExcelKeywords.setValueToCellByIndex(sheet01, n, 3, err)

        ExcelKeywords.setValueToCellByIndex(sheet01, n, 4, page)

        n = (n + 1)

        ExcelKeywords.saveWorkbook(GlobalVariable.excelFilePathFilials, workbook01)

        WebUI.closeBrowser()
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

