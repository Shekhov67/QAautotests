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

def test0 = Test0()

def test1 = Test1()

def test2 = Test2()

def test3 = Test3()

def test4 = Test4()

def test5 = Test5()

def test6 = Test6()

def test7 = Test7()

def test8 = Test8()

def test9 = Test9()

def test10 = Test10()

def test11 = Test11()

def test12 = Test12()

def test13 = Test13()

def test14 = Test14()

def test15 = Test15()

def test16 = Test16()

def test17 = Test17()

static def OpenBrowser() {
    WebUI.openBrowser('')

    WebUI.navigateToUrl(findTestData('Test Data').getValue(7, 3))

    WebUI.delay(5)

    if (WebUI.verifyElementText(findTestObject('Общие/button_'), 'Вход') == true) {
        WebUI.setText(findTestObject('Общие/input__username'), findTestData('Test Data').getValue(5, 1))

        WebUI.setText(findTestObject('Общие/input__password'), findTestData('Test Data').getValue(6, 1))

        WebUI.click(findTestObject('Общие в сеть/button_'))

        WebUI.delay(5)
    } else {
        WebUI.refresh()

        WebUI.delay(5)

        WebUI.setText(findTestObject('Общие/input__username'), findTestData('Test Data').getValue(5, 1))

        WebUI.setText(findTestObject('Общие/input__password'), findTestData('Test Data').getValue(6, 1))

        WebUI.click(findTestObject('Общие в сеть/button_'))
    }
}

static def SelectDzo() {
    WebUI.click(findTestObject('Объем потерь (Данные в виджетах)/фильтр ДЗО'))

    WebUI.click(findTestObject('Объем потерь (Данные в виджетах)/Снять выделения в фильтре ДЗО'))

    WebUI.click(findTestObject('Объем потерь (Данные в виджетах)/Применить в фильтре ДЗО'))

    WebUI.click(findTestObject('Объем потерь (Данные в виджетах)/фильтр ДЗО'))

    WebUI.click(findTestObject('Объем потерь сверка/ПАО Росссети'))

    WebUI.click(findTestObject('Объем потерь сверка/РаспредКомплекс'))
}

static def SelectDate() {
    WebUI.click(findTestObject('Объем потерь (Данные в виджетах)/i_close'))

    WebUI.click(findTestObject('Объем потерь (Данные в виджетах)/фильтр Дата'))

    WebUI.scrollToElement(findTestObject('Объем потерь сверка/2024 год'), 30)

    WebUI.scrollToElement(findTestObject('Объем потерь (Данные в виджетах)/скролл'), 30)

    WebUI.click(findTestObject('Объем потерь сверка/2024 год'), FailureHandling.CONTINUE_ON_FAILURE)

    WebUI.scrollToElement(findTestObject('Объем потерь сверка/1 квартал'), 30)

    WebUI.scrollToElement(findTestObject('Объем потерь (Данные в виджетах)/скролл'), 30)

    WebUI.click(findTestObject('Объем потерь сверка/1 квартал'), FailureHandling.CONTINUE_ON_FAILURE)

    WebUI.scrollToElement(findTestObject('Объем потерь сверка/Январь'), 30)

    WebUI.scrollToElement(findTestObject('Объем потерь (Данные в виджетах)/скролл'), 30)

    WebUI.click(findTestObject('Объем потерь сверка/Январь'), FailureHandling.CONTINUE_ON_FAILURE)

    WebUI.click(findTestObject('Объем потерь сверка/Февраль'), FailureHandling.CONTINUE_ON_FAILURE)

    WebUI.click(findTestObject('Объем потерь (Данные в виджетах)/Март'), FailureHandling.CONTINUE_ON_FAILURE)

    WebUI.click(findTestObject('Объем потерь (Данные в виджетах)/Применить в фильтре Дата'))
}

static def ScannErrors(def path) {
    if (WebUI.verifyTextNotPresent('нет данных', false) == false) {
        def write = WriteToExcel2(def err = 'нет данных')
    } else if (WebUI.verifyTextNotPresent('Ошибка запроса данных', false) == false) {
        def write = WriteToExcel2(def err = 'Ошибка запроса данных')
    } else if (WebUI.verifyTextNotPresent('Произошла ошибка при выполнении пользовательского кода', false) == false) {
        def write = WriteToExcel2(def err = 'Произошла ошибка при выполнении пользовательского кода')
    } else if (WebUI.verifyTextNotPresent('У виджета нет данных', false) == false) {
        def write = WriteToExcel2(def err = 'У виджета нет данных')
    } else if (WebUI.verifyTextNotPresent('Некорректные фильтры', false) == false) {
        def write = WriteToExcel2(def err = 'Некорректные фильтры')
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
        def write = WriteToExcel(file, page, path, def typeDate = 'Объем потерь')

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
        def write = WriteToExcel(file, page, path, def typeDate = 'Уровень потерь')

        return false
    }
}

static def WriteToExcel(def file, def page, def path, def typeDate) {
    String sheetName = 'List1'

    def data = findTestData('Test Data')

    int n = data.getRowNumbers() + 1

    String dashboardName = 'Объем потерь'

    def workbook01 = ExcelKeywords.getWorkbook(GlobalVariable.excelFilePath)

    def sheet01 = ExcelKeywords.getExcelSheet(workbook01, sheetName)

    println(n)

    println(path)

    path = path.replaceAll('Объем потерь сверка/Данные со страницы Объем потерь/', '')

    String dZO = WebUI.getText(findTestObject('Объем потерь (Данные в виджетах)/фильтр ДЗО'))

    println(dZO)

    println(typeDate)

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

    ExcelKeywords.setValueToCellByIndex(sheet01, n, 6, typeDate)

    n = (n + 1)

    ExcelKeywords.saveWorkbook(GlobalVariable.excelFilePath, workbook01)
}

static def WriteToExcel2(def err) {
    String sheetName = 'List1'

    def data = findTestData('Test Data')

    int n = data.getRowNumbers() + 1

    String dZO = WebUI.getText(findTestObject('Объем потерь (Данные в виджетах)/фильтр ДЗО'))

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
    }
}

static def Raspred() {
    WebUI.click(findTestObject('Объем потерь (Данные в виджетах)/фильтр ДЗО'))

    WebUI.click(findTestObject('Объем потерь (Данные в виджетах)/Снять выделения в фильтре ДЗО'))

    WebUI.click(findTestObject('Объем потерь (Данные в виджетах)/Применить в фильтре ДЗО'))

    WebUI.click(findTestObject('Объем потерь (Данные в виджетах)/фильтр ДЗО'))

    WebUI.click(findTestObject('Объем потерь сверка/ПАО Росссети'))

    WebUI.click(findTestObject('Объем потерь сверка/Выбрать РаспердКомплекс'))

    WebUI.click(findTestObject('Объем потерь (Данные в виджетах)/Применить в фильтре ДЗО'))
}

static def Test0() {
    '0'
    def start = OpenBrowser()

    String obemPoter = findTestData('Test Data').getValue(2, 87)

    println(obemPoter)

    String percentPoter = findTestData('Test Data').getValue(3, 87)

    println(percentPoter)

    def raspred = Raspred()

    def selectDate = SelectDate()

    def write2 = WriteToExcel2(def err = 'Сверка не прошла')

    def path = 'Объем потерь сверка/Данные со страницы Объем потерь/Объем потерь АО Тываэнерго'

    def pageDataString = WebUI.getText(findTestObject(path))

    println(pageDataString)

    if (pageDataString == null) {
        WebUI.closeBrowser()

        return 
    } else {
        pageDataString = WebUI.getText(findTestObject(path)).replace('Объем потерь', '').replaceAll('\\s+', '').substring(
            0, obemPoter.length())

        def fileDataString = obemPoter

        println(fileDataString)

        def check = Check(def pageString = pageDataString, def fileString = fileDataString, path)

        path = 'Объем потерь сверка/Данные со страницы Объем потерь/Уровень потерь АО Тываэнерго'

        pageDataString = WebUI.getText(findTestObject(path)).replace('Уровень потерь', '').replaceAll('\\s+', '').substring(
            0, percentPoter.length())

        fileDataString = percentPoter

        def checkPercents = CheckPercents(pageString = pageDataString, fileString = fileDataString, path)

        def scanErr = ScannErrors(path)

        WebUI.deleteAllCookies()

        WebUI.closeBrowser()
    }
    
    WebUI.closeBrowser()
}

static def Test1() {
    '1'
    def start = OpenBrowser()

    def obemPoter = findTestData('Test Data').getValue(2, 76)

    println(obemPoter)

    def percentPoter = findTestData('Test Data').getValue(3, 76)

    println(percentPoter)

    def filterDzo = SelectDzo()

    WebUI.click(findTestObject('Объем потерь сверка/выбрать АО Тываэнерго'))

    WebUI.click(findTestObject('Объем потерь (Данные в виджетах)/Применить в фильтре ДЗО'))

    def selectDate = SelectDate()

    def write2 = WriteToExcel2(def err = 'Сверка не прошла')

    def path = 'Объем потерь сверка/Данные со страницы Объем потерь/Объем потерь АО Тываэнерго'

    def pageDataString = WebUI.getText(findTestObject(path))

    println(pageDataString)

    if (pageDataString == null) {
        WebUI.closeBrowser()

        return 
    } else {
        pageDataString = WebUI.getText(findTestObject(path)).replace('Объем потерь', '').replaceAll('\\s+', '').substring(
            0, obemPoter.length())

        def fileDataString = obemPoter

        println(fileDataString)

        def check = Check(def pageString = pageDataString, def fileString = fileDataString, path)

        path = 'Объем потерь сверка/Данные со страницы Объем потерь/Уровень потерь АО Тываэнерго'

        pageDataString = WebUI.getText(findTestObject(path)).replace('Уровень потерь', '').replaceAll('\\s+', '').substring(
            0, percentPoter.length())

        fileDataString = percentPoter

        def checkPercents = CheckPercents(pageString = pageDataString, fileString = fileDataString, path)

        def scanErr = ScannErrors(path)

        WebUI.deleteAllCookies()

        WebUI.closeBrowser()
    }
}

static def Test2() {
    '2'
    def start = OpenBrowser()

    def obemPoter = findTestData('Test Data').getValue(2, 81)

    println(obemPoter)

    def percentPoter = findTestData('Test Data').getValue(3, 81)

    println(percentPoter)

    def filterDzo = SelectDzo()

    WebUI.click(findTestObject('Объем потерь сверка/выбрать АО Чеченэнерго'))

    WebUI.click(findTestObject('Объем потерь (Данные в виджетах)/Применить в фильтре ДЗО'))

    def selectDate = SelectDate()

    def write2 = WriteToExcel2(def err = 'Сверка не прошла')

    def path = 'Объем потерь сверка/Данные со страницы Объем потерь/Объем потерь АО Тываэнерго'

    def pageDataString = WebUI.getText(findTestObject(path))

    println(pageDataString)

    if (pageDataString == null) {
        WebUI.closeBrowser()

        return 
    } else {
        pageDataString = WebUI.getText(findTestObject(path)).replace('Объем потерь', '').replaceAll('\\s+', '').substring(
            0, obemPoter.length())

        def fileDataString = obemPoter

        println(fileDataString)

        def check = Check(def pageString = pageDataString, def fileString = fileDataString, path)

        path = 'Объем потерь сверка/Данные со страницы Объем потерь/Уровень потерь АО Тываэнерго'

        pageDataString = WebUI.getText(findTestObject(path)).replace('Уровень потерь', '').replaceAll('\\s+', '').substring(
            0, percentPoter.length())

        fileDataString = percentPoter

        def checkPercents = CheckPercents(pageString = pageDataString, fileString = fileDataString, path)

        def scanErr = ScannErrors(path)

        WebUI.deleteAllCookies()

        WebUI.closeBrowser()
    }
}

static def Test3() {
    '3'
    def start = OpenBrowser()

    def obemPoter = findTestData('Test Data').getValue(2, 73)

    println(obemPoter)

    def percentPoter = findTestData('Test Data').getValue(3, 73)

    println(percentPoter)

    def filterDzo = SelectDzo()

    WebUI.click(findTestObject('Объем потерь сверка/выбрать Россети Волга'))

    WebUI.click(findTestObject('Объем потерь (Данные в виджетах)/Применить в фильтре ДЗО'))

    def selectDate = SelectDate()

    def write2 = WriteToExcel2(def err = 'Сверка не прошла')

    def path = 'Объем потерь сверка/Данные со страницы Объем потерь/Объем потерь АО Тываэнерго'

    def pageDataString = WebUI.getText(findTestObject(path))

    println(pageDataString)

    if (pageDataString == null) {
        WebUI.closeBrowser()

        return 
    } else {
        pageDataString = WebUI.getText(findTestObject(path)).replace('Объем потерь', '').replaceAll('\\s+', '').substring(
            0, obemPoter.length())

        def fileDataString = obemPoter

        println(fileDataString)

        def check = Check(def pageString = pageDataString, def fileString = fileDataString, path)

        path = 'Объем потерь сверка/Данные со страницы Объем потерь/Уровень потерь АО Тываэнерго'

        pageDataString = WebUI.getText(findTestObject(path))

        pageDataString = WebUI.getText(findTestObject(path)).replace('Уровень потерь', '').replaceAll('\\s+', '').substring(
            0, percentPoter.length())

        fileDataString = percentPoter

        def checkPercents = CheckPercents(pageString = pageDataString, fileString = fileDataString, path)

        if (check == false) {
            println('Start filial')

            WebUI.callTestCase(findTestCase('Объем потерь сверка/Сверка по филиалам/Объем потерь Филиалы Россети Волга'), 
                [:], FailureHandling.CONTINUE_ON_FAILURE)
        } else if (checkPercents == false) {
            println('Start filial')

            WebUI.callTestCase(findTestCase('Объем потерь сверка/Сверка по филиалам/Объем потерь Филиалы Россети Волга'), 
                [:], FailureHandling.CONTINUE_ON_FAILURE)
        } else {
            println('End case DZO')

            def scanErr = ScannErrors(path)

            WebUI.deleteAllCookies()

            WebUI.closeBrowser()
        }
    }
}

static def Test4() {
    '4'
    def start = OpenBrowser()

    def obemPoter = findTestData('Test Data').getValue(2, 82)

    println(obemPoter)

    def percentPoter = findTestData('Test Data').getValue(3, 82)

    println(percentPoter)

    def filterDzo = SelectDzo()

    WebUI.scrollToElement(findTestObject('Объем потерь (Данные в виджетах)/Россети Кубань'), 30)

    WebUI.scrollToElement(findTestObject('Объем потерь (Данные в виджетах)/скролл'), 30)

    WebUI.click(findTestObject('Объем потерь сверка/выбрать Россети Кубань'))

    WebUI.click(findTestObject('Объем потерь (Данные в виджетах)/Применить в фильтре ДЗО'))

    def selectDate = SelectDate()

    def write2 = WriteToExcel2(def err = 'Сверка не прошла')

    def path = 'Объем потерь сверка/Данные со страницы Объем потерь/Объем потерь АО Тываэнерго'

    def pageDataString = WebUI.getText(findTestObject(path))

    println(pageDataString)

    if (pageDataString == null) {
        WebUI.closeBrowser()

        return 
    } else {
        pageDataString = WebUI.getText(findTestObject(path)).replace('Объем потерь', '').replaceAll('\\s+', '').substring(
            0, obemPoter.length())

        println(pageDataString)

        def fileDataString = obemPoter

        println(fileDataString)

        def check = Check(def pageString = pageDataString, def fileString = fileDataString, path)

        path = 'Объем потерь сверка/Данные со страницы Объем потерь/Уровень потерь АО Тываэнерго'

        pageDataString = WebUI.getText(findTestObject(path)).replace('Уровень потерь', '').replaceAll('\\s+', '').substring(
            0, percentPoter.length())

        fileDataString = percentPoter

        def checkPercents = CheckPercents(pageString = pageDataString, fileString = fileDataString, path)

        def scanErr = ScannErrors(path)

        WebUI.deleteAllCookies()

        WebUI.closeBrowser()
    }
}

static def Test5() {
    '5'
    def start = OpenBrowser()

    def obemPoter = findTestData('Test Data').getValue(2, 84)

    println(obemPoter)

    def percentPoter = findTestData('Test Data').getValue(3, 84)

    println(percentPoter)

    def filterDzo = SelectDzo()

    WebUI.scrollToElement(findTestObject('Объем потерь (Данные в виджетах)/Россети Ленэнерго(ГК)'), 30)

    WebUI.click(findTestObject('Объем потерь сверка/выбрать Росссети Ленэнерго(ГК)'))

    WebUI.scrollToElement(findTestObject('Объем потерь (Данные в виджетах)/скролл'), 30)

    WebUI.click(findTestObject('Объем потерь (Данные в виджетах)/Применить в фильтре ДЗО'))

    def selectDate = SelectDate()

    def write2 = WriteToExcel2(def err = 'Сверка не прошла')

    def path = 'Объем потерь сверка/Данные со страницы Объем потерь/Объем потерь АО Тываэнерго'

    def pageDataString = WebUI.getText(findTestObject(path))

    println(pageDataString)

    if (pageDataString == null) {
        WebUI.closeBrowser()

        return 
    } else {
        pageDataString = WebUI.getText(findTestObject(path)).replace('Объем потерь', '').replaceAll('\\s+', '').substring(
            0, obemPoter.length())

        println(pageDataString)

        def fileDataString = obemPoter

        println(fileDataString)

        def check = Check(def pageString = pageDataString, def fileString = fileDataString, path)

        path = 'Объем потерь сверка/Данные со страницы Объем потерь/Уровень потерь АО Тываэнерго'

        pageDataString = WebUI.getText(findTestObject(path)).replace('Уровень потерь', '').replaceAll('\\s+', '').substring(
            0, percentPoter.length())

        fileDataString = percentPoter

        def checkPercents = CheckPercents(pageString = pageDataString, fileString = fileDataString, path)

        def scanErr = ScannErrors(path)

        WebUI.deleteAllCookies()

        WebUI.closeBrowser()
    }
}

static def Test6() {
    '6'
    def start = OpenBrowser()

    def obemPoter = findTestData('Test Data').getValue(2, 83)

    println(obemPoter)

    def percentPoter = findTestData('Test Data').getValue(3, 83)

    println('Эксель файл' + percentPoter)

    def filterDzo = SelectDzo()

    WebUI.scrollToElement(findTestObject('Объем потерь (Данные в виджетах)/выбрать Россети Московский регион'), 30)

    WebUI.scrollToElement(findTestObject('Объем потерь (Данные в виджетах)/скролл'), 30)

    WebUI.click(findTestObject('Объем потерь сверка/выбрать Россети Московский регион'))

    WebUI.click(findTestObject('Объем потерь (Данные в виджетах)/Применить в фильтре ДЗО'))

    def selectDate = SelectDate()

    def write2 = WriteToExcel2(def err = 'Сверка не прошла')

    def path = 'Объем потерь сверка/Данные со страницы Объем потерь/Объем потерь АО Тываэнерго'

    def pageDataString = WebUI.getText(findTestObject(path))

    println(pageDataString)

    if (pageDataString == null) {
        WebUI.closeBrowser()

        return 
    } else {
        pageDataString = WebUI.getText(findTestObject(path)).replace('Объем потерь', '').replaceAll('\\s+', '').substring(
            0, obemPoter.length())

        def fileDataString = obemPoter

        println(fileDataString)

        def check = Check(def pageString = pageDataString, def fileString = fileDataString, path)

        path = 'Объем потерь сверка/Данные со страницы Объем потерь/Уровень потерь АО Тываэнерго'

        pageDataString = WebUI.getText(findTestObject(path)).replace('Уровень потерь', '').replaceAll('\\s+', '').substring(
            0, percentPoter.length())

        fileDataString = percentPoter

        def checkPercents = CheckPercents(pageString = pageDataString, fileString = fileDataString, path)

        if (check == false) {
            println('Start filial')

            WebUI.callTestCase(findTestCase('Объем потерь сверка/Сверка по филиалам/Объем потерь Филиалы Россети Московский регион'), 
                [:], FailureHandling.CONTINUE_ON_FAILURE)
        } else if (checkPercents == false) {
            println('Start filial')

            WebUI.callTestCase(findTestCase('Объем потерь сверка/Сверка по филиалам/Объем потерь Филиалы Россети Московский регион'), 
                [:], FailureHandling.CONTINUE_ON_FAILURE)
        } else {
            println('End case DZO')

            def scanErr = ScannErrors(path)

            WebUI.deleteAllCookies()

            WebUI.closeBrowser()
        }
    }
}

static def Test7() {
    '7'
    def start = OpenBrowser()

    def obemPoter = findTestData('Test Data').getValue(2, 80)

    println(obemPoter)

    def percentPoter = findTestData('Test Data').getValue(3, 80)

    println(percentPoter)

    def filterDzo = SelectDzo()

    WebUI.scrollToElement(findTestObject('Объем потерь (Данные в виджетах)/выбрать Россети Северный Кавказ(ГК)'), 30)

    WebUI.scrollToElement(findTestObject('Объем потерь (Данные в виджетах)/скролл'), 30)

    WebUI.click(findTestObject('Объем потерь сверка/выбрать Россети Северный Кавказ(ГК)'))

    WebUI.click(findTestObject('Объем потерь (Данные в виджетах)/Применить в фильтре ДЗО'))

    def selectDate = SelectDate()

    def write2 = WriteToExcel2(def err = 'Сверка не прошла')

    def path = 'Объем потерь сверка/Данные со страницы Объем потерь/Объем потерь АО Тываэнерго'

    def pageDataString = WebUI.getText(findTestObject(path))

    println(pageDataString)

    if (pageDataString == null) {
        return 
        
        WebUI.closeBrowser()
    } else {
        pageDataString = WebUI.getText(findTestObject(path)).replace('Объем потерь', '').replaceAll('\\s+', '').substring(
            0, obemPoter.length())

        def fileDataString = obemPoter

        println(fileDataString)

        def check = Check(def pageString = pageDataString, def fileString = fileDataString, path)

        path = 'Объем потерь сверка/Данные со страницы Объем потерь/Уровень потерь АО Тываэнерго'

        pageDataString = WebUI.getText(findTestObject(path))

        println(pageDataString)

        pageDataString = WebUI.getText(findTestObject(path)).replace('Уровень потерь', '').replaceAll('\\s+', '').substring(
            0, percentPoter.length())

        fileDataString = percentPoter

        def checkPercents = CheckPercents(pageString = pageDataString, fileString = fileDataString, path)

        if (check == false) {
            println('Start filial')

            WebUI.callTestCase(findTestCase('Объем потерь сверка/Сверка по филиалам/Объем потерь Филиалы Россети Северный Кавказ'), 
                [:], FailureHandling.CONTINUE_ON_FAILURE)
        } else if (checkPercents == false) {
            println('Start filial')

            WebUI.callTestCase(findTestCase('Объем потерь сверка/Сверка по филиалам/Объем потерь Филиалы Россети Северный Кавказ'), 
                [:], FailureHandling.CONTINUE_ON_FAILURE)
        } else {
            println('End case DZO')

            def scanErr = ScannErrors(path)

            WebUI.deleteAllCookies()

            WebUI.closeBrowser()
        }
    }
}

static def Test8() {
    '8'
    def start = OpenBrowser()

    def obemPoter = findTestData('Test Data').getValue(2, 74)

    println(obemPoter)

    def percentPoter = findTestData('Test Data').getValue(3, 74)

    println(percentPoter)

    def filterDzo = SelectDzo()

    WebUI.scrollToElement(findTestObject('Объем потерь сверка/Россети Северо-Запад'), 30)

    WebUI.scrollToElement(findTestObject('Объем потерь (Данные в виджетах)/скролл'), 30)

    WebUI.click(findTestObject('Объем потерь сверка/выбрать Россети Северо-Запад'))

    WebUI.click(findTestObject('Объем потерь (Данные в виджетах)/Применить в фильтре ДЗО'))

    def selectDate = SelectDate()

    def write2 = WriteToExcel2(def err = 'Сверка не прошла')

    def path = 'Объем потерь сверка/Данные со страницы Объем потерь/Объем потерь АО Тываэнерго'

    def pageDataString = WebUI.getText(findTestObject(path))

    println(pageDataString)

    if (pageDataString == null) {
        WebUI.closeBrowser()

        return 
    } else {
        pageDataString = WebUI.getText(findTestObject(path)).replace('Объем потерь', '').replaceAll('\\s+', '').substring(
            0, obemPoter.length())

        def fileDataString = obemPoter

        println(fileDataString)

        def check = Check(def pageString = pageDataString, def fileString = fileDataString, path)

        path = 'Объем потерь сверка/Данные со страницы Объем потерь/Уровень потерь АО Тываэнерго'

        pageDataString = WebUI.getText(findTestObject(path))

        println(pageDataString)

        pageDataString = WebUI.getText(findTestObject(path)).replace('Уровень потерь', '').replaceAll('\\s+', '').substring(
            0, percentPoter.length())

        fileDataString = percentPoter

        def checkPercents = CheckPercents(pageString = pageDataString, fileString = fileDataString, path)

        if (check == false) {
            println('Start filial')

            WebUI.callTestCase(findTestCase('Объем потерь сверка/Сверка по филиалам/Объем потерь Филиалы Россети Северо-Запад'), 
                [:], FailureHandling.CONTINUE_ON_FAILURE)
        } else if (checkPercents == false) {
            println('Start filial')

            WebUI.callTestCase(findTestCase('Объем потерь сверка/Сверка по филиалам/Объем потерь Филиалы Россети Северо-Запад'), 
                [:], FailureHandling.CONTINUE_ON_FAILURE)
        } else {
            println('End case DZO')

            def scanErr = ScannErrors(path)

            WebUI.deleteAllCookies()

            WebUI.closeBrowser()
        }
    }
}

static def Test9() {
    '9'
    def start = OpenBrowser()

    def obemPoter = findTestData('Test Data').getValue(2, 75)

    println(obemPoter)

    def percentPoter = findTestData('Test Data').getValue(3, 75)

    println(percentPoter)

    def filterDzo = SelectDzo()

    WebUI.scrollToElement(findTestObject('Объем потерь (Данные в виджетах)/Россети Сибирь (ГК)'), 30)

    WebUI.scrollToElement(findTestObject('Объем потерь (Данные в виджетах)/скролл'), 30)

    WebUI.click(findTestObject('Объем потерь сверка/выбрать Россети Сибирь(ГК)'))

    WebUI.click(findTestObject('Объем потерь (Данные в виджетах)/Применить в фильтре ДЗО'))

    def selectDate = SelectDate()

    def write2 = WriteToExcel2(def err = 'Сверка не прошла')

    def path = 'Объем потерь сверка/Данные со страницы Объем потерь/Объем потерь АО Тываэнерго'

    def pageDataString = WebUI.getText(findTestObject(path))

    println(pageDataString)

    if (pageDataString == null) {
        WebUI.closeBrowser()

        return 
    } else {
        pageDataString = WebUI.getText(findTestObject(path)).replace('Объем потерь', '').replaceAll('\\s+', '').substring(
            0, obemPoter.length())

        def fileDataString = obemPoter

        println(fileDataString)

        def check = Check(def pageString = pageDataString, def fileString = fileDataString, path)

        path = 'Объем потерь сверка/Данные со страницы Объем потерь/Уровень потерь АО Тываэнерго'

        pageDataString = WebUI.getText(findTestObject(path)).replace('Уровень потерь', '').replaceAll('\\s+', '').substring(
            0, percentPoter.length())

        fileDataString = percentPoter

        def checkPercents = CheckPercents(pageString = pageDataString, fileString = fileDataString, path)

        if (check == false) {
            println('Start filial')

            WebUI.callTestCase(findTestCase('Объем потерь сверка/Сверка по филиалам/Объем потерь Филиалы Россети Сибирь(ГК)'), 
                [:], FailureHandling.CONTINUE_ON_FAILURE)
        } else if (checkPercents == false) {
            println('Start filial')

            WebUI.callTestCase(findTestCase('Объем потерь сверка/Сверка по филиалам/Объем потерь Филиалы Россети Сибирь(ГК)'), 
                [:], FailureHandling.CONTINUE_ON_FAILURE)
        } else {
            println('End case DZO')

            def scanErr = ScannErrors(path)

            WebUI.deleteAllCookies()

            WebUI.closeBrowser()
        }
    }
}

static def Test10() {
    '10'
    def start = OpenBrowser()

    def obemPoter = findTestData('Test Data').getValue(2, 77)

    println(obemPoter)

    def percentPoter = findTestData('Test Data').getValue(3, 77)

    println(percentPoter)

    def filterDzo = SelectDzo()

    WebUI.scrollToElement(findTestObject('Объем потерь сверка/Россети Центр'), 30)

    WebUI.scrollToElement(findTestObject('Объем потерь (Данные в виджетах)/скролл'), 30)

    WebUI.click(findTestObject('Объем потерь сверка/выбрать Россети Томск'))

    WebUI.click(findTestObject('Объем потерь (Данные в виджетах)/Применить в фильтре ДЗО'))

    def selectDate = SelectDate()

    def write2 = WriteToExcel2(def err = 'Сверка не прошла')

    def path = 'Объем потерь сверка/Данные со страницы Объем потерь/Объем потерь АО Тываэнерго'

    def pageDataString = WebUI.getText(findTestObject(path))

    println(pageDataString)

    if (pageDataString == null) {
        WebUI.closeBrowser()

        return 
    } else {
        pageDataString = WebUI.getText(findTestObject(path)).replace('Объем потерь', '').replaceAll('\\s+', '').substring(
            0, obemPoter.length())

        def fileDataString = obemPoter

        println(fileDataString)

        def check = Check(def pageString = pageDataString, def fileString = fileDataString, path)

        path = 'Объем потерь сверка/Данные со страницы Объем потерь/Уровень потерь АО Тываэнерго'

        pageDataString = WebUI.getText(findTestObject(path)).replace('Уровень потерь', '').replaceAll('\\s+', '').substring(
            0, percentPoter.length())

        fileDataString = percentPoter

        def checkPercents = CheckPercents(pageString = pageDataString, fileString = fileDataString, path)

        def scanErr = ScannErrors(path)

        WebUI.deleteAllCookies()

        WebUI.closeBrowser()
    }
}

static def Test11() {
    '11'
    def start = OpenBrowser()

    def obemPoter = findTestData('Test Data').getValue(2, 85)

    println(obemPoter)

    def percentPoter = findTestData('Test Data').getValue(3, 85)

    println(percentPoter)

    def filterDzo = SelectDzo()

    WebUI.scrollToElement(findTestObject('Объем потерь сверка/Россети Центр'), 30)

    WebUI.scrollToElement(findTestObject('Объем потерь (Данные в виджетах)/скролл'), 30)

    WebUI.click(findTestObject('Объем потерь сверка/выбрать Тюмень'))

    WebUI.click(findTestObject('Объем потерь (Данные в виджетах)/Применить в фильтре ДЗО'))

    def selectDate = SelectDate()

    def write2 = WriteToExcel2(def err = 'Сверка не прошла')

    def path = 'Объем потерь сверка/Данные со страницы Объем потерь/Объем потерь АО Тываэнерго'

    def pageDataString = WebUI.getText(findTestObject(path))

    println(pageDataString)

    if (pageDataString == null) {
        WebUI.closeBrowser()

        return 
    } else {
        pageDataString = WebUI.getText(findTestObject(path)).replace('Объем потерь', '').replaceAll('\\s+', '').substring(
            0, obemPoter.length())

        def fileDataString = obemPoter

        println(fileDataString)

        def check = Check(def pageString = pageDataString, def fileString = fileDataString, path)

        path = 'Объем потерь сверка/Данные со страницы Объем потерь/Уровень потерь АО Тываэнерго'

        pageDataString = WebUI.getText(findTestObject(path)).replace('Уровень потерь', '').replaceAll('\\s+', '').substring(
            0, percentPoter.length())

        fileDataString = percentPoter

        def checkPercents = CheckPercents(pageString = pageDataString, fileString = fileDataString, path)

        def scanErr = ScannErrors(path)

        WebUI.deleteAllCookies()

        WebUI.closeBrowser()
    }
}

static def Test12() {
    '12'
    def start = OpenBrowser()

    def obemPoter = findTestData('Test Data').getValue(2, 78)

    println(obemPoter)

    def percentPoter = findTestData('Test Data').getValue(3, 78)

    println(percentPoter)

    def filterDzo = SelectDzo()

    WebUI.scrollToElement(findTestObject('Объем потерь сверка/Россети Центр'), 30)

    WebUI.scrollToElement(findTestObject('Объем потерь (Данные в виджетах)/скролл'), 30)

    WebUI.click(findTestObject('Объем потерь сверка/выбрать Россети Урал(ГК)'))

    WebUI.click(findTestObject('Объем потерь (Данные в виджетах)/Применить в фильтре ДЗО'))

    def selectDate = SelectDate()

    def write2 = WriteToExcel2(def err = 'Сверка не прошла')

    def path = 'Объем потерь сверка/Данные со страницы Объем потерь/Объем потерь АО Тываэнерго'

    def pageDataString = WebUI.getText(findTestObject(path)).replace('Объем потерь', '').replaceAll('\\s+', '').substring(
        0, obemPoter.length())

    println(pageDataString)

    if (pageDataString == null) {
        WebUI.closeBrowser()

        return 
    } else {
        pageDataString = WebUI.getText(findTestObject(path)).replace('Объем потерь', '').replaceAll('\\s+', '').substring(
            0, obemPoter.length())

        println(pageDataString)

        def fileDataString = obemPoter

        println(fileDataString)

        def check = Check(def pageString = pageDataString, def fileString = fileDataString, path)

        path = 'Объем потерь сверка/Данные со страницы Объем потерь/Уровень потерь АО Тываэнерго'

        pageDataString = WebUI.getText(findTestObject(path))

        pageDataString = WebUI.getText(findTestObject(path)).replace('Уровень потерь', '').replaceAll('\\s+', '').substring(
            0, percentPoter.length())

        fileDataString = percentPoter

        def checkPercents = CheckPercents(pageString = pageDataString, fileString = fileDataString, path)

        if (check == false) {
            println('Start filial')

            WebUI.callTestCase(findTestCase('Объем потерь сверка/Сверка по филиалам/Объем потерь Филиалы Россети Урал(ГК)'), 
                [:], FailureHandling.CONTINUE_ON_FAILURE)
        } else if (checkPercents == false) {
            println('Start filial')

            WebUI.callTestCase(findTestCase('Объем потерь сверка/Сверка по филиалам/Объем потерь Филиалы Россети Урал(ГК)'), 
                [:], FailureHandling.CONTINUE_ON_FAILURE)
        } else {
            println('End case DZO')

            def scanErr = ScannErrors(path)

            WebUI.deleteAllCookies()

            WebUI.closeBrowser()
        }
    }
}

static def Test13() {
    '13'
    def start = OpenBrowser()

    def obemPoter = findTestData('Test Data').getValue(2, 88)

    println(obemPoter)

    def percentPoter = findTestData('Test Data').getValue(3, 88)

    println(percentPoter)

    WebUI.click(findTestObject('Объем потерь (Данные в виджетах)/фильтр ДЗО'))

    WebUI.click(findTestObject('Объем потерь (Данные в виджетах)/Снять выделения в фильтре ДЗО'))

    WebUI.click(findTestObject('Объем потерь (Данные в виджетах)/Применить в фильтре ДЗО'))

    WebUI.click(findTestObject('Объем потерь (Данные в виджетах)/фильтр ДЗО'))

    WebUI.click(findTestObject('Объем потерь сверка/ПАО Росссети'))

    WebUI.click(findTestObject('Объем потерь сверка/Магистральные сети'))

    WebUI.click(findTestObject('Объем потерь сверка/Россети ФСК ЕЭС'))

    WebUI.click(findTestObject('Объем потерь (Данные в виджетах)/Применить в фильтре ДЗО'))

    def selectDate = SelectDate()

    def write2 = WriteToExcel2(def err = 'Сверка не прошла')

    def path = 'Объем потерь сверка/Данные со страницы Объем потерь/Объем потерь АО Тываэнерго'

    def pageDataString = WebUI.getText(findTestObject(path))

    println(pageDataString)

    if (pageDataString == null) {
        WebUI.closeBrowser()

        return 
    } else {
        pageDataString = WebUI.getText(findTestObject(path)).replace('Объем потерь', '').replaceAll('\\s+', '').substring(
            0, obemPoter.length())

        println(pageDataString)

        def fileDataString = obemPoter

        println(fileDataString)

        def check = Check(def pageString = pageDataString, def fileString = fileDataString, path)

        path = 'Объем потерь сверка/Данные со страницы Объем потерь/Уровень потерь АО Тываэнерго'

        pageDataString = WebUI.getText(findTestObject(path)).replace('Уровень потерь', '').replaceAll('\\s+', '').substring(
            0, percentPoter.length())

        fileDataString = percentPoter

        def checkPercents = CheckPercents(pageString = pageDataString, fileString = fileDataString, path)

        def scanErr = ScannErrors(path)

        WebUI.deleteAllCookies()

        WebUI.closeBrowser()
    }
}

static def Test14() {
    '14'
    def start = OpenBrowser()

    def obemPoter = findTestData('Test Data').getValue(2, 71)

    println(obemPoter)

    def percentPoter = findTestData('Test Data').getValue(3, 71)

    println(percentPoter)

    def filterDzo = SelectDzo()

    WebUI.scrollToElement(findTestObject('Объем потерь сверка/Россети Янтарь'), 30)

    WebUI.scrollToElement(findTestObject('Объем потерь (Данные в виджетах)/скролл'), 30)

    WebUI.click(findTestObject('Объем потерь сверка/выбрать Россети Центр'))

    WebUI.click(findTestObject('Объем потерь (Данные в виджетах)/Применить в фильтре ДЗО'))

    def selectDate = SelectDate()

    def write2 = WriteToExcel2(def err = 'Сверка не прошла')

    def path = 'Объем потерь сверка/Данные со страницы Объем потерь/Объем потерь АО Тываэнерго'

    def pageDataString = WebUI.getText(findTestObject(path))

    println(pageDataString)

    if (pageDataString == null) {
        WebUI.closeBrowser()

        return 
    } else {
        pageDataString = WebUI.getText(findTestObject(path)).replace('Объем потерь', '').replaceAll('\\s+', '').substring(
            0, obemPoter.length())

        println(pageDataString)

        def fileDataString = obemPoter

        println(fileDataString)

        def check = Check(def pageString = pageDataString, def fileString = fileDataString, path)

        path = 'Объем потерь сверка/Данные со страницы Объем потерь/Уровень потерь АО Тываэнерго'

        pageDataString = WebUI.getText(findTestObject(path)).replace('Уровень потерь', '').replaceAll('\\s+', '').substring(
            0, percentPoter.length())

        fileDataString = percentPoter

        def checkPercents = CheckPercents(pageString = pageDataString, fileString = fileDataString, path)

        if (check == false) {
            println('Start filial')

            WebUI.callTestCase(findTestCase('Объем потерь сверка/Сверка по филиалам/Объем потерь Филиалы Россети Центр'), 
                [:], FailureHandling.CONTINUE_ON_FAILURE)
        } else if (checkPercents == false) {
            println('Start filial')

            WebUI.callTestCase(findTestCase('Объем потерь сверка/Сверка по филиалам/Объем потерь Филиалы Россети Центр'), 
                [:], FailureHandling.CONTINUE_ON_FAILURE)
        } else {
            println('End case DZO')

            def scanErr = ScannErrors(path)

            WebUI.deleteAllCookies()

            WebUI.closeBrowser()
        }
    }
}

static def Test15() {
    '15'
    def start = OpenBrowser()

    def obemPoter = findTestData('Test Data').getValue(2, 72)

    println(obemPoter)

    def percentPoter = findTestData('Test Data').getValue(3, 72)

    println(percentPoter)

    def filterDzo = SelectDzo()

    WebUI.scrollToElement(findTestObject('Объем потерь сверка/Россети Янтарь'), 30)

    WebUI.scrollToElement(findTestObject('Объем потерь (Данные в виджетах)/скролл'), 30)

    WebUI.click(findTestObject('Объем потерь сверка/выбрать Россети Центр и Приволжье(ГК)'))

    WebUI.click(findTestObject('Объем потерь (Данные в виджетах)/Применить в фильтре ДЗО'))

    def selectDate = SelectDate()

    def write2 = WriteToExcel2(def err = 'Сверка не прошла')

    def path = 'Объем потерь сверка/Данные со страницы Объем потерь/Объем потерь АО Тываэнерго'

    def pageDataString = WebUI.getText(findTestObject(path))

    println(pageDataString)

    if (pageDataString == null) {
        WebUI.closeBrowser()

        return 
    } else {
        pageDataString = WebUI.getText(findTestObject(path)).replace('Объем потерь', '').replaceAll('\\s+', '').substring(
            0, obemPoter.length())

        println(pageDataString)

        def fileDataString = obemPoter

        println(fileDataString)

        def check = Check(def pageString = pageDataString, def fileString = fileDataString, path)

        path = 'Объем потерь сверка/Данные со страницы Объем потерь/Уровень потерь АО Тываэнерго'

        pageDataString = WebUI.getText(findTestObject(path)).replace('Уровень потерь', '').replaceAll('\\s+', '').substring(
            0, percentPoter.length())

        fileDataString = percentPoter

        def checkPercents = CheckPercents(pageString = pageDataString, fileString = fileDataString, path)

        if (check == false) {
            println('Start filial')

            WebUI.callTestCase(findTestCase('Объем потерь сверка/Сверка по филиалам/Объем потерь Филиалы Россети Центр и Приволжье(ГК)'), 
                [:], FailureHandling.CONTINUE_ON_FAILURE)
        } else if (checkPercents == false) {
            println('Start filial')

            WebUI.callTestCase(findTestCase('Объем потерь сверка/Сверка по филиалам/Объем потерь Филиалы Россети Центр и Приволжье(ГК)'), 
                [:], FailureHandling.CONTINUE_ON_FAILURE)
        } else {
            println('End case DZO')

            def scanErr = ScannErrors(path)

            WebUI.deleteAllCookies()

            WebUI.closeBrowser()
        }
    }
}

static def Test16() {
    '16'
    def start = OpenBrowser()

    def obemPoter = findTestData('Test Data').getValue(2, 79)

    println(obemPoter)

    def percentPoter = findTestData('Test Data').getValue(3, 79)

    println(percentPoter)

    def filterDzo = SelectDzo()

    WebUI.scrollToElement(findTestObject('Объем потерь сверка/Россети Янтарь'), 30)

    WebUI.scrollToElement(findTestObject('Объем потерь (Данные в виджетах)/скролл'), 30)

    WebUI.click(findTestObject('Объем потерь сверка/выбрать Россети Юг(ГК)'))

    WebUI.click(findTestObject('Объем потерь (Данные в виджетах)/Применить в фильтре ДЗО'))

    def selectDate = SelectDate()

    def write2 = WriteToExcel2(def err = 'Сверка не прошла')

    def path = 'Объем потерь сверка/Данные со страницы Объем потерь/Объем потерь АО Тываэнерго'

    def pageDataString = WebUI.getText(findTestObject(path))

    println(pageDataString)

    if (pageDataString == null) {
        WebUI.closeBrowser()

        return 
    } else {
        pageDataString = WebUI.getText(findTestObject(path)).replace('Объем потерь', '').replaceAll('\\s+', '').substring(
            0, obemPoter.length())

        println(pageDataString)

        def fileDataString = obemPoter

        println(fileDataString)

        def check = Check(def pageString = pageDataString, def fileString = fileDataString, path)

        path = 'Объем потерь сверка/Данные со страницы Объем потерь/Уровень потерь АО Тываэнерго'

        pageDataString = WebUI.getText(findTestObject(path)).replace('Уровень потерь', '').replaceAll('\\s+', '').substring(
            0, percentPoter.length())

        fileDataString = percentPoter

        def checkPercents = CheckPercents(pageString = pageDataString, fileString = fileDataString, path)

        if (check == false) {
            println('Start filial')

            WebUI.callTestCase(findTestCase('Объем потерь сверка/Сверка по филиалам/Объем потерь Филиалы Россети Юг(ГК)'), 
                [:], FailureHandling.CONTINUE_ON_FAILURE)
        } else if (checkPercents == false) {
            println('Start filial')

            WebUI.callTestCase(findTestCase('Объем потерь сверка/Сверка по филиалам/Объем потерь Филиалы Россети Юг(ГК)'), 
                [:], FailureHandling.CONTINUE_ON_FAILURE)
        } else {
            println('End case DZO')

            def scanErr = ScannErrors(path)

            WebUI.deleteAllCookies()

            WebUI.closeBrowser()
        }
    }
}

static def Test17() {
    '17'
    def start = OpenBrowser()

    def obemPoter = findTestData('Test Data').getValue(2, 86)

    println(obemPoter)

    def percentPoter = findTestData('Test Data').getValue(3, 86)

    println(percentPoter)

    def filterDzo = SelectDzo()

    WebUI.scrollToElement(findTestObject('Объем потерь сверка/Россети Янтарь'), 30)

    WebUI.scrollToElement(findTestObject('Объем потерь (Данные в виджетах)/скролл'), 30)

    WebUI.click(findTestObject('Объем потерь сверка/выбрать Россети Янтарь'))

    WebUI.click(findTestObject('Объем потерь (Данные в виджетах)/Применить в фильтре ДЗО'))

    def selectDate = SelectDate()

    def write2 = WriteToExcel2(def err = 'Сверка не прошла')

    def path = 'Объем потерь сверка/Данные со страницы Объем потерь/Объем потерь АО Тываэнерго'

    def pageDataString = WebUI.getText(findTestObject(path))

    println(pageDataString)

    if (pageDataString == null) {
        WebUI.closeBrowser()

        return 
    } else {
        pageDataString = WebUI.getText(findTestObject(path)).replace('Объем потерь', '').replaceAll('\\s+', '').substring(
            0, obemPoter.length())

        def fileDataString = obemPoter

        println(fileDataString)

        def check = Check(def pageString = pageDataString, def fileString = fileDataString, path)

        path = 'Объем потерь сверка/Данные со страницы Объем потерь/Уровень потерь АО Тываэнерго'

        pageDataString = WebUI.getText(findTestObject(path)).replace('Уровень потерь', '').replaceAll('\\s+', '').substring(
            0, percentPoter.length())

        fileDataString = percentPoter

        def checkPercents = CheckPercents(pageString = pageDataString, fileString = fileDataString, path)

        def scanErr = ScannErrors(path)

        WebUI.deleteAllCookies()

        WebUI.closeBrowser()
    }
}

