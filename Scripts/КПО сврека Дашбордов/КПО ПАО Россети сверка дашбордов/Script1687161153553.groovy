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

def test1 = VyruchkaVseToggleMln1()

def test2 = VyruchkaOneToggleMln2()

def test3 = VyruchkaTwoToggleMln3()

def test4 = VyruchkaVseToggleProc4()

def test5 = VyruchkaOneToggleProc5()

def test6 = VyruchkaTwoToggleProc6()

static def VyruchkaVseToggleMln1() {
    WebUI.openBrowser('')

    'БЛОК РУКОВОДИТЕЛЕЙ'
    WebUI.navigateToUrl(findTestData('Test Data').getValue(7, 1))

    WebUI.setText(findTestObject('Object Repository/КПО/input__password'), findTestData('Test Data').getValue(6, 1))

    WebUI.setText(findTestObject('Object Repository/КПО/input__username'), findTestData('Test Data').getValue(5, 1))

    WebUI.click(findTestObject('Object Repository/КПО/button_'))

    WebUI.click(findTestObject('КПО/фильтр Выручка'))

    WebUI.click(findTestObject('КПО/снять выделения в фильтре Выручка'))

    WebUI.click(findTestObject('КПО/применить в фильтре Выручка'))

    WebUI.click(findTestObject('КПО/фильтр Дата'))

    WebUI.click(findTestObject('КПО/снять выделения в фильтре Дата'))

    WebUI.click(findTestObject('КПО/применить в фильтре Дата'))

    WebUI.click(findTestObject('КПО/фильтр Дата'))

    WebUI.scrollToElement(findTestObject('КПО/2024 год'), 30)

    WebUI.scrollToElement(findTestObject('КПО/скролл Фильтр дата'), 30)

    WebUI.click(findTestObject('КПО/2024 год'), FailureHandling.CONTINUE_ON_FAILURE)

    WebUI.scrollToElement(findTestObject('КПО/раскрыть 1 квартла 2024'), 30)

    WebUI.scrollToElement(findTestObject('КПО/скролл Фильтр дата'), 30)

    WebUI.click(findTestObject('КПО/раскрыть 1 квартла 2024'))

    WebUI.scrollToElement(findTestObject('КПО/Январь 2024'), 30)

    WebUI.scrollToElement(findTestObject('КПО/скролл Фильтр дата'), 30)

    WebUI.click(findTestObject('КПО/Январь 2024'))

    WebUI.click(findTestObject('КПО/Февраль 2024'))

    WebUI.click(findTestObject('КПО/применить в фильтре Дата'))

    WebUI.click(findTestObject('Object Repository/КПО/фильтр ДЗО'))

    WebUI.click(findTestObject('КПО/снять выделения в фильтре ДЗО'))

    WebUI.click(findTestObject('КПО/применить в фильтре ДЗО'))

    WebUI.delay(5)

    String a = WebUI.getText(findTestObject('КПО/Данные с виджета факт1'))

    String a1 = WebUI.getText(findTestObject('КПО/Данные с виджета ПЛАН1'))

    String c = WebUI.getText(findTestObject('КПО/Данные с виджета Отклонения - тогл млн/ПАО Россети'))

    String d = WebUI.getText(findTestObject('КПО/Данные с виджета Отклонения - тогл млн/CentIPriv'))

    String e = WebUI.getText(findTestObject('КПО/Данные с виджета Отклонения - тогл млн/Centr'))

    String f = WebUI.getText(findTestObject('КПО/Данные с виджета Отклонения - тогл млн/chech'))

    String g = WebUI.getText(findTestObject('КПО/Данные с виджета Отклонения - тогл млн/kub'))

    String m = WebUI.getText(findTestObject('КПО/Данные с виджета Отклонения - тогл млн/len'))

    String n = WebUI.getText(findTestObject('КПО/Данные с виджета Отклонения - тогл млн/mo'))

    String o = WebUI.getText(findTestObject('КПО/Данные с виджета Отклонения - тогл млн/sevKaz'))

    String p = WebUI.getText(findTestObject('КПО/Данные с виджета Отклонения - тогл млн/SevZap'))

    String r = WebUI.getText(findTestObject('КПО/Данные с виджета Отклонения - тогл млн/Sibir'))

    String s = WebUI.getText(findTestObject('КПО/Данные с виджета Отклонения - тогл млн/Tomsk'))

    String t = WebUI.getText(findTestObject('КПО/Данные с виджета Отклонения - тогл млн/Tumen'))

    String u = WebUI.getText(findTestObject('КПО/Данные с виджета Отклонения - тогл млн/tuv'))

    String y = WebUI.getText(findTestObject('КПО/Данные с виджета Отклонения - тогл млн/Ug'))

    String z = WebUI.getText(findTestObject('КПО/Данные с виджета Отклонения - тогл млн/Ural'))

    String x = WebUI.getText(findTestObject('КПО/Данные с виджета Отклонения - тогл млн/volg'))

    String w = WebUI.getText(findTestObject('КПО/Данные с виджета Отклонения - тогл млн/Yantar'))

    println(a)

    println(a1)

    println(c)

    'ВЫРУЧКА'
    WebUI.navigateToUrl(findTestData('Test Data').getValue(7, 2))

    WebUI.delay(15)

    if (WebUI.verifyTextPresent('Просьба обратить внимание', true) == true) {
        WebUI.click(findTestObject('КПО для раздела Выручка/закрыть уведомление'))
    }
    
    WebUI.click(findTestObject('КПО для раздела Выручка/Фильтр ДЗО'))

    WebUI.click(findTestObject('КПО для раздела Выручка/снять выделения в ДЗО'))

    WebUI.click(findTestObject('КПО для раздела Выручка/применить в фильтре ДЗО'))

    WebUI.delay(5)

    String b = WebUI.getText(findTestObject('КПО для раздела Выручка/Данные с виджета факт'))

    String b1 = WebUI.getText(findTestObject('КПО для раздела Выручка/Данные с виджета ПЛАН'))

    String с1 = WebUI.getText(findTestObject('КПО для раздела Выручка/Данные с виджета Отклонения - тогл млн/ПАО Россети'))

    String d1 = WebUI.getText(findTestObject('КПО для раздела Выручка/Данные с виджета Отклонения - тогл млн/CentIPriv'))

    String e1 = WebUI.getText(findTestObject('КПО для раздела Выручка/Данные с виджета Отклонения - тогл млн/Centr'))

    String f1 = WebUI.getText(findTestObject('КПО для раздела Выручка/Данные с виджета Отклонения - тогл млн/chech'))

    String g1 = WebUI.getText(findTestObject('КПО для раздела Выручка/Данные с виджета Отклонения - тогл млн/kub'))

    String m1 = WebUI.getText(findTestObject('КПО для раздела Выручка/Данные с виджета Отклонения - тогл млн/len'))

    String n1 = WebUI.getText(findTestObject('КПО для раздела Выручка/Данные с виджета Отклонения - тогл млн/mo'))

    String o1 = WebUI.getText(findTestObject('КПО для раздела Выручка/Данные с виджета Отклонения - тогл млн/sevKaz'))

    String p1 = WebUI.getText(findTestObject('КПО для раздела Выручка/Данные с виджета Отклонения - тогл млн/SevZap'))

    String r1 = WebUI.getText(findTestObject('КПО для раздела Выручка/Данные с виджета Отклонения - тогл млн/Sibir'))

    String s1 = WebUI.getText(findTestObject('КПО для раздела Выручка/Данные с виджета Отклонения - тогл млн/Tomsk'))

    String t1 = WebUI.getText(findTestObject('КПО для раздела Выручка/Данные с виджета Отклонения - тогл млн/Tumen'))

    String u1 = WebUI.getText(findTestObject('КПО для раздела Выручка/Данные с виджета Отклонения - тогл млн/tuv'))

    String y1 = WebUI.getText(findTestObject('КПО для раздела Выручка/Данные с виджета Отклонения - тогл млн/Ug'))

    String z1 = WebUI.getText(findTestObject('КПО для раздела Выручка/Данные с виджета Отклонения - тогл млн/Ural'))

    String x1 = WebUI.getText(findTestObject('КПО для раздела Выручка/Данные с виджета Отклонения - тогл млн/volg'))

    String w1 = WebUI.getText(findTestObject('КПО для раздела Выручка/Данные с виджета Отклонения - тогл млн/Yantar'))

    println(b)

    println(b1)

    println(с1)

    if (WebUI.verifyEqual(a, b) && WebUI.verifyEqual(a1, b1)) {
        println('GOOD')
    } else {
        def write = WriteToExcel(String widget = 'Виджет: ВЫполнение плановых показателй', String dzo = 'ПАО Россети', String togle = 'Переключатель на : МЛН')
    }
    
    if (WebUI.verifyEqual(c, с1)) {
        println('GOOD')
    } else {
        def write = WriteToExcel(String widget = 'Виджет: Отклонение котлового полезного отпуска', String dzo = 'ПАО Россети', 
            String togle = 'Переключатель на : МЛН')
    }
    
    if (WebUI.verifyEqual(d, d1)) {
        println('GOOD')
    } else {
        def write = WriteToExcel(String widget = 'Виджет: Отклонение котлового полезного отпуска', String dzo = 'Центр и Приволжье', 
            String togle = 'Переключатель на : МЛН')
    }
    
    if (WebUI.verifyEqual(e, e1)) {
        println('GOOD')
    } else {
        def write = WriteToExcel(String widget = 'Виджет: Отклонение котлового полезного отпуска', String dzo = 'Центр', 
            String togle = 'Переключатель на : МЛН')
    }
    
    if (WebUI.verifyEqual(f, f1)) {
        println('GOOD')
    } else {
        def write = WriteToExcel(String widget = 'Виджет: Отклонение котлового полезного отпуска', String dzo = 'Чеченэнерго', 
            String togle = 'Переключатель на : МЛН')
    }
    
    if (WebUI.verifyEqual(g, g1)) {
        println('GOOD')
    } else {
        def write = WriteToExcel(String widget = 'Виджет: Отклонение котлового полезного отпуска', String dzo = 'Россети Кубань', 
            String togle = 'Переключатель на : МЛН')
    }
    
    if (WebUI.verifyEqual(m, m1)) {
        println('GOOD')
    } else {
        def write = WriteToExcel(String widget = 'Виджет: Отклонение котлового полезного отпуска', String dzo = 'Россети Ленэнерго(ГК)', 
            String togle = 'Переключатель на : МЛН')
    }
    
    if (WebUI.verifyEqual(n, n1)) {
        println('GOOD')
    } else {
        def write = WriteToExcel(String widget = 'Виджет: Отклонение котлового полезного отпуска', String dzo = 'Россети Москвоский регион', 
            String togle = 'Переключатель на : МЛН')
    }
    
    if (WebUI.verifyEqual(o, o1)) {
        println('GOOD')
    } else {
        def write = WriteToExcel(String widget = 'Виджет: Отклонение котлового полезного отпуска', String dzo = 'Россети Северный Кавказ(ГК)', 
            String togle = 'Переключатель на : МЛН')
    }
    
    if (WebUI.verifyEqual(p, p1)) {
        println('GOOD')
    } else {
        def write = WriteToExcel(String widget = 'Виджет: Отклонение котлового полезного отпуска', String dzo = 'Россети Северо-Запад', 
            String togle = 'Переключатель на : МЛН')
    }
    
    if (WebUI.verifyEqual(r, r1)) {
        println('GOOD')
    } else {
        def write = WriteToExcel(String widget = 'Виджет: Отклонение котлового полезного отпуска', String dzo = 'Россети Сибирь(ГК)', 
            String togle = 'Переключатель на : МЛН')
    }
    
    if (WebUI.verifyEqual(s, s1)) {
        println('GOOD')
    } else {
        def write = WriteToExcel(String widget = 'Виджет: Отклонение котлового полезного отпуска', String dzo = 'Россети Томск', 
            String togle = 'Переключатель на : МЛН')
    }
    
    if (WebUI.verifyEqual(t, t1)) {
        println('GOOD')
    } else {
        def write = WriteToExcel(String widget = 'Виджет: Отклонение котлового полезного отпуска', String dzo = 'Россети Тюмень', 
            String togle = 'Переключатель на : МЛН')
    }
    
    if (WebUI.verifyEqual(u, u1)) {
        println('GOOD')
    } else {
        def write = WriteToExcel(String widget = 'Виджет: Отклонение котлового полезного отпуска', String dzo = 'Тываэнерго', 
            String togle = 'Переключатель на : МЛН')
    }
    
    if (WebUI.verifyEqual(y, y1)) {
        println('GOOD')
    } else {
        def write = WriteToExcel(String widget = 'Виджет: Отклонение котлового полезного отпуска', String dzo = 'Россети Юг(ГК)', 
            String togle = 'Переключатель на : МЛН')
    }
    
    if (WebUI.verifyEqual(z, z1)) {
        println('GOOD')
    } else {
        def write = WriteToExcel(String widget = 'Виджет: Отклонение котлового полезного отпуска', String dzo = 'Россети Урал(ГК)', 
            String togle = 'Переключатель на : МЛН')
    }
    
    if (WebUI.verifyEqual(x, x1)) {
        println('GOOD')
    } else {
        def write = WriteToExcel(String widget = 'Виджет: Отклонение котлового полезного отпуска', String dzo = 'Россети Волга', 
            String togle = 'Переключатель на : МЛН')
    }
    
    if (WebUI.verifyEqual(w, w1)) {
        println('GOOD')
    } else {
        def write = WriteToExcel(String widget = 'Виджет: Отклонение котлового полезного отпуска', String dzo = 'Россети Янтарь', 
            String togle = 'Переключатель на : МЛН')
    }
    
    WebUI.closeBrowser()
}

static def VyruchkaOneToggleMln2() {
    WebUI.openBrowser('')

    'БЛОК РУКОВОДИТЕЛЕЙ'
    WebUI.navigateToUrl(findTestData('Test Data').getValue(7, 1))

    WebUI.setText(findTestObject('Object Repository/КПО/input__password'), findTestData('Test Data').getValue(6, 1))

    WebUI.setText(findTestObject('Object Repository/КПО/input__username'), findTestData('Test Data').getValue(5, 1))

    WebUI.click(findTestObject('Object Repository/КПО/button_'))

    WebUI.click(findTestObject('КПО/фильтр Выручка'))

    WebUI.click(findTestObject('КПО/снять выделения в фильтре Выручка'))

    WebUI.click(findTestObject('КПО/выбрать Объем Электроэнергии'))

    WebUI.click(findTestObject('КПО/применить в фильтре Выручка'))

    WebUI.click(findTestObject('КПО/фильтр Дата'))

    WebUI.click(findTestObject('КПО/снять выделения в фильтре Дата'))

    WebUI.click(findTestObject('КПО/применить в фильтре Дата'))

    WebUI.click(findTestObject('КПО/фильтр Дата'))

    WebUI.scrollToElement(findTestObject('КПО/2024 год'), 30)

    WebUI.scrollToElement(findTestObject('КПО/скролл Фильтр дата'), 30)

    WebUI.click(findTestObject('КПО/2024 год'), FailureHandling.CONTINUE_ON_FAILURE)

    WebUI.scrollToElement(findTestObject('КПО/раскрыть 1 квартла 2024'), 30)

    WebUI.scrollToElement(findTestObject('КПО/скролл Фильтр дата'), 30)

    WebUI.click(findTestObject('КПО/раскрыть 1 квартла 2024'))

    WebUI.scrollToElement(findTestObject('КПО/Январь 2024'), 30)

    WebUI.scrollToElement(findTestObject('КПО/скролл Фильтр дата'), 30)

    WebUI.click(findTestObject('КПО/Январь 2024'))

    WebUI.click(findTestObject('КПО/Февраль 2024'))

    WebUI.click(findTestObject('КПО/применить в фильтре Дата'))

    WebUI.click(findTestObject('Object Repository/КПО/фильтр ДЗО'))

    WebUI.click(findTestObject('КПО/снять выделения в фильтре ДЗО'))

    WebUI.click(findTestObject('КПО/применить в фильтре ДЗО'))

    WebUI.delay(5)

    String a = WebUI.getText(findTestObject('КПО/Данные с виджета факт1'))

    String a1 = WebUI.getText(findTestObject('КПО/Данные с виджета ПЛАН1'))

    String c = WebUI.getText(findTestObject('КПО/Данные с виджета Отклонения - тогл млн - Выручка Объем электроэнергии/PAO one'))

    String d = WebUI.getText(findTestObject('КПО/Данные с виджета Отклонения - тогл млн - Выручка Объем электроэнергии/tuv one'))

    String e = WebUI.getText(findTestObject('КПО/Данные с виджета Отклонения - тогл млн - Выручка Объем электроэнергии/chech onee'))

    String f = WebUI.getText(findTestObject('КПО/Данные с виджета Отклонения - тогл млн - Выручка Объем электроэнергии/sevkax one'))

    String g = WebUI.getText(findTestObject('КПО/Данные с виджета Отклонения - тогл млн - Выручка Объем электроэнергии/ug one'))

    println(a)

    println(a1)

    println(c)

    'ВЫРУЧКА'
    WebUI.navigateToUrl(findTestData('Test Data').getValue(7, 2))

    WebUI.delay(15)

    if (WebUI.verifyTextPresent('Просьба обратить внимание', true) == true) {
        WebUI.click(findTestObject('КПО для раздела Выручка/закрыть уведомление'))
    }
    
    WebUI.click(findTestObject('КПО для раздела Выручка/Фильтр ДЗО'))

    WebUI.click(findTestObject('КПО для раздела Выручка/снять выделения в ДЗО'))

    WebUI.click(findTestObject('КПО для раздела Выручка/применить в фильтре ДЗО'))

    WebUI.click(findTestObject('КПО для раздела Выручка/Фильтр Выручка'))

    WebUI.click(findTestObject('КПО для раздела Выручка/dsNATb'))

    WebUI.click(findTestObject('КПО для раздела Выручка/OBEM'))

    WebUI.click(findTestObject('КПО для раздела Выручка/PRIM'))

    WebUI.delay(5)

    String b = WebUI.getText(findTestObject('КПО для раздела Выручка/Данные с виджета факт'))

    String b1 = WebUI.getText(findTestObject('КПО для раздела Выручка/Данные с виджета ПЛАН'))

    String с1 = WebUI.getText(findTestObject('КПО для раздела Выручка/Данные с виджета Отклонения - тогл млн - Выручка Объем электроэнергии/PAO vurch'))

    String d1 = WebUI.getText(findTestObject('КПО для раздела Выручка/Данные с виджета Отклонения - тогл млн - Выручка Объем электроэнергии/tuv vurch'))

    String e1 = WebUI.getText(findTestObject('КПО для раздела Выручка/Данные с виджета Отклонения - тогл млн - Выручка Объем электроэнергии/chech vurch'))

    String f1 = WebUI.getText(findTestObject('КПО для раздела Выручка/Данные с виджета Отклонения - тогл млн - Выручка Объем электроэнергии/sevkaz vurch'))

    String g1 = WebUI.getText(findTestObject('КПО для раздела Выручка/Данные с виджета Отклонения - тогл млн - Выручка Объем электроэнергии/ug vyruch'))

    println(b)

    println(b1)

    println(с1)

    if (WebUI.verifyEqual(a, b) && WebUI.verifyEqual(a1, b1)) {
        println('GOOD')
    } else {
        def write = WriteToExcel(String widget = 'Виджет: ВЫполнение плановых показателй', String dzo = 'ПАО Россети', String togle = 'Переключатель на : МЛН')
    }
    
    if (WebUI.verifyEqual(c, с1)) {
        println('GOOD')
    } else {
        def write = WriteToExcel(String widget = 'Виджет: Отклонение котлового полезного отпуска', String dzo = 'ПАО Россети', 
            String togle = 'Переключатель на : МЛН')
    }
    
    if (WebUI.verifyEqual(d, d1)) {
        println('GOOD')
    } else {
        def write = WriteToExcel(String widget = 'Виджет: Отклонение котлового полезного отпуска', String dzo = 'АО Тываэнерго', 
            String togle = 'Переключатель на : МЛН')
    }
    
    if (WebUI.verifyEqual(e, e1)) {
        println('GOOD')
    } else {
        def write = WriteToExcel(String widget = 'Виджет: Отклонение котлового полезного отпуска', String dzo = 'АО Чеченэнерго', 
            String togle = 'Переключатель на : МЛН')
    }
    
    if (WebUI.verifyEqual(f, f1)) {
        println('GOOD')
    } else {
        def write = WriteToExcel(String widget = 'Виджет: Отклонение котлового полезного отпуска', String dzo = 'Россети Северный Кавказ(ГК)', 
            String togle = 'Переключатель на : МЛН')
    }
    
    if (WebUI.verifyEqual(g, g1)) {
        println('GOOD')
    } else {
        def write = WriteToExcel(String widget = 'Виджет: Отклонение котлового полезного отпуска', String dzo = 'Россети Юг(ГК)', 
            String togle = 'Переключатель на : МЛН')
    }
    
    WebUI.closeBrowser()
}

static def VyruchkaTwoToggleMln3() {
    WebUI.openBrowser('')

    'БЛОК РУКОВОДИТЕЛЕЙ'
    WebUI.navigateToUrl(findTestData('Test Data').getValue(7, 1))

    WebUI.setText(findTestObject('Object Repository/КПО/input__password'), findTestData('Test Data').getValue(6, 1))

    WebUI.setText(findTestObject('Object Repository/КПО/input__username'), findTestData('Test Data').getValue(5, 1))

    WebUI.click(findTestObject('Object Repository/КПО/button_'))

    WebUI.click(findTestObject('КПО/фильтр Выручка'))

    WebUI.click(findTestObject('КПО/выбрать Объем реалицаии передачи ээ'))

    WebUI.click(findTestObject('КПО/применить в фильтре Выручка'))

    WebUI.click(findTestObject('КПО/фильтр Дата'))

    WebUI.click(findTestObject('КПО/снять выделения в фильтре Дата'))

    WebUI.click(findTestObject('КПО/применить в фильтре Дата'))

    WebUI.click(findTestObject('КПО/фильтр Дата'))

    WebUI.scrollToElement(findTestObject('КПО/2024 год'), 30)

    WebUI.scrollToElement(findTestObject('КПО/скролл Фильтр дата'), 30)

    WebUI.click(findTestObject('КПО/2024 год'), FailureHandling.CONTINUE_ON_FAILURE)

    WebUI.scrollToElement(findTestObject('КПО/раскрыть 1 квартла 2024'), 30)

    WebUI.scrollToElement(findTestObject('КПО/скролл Фильтр дата'), 30)

    WebUI.click(findTestObject('КПО/раскрыть 1 квартла 2024'))

    WebUI.scrollToElement(findTestObject('КПО/Январь 2024'), 30)

    WebUI.scrollToElement(findTestObject('КПО/скролл Фильтр дата'), 30)

    WebUI.click(findTestObject('КПО/Январь 2024'))

    WebUI.click(findTestObject('КПО/Февраль 2024'))

    WebUI.click(findTestObject('КПО/применить в фильтре Дата'))

    WebUI.click(findTestObject('Object Repository/КПО/фильтр ДЗО'))

    WebUI.click(findTestObject('КПО/снять выделения в фильтре ДЗО'))

    WebUI.click(findTestObject('КПО/применить в фильтре ДЗО'))

    WebUI.delay(5)

    String a = WebUI.getText(findTestObject('КПО/Данные с виджета факт1'))

    String a1 = WebUI.getText(findTestObject('КПО/Данные с виджета ПЛАН1'))

    String c = WebUI.getText(findTestObject('КПО/Данные с виджета Отклонения - тогл млн/ПАО Россети'))

    String d = WebUI.getText(findTestObject('КПО/Данные с виджета Отклонения - тогл млн/CentIPriv'))

    String e = WebUI.getText(findTestObject('КПО/Данные с виджета Отклонения - тогл млн/Centr'))

    String f = WebUI.getText(findTestObject('КПО/Данные с виджета Отклонения - тогл млн/chech'))

    String g = WebUI.getText(findTestObject('КПО/Данные с виджета Отклонения - тогл млн/kub'))

    String m = WebUI.getText(findTestObject('КПО/Данные с виджета Отклонения - тогл млн/len'))

    String n = WebUI.getText(findTestObject('КПО/Данные с виджета Отклонения - тогл млн/mo'))

    String o = WebUI.getText(findTestObject('КПО/Данные с виджета Отклонения - тогл млн/sevKaz'))

    String p = WebUI.getText(findTestObject('КПО/Данные с виджета Отклонения - тогл млн/SevZap'))

    String r = WebUI.getText(findTestObject('КПО/Данные с виджета Отклонения - тогл млн/Sibir'))

    String s = WebUI.getText(findTestObject('КПО/Данные с виджета Отклонения - тогл млн/Tomsk'))

    String t = WebUI.getText(findTestObject('КПО/Данные с виджета Отклонения - тогл млн/Tumen'))

    String u = WebUI.getText(findTestObject('КПО/Данные с виджета Отклонения - тогл млн/tuv'))

    String y = WebUI.getText(findTestObject('КПО/Данные с виджета Отклонения - тогл млн/Ug'))

    String z = WebUI.getText(findTestObject('КПО/Данные с виджета Отклонения - тогл млн/Ural'))

    String x = WebUI.getText(findTestObject('КПО/Данные с виджета Отклонения - тогл млн/volg'))

    String w = WebUI.getText(findTestObject('КПО/Данные с виджета Отклонения - тогл млн/Yantar'))

    println(a)

    println(a1)

    println(c)

    'ВЫРУЧКА'
    WebUI.navigateToUrl(findTestData('Test Data').getValue(7, 2))

    WebUI.delay(15)

    if (WebUI.verifyTextPresent('Просьба обратить внимание', true) == true) {
        WebUI.click(findTestObject('КПО для раздела Выручка/закрыть уведомление'))
    }
    
    WebUI.click(findTestObject('КПО для раздела Выручка/Фильтр ДЗО'))

    WebUI.click(findTestObject('КПО для раздела Выручка/снять выделения в ДЗО'))

    WebUI.click(findTestObject('КПО для раздела Выручка/применить в фильтре ДЗО'))

    WebUI.click(findTestObject('КПО для раздела Выручка/Фильтр Выручка'))

    WebUI.click(findTestObject('КПО для раздела Выручка/dsNATb'))

    WebUI.click(findTestObject('КПО для раздела Выручка/div_EE'))

    WebUI.click(findTestObject('КПО для раздела Выручка/PRIM'))

    WebUI.delay(5)

    String b = WebUI.getText(findTestObject('КПО для раздела Выручка/Данные с виджета факт'))

    String b1 = WebUI.getText(findTestObject('КПО для раздела Выручка/Данные с виджета ПЛАН'))

    String с1 = WebUI.getText(findTestObject('КПО для раздела Выручка/Данные с виджета Отклонения - тогл млн/ПАО Россети'))

    String d1 = WebUI.getText(findTestObject('КПО для раздела Выручка/Данные с виджета Отклонения - тогл млн/CentIPriv'))

    String e1 = WebUI.getText(findTestObject('КПО для раздела Выручка/Данные с виджета Отклонения - тогл млн/Centr'))

    String f1 = WebUI.getText(findTestObject('КПО для раздела Выручка/Данные с виджета Отклонения - тогл млн/chech'))

    String g1 = WebUI.getText(findTestObject('КПО для раздела Выручка/Данные с виджета Отклонения - тогл млн/kub'))

    String m1 = WebUI.getText(findTestObject('КПО для раздела Выручка/Данные с виджета Отклонения - тогл млн/len'))

    String n1 = WebUI.getText(findTestObject('КПО для раздела Выручка/Данные с виджета Отклонения - тогл млн/mo'))

    String o1 = WebUI.getText(findTestObject('КПО для раздела Выручка/Данные с виджета Отклонения - тогл млн/sevKaz'))

    String p1 = WebUI.getText(findTestObject('КПО для раздела Выручка/Данные с виджета Отклонения - тогл млн/SevZap'))

    String r1 = WebUI.getText(findTestObject('КПО для раздела Выручка/Данные с виджета Отклонения - тогл млн/Sibir'))

    String s1 = WebUI.getText(findTestObject('КПО для раздела Выручка/Данные с виджета Отклонения - тогл млн/Tomsk'))

    String t1 = WebUI.getText(findTestObject('КПО для раздела Выручка/Данные с виджета Отклонения - тогл млн/Tumen'))

    String u1 = WebUI.getText(findTestObject('КПО для раздела Выручка/Данные с виджета Отклонения - тогл млн/tuv'))

    String y1 = WebUI.getText(findTestObject('КПО для раздела Выручка/Данные с виджета Отклонения - тогл млн/Ug'))

    String z1 = WebUI.getText(findTestObject('КПО для раздела Выручка/Данные с виджета Отклонения - тогл млн/Ural'))

    String x1 = WebUI.getText(findTestObject('КПО для раздела Выручка/Данные с виджета Отклонения - тогл млн/volg'))

    String w1 = WebUI.getText(findTestObject('КПО для раздела Выручка/Данные с виджета Отклонения - тогл млн/Yantar'))

    println(b)

    println(b1)

    println(с1)

    if (WebUI.verifyEqual(a, b) && WebUI.verifyEqual(a1, b1)) {
        println('GOOD')
    } else {
        def write = WriteToExcel(String widget = 'Виджет: ВЫполнение плановых показателй', String dzo = 'ПАО Россети', String togle = 'Переключатель на : МЛН')
    }
    
    if (WebUI.verifyEqual(c, с1)) {
        println('GOOD')
    } else {
        def write = WriteToExcel(String widget = 'Виджет: Отклонение котлового полезного отпуска', String dzo = 'ПАО Россети', 
            String togle = 'Переключатель на : МЛН')
    }
    
    if (WebUI.verifyEqual(d, d1)) {
        println('GOOD')
    } else {
        def write = WriteToExcel(String widget = 'Виджет: Отклонение котлового полезного отпуска', String dzo = 'Центр и Приволжье', 
            String togle = 'Переключатель на : МЛН')
    }
    
    if (WebUI.verifyEqual(e, e1)) {
        println('GOOD')
    } else {
        def write = WriteToExcel(String widget = 'Виджет: Отклонение котлового полезного отпуска', String dzo = 'Центр', 
            String togle = 'Переключатель на : МЛН')
    }
    
    if (WebUI.verifyEqual(f, f1)) {
        println('GOOD')
    } else {
        def write = WriteToExcel(String widget = 'Виджет: Отклонение котлового полезного отпуска', String dzo = 'Чеченэнерго', 
            String togle = 'Переключатель на : МЛН')
    }
    
    if (WebUI.verifyEqual(g, g1)) {
        println('GOOD')
    } else {
        def write = WriteToExcel(String widget = 'Виджет: Отклонение котлового полезного отпуска', String dzo = 'Россети Кубань', 
            String togle = 'Переключатель на : МЛН')
    }
    
    if (WebUI.verifyEqual(m, m1)) {
        println('GOOD')
    } else {
        def write = WriteToExcel(String widget = 'Виджет: Отклонение котлового полезного отпуска', String dzo = 'Россети Ленэнерго(ГК)', 
            String togle = 'Переключатель на : МЛН')
    }
    
    if (WebUI.verifyEqual(n, n1)) {
        println('GOOD')
    } else {
        def write = WriteToExcel(String widget = 'Виджет: Отклонение котлового полезного отпуска', String dzo = 'Россети Москвоский регион', 
            String togle = 'Переключатель на : МЛН')
    }
    
    if (WebUI.verifyEqual(o, o1)) {
        println('GOOD')
    } else {
        def write = WriteToExcel(String widget = 'Виджет: Отклонение котлового полезного отпуска', String dzo = 'Россети Северный Кавказ(ГК)', 
            String togle = 'Переключатель на : МЛН')
    }
    
    if (WebUI.verifyEqual(p, p1)) {
        println('GOOD')
    } else {
        def write = WriteToExcel(String widget = 'Виджет: Отклонение котлового полезного отпуска', String dzo = 'Россети Северо-Запад', 
            String togle = 'Переключатель на : МЛН')
    }
    
    if (WebUI.verifyEqual(r, r1)) {
        println('GOOD')
    } else {
        def write = WriteToExcel(String widget = 'Виджет: Отклонение котлового полезного отпуска', String dzo = 'Россети Сибирь(ГК)', 
            String togle = 'Переключатель на : МЛН')
    }
    
    if (WebUI.verifyEqual(s, s1)) {
        println('GOOD')
    } else {
        def write = WriteToExcel(String widget = 'Виджет: Отклонение котлового полезного отпуска', String dzo = 'Россети Томск', 
            String togle = 'Переключатель на : МЛН')
    }
    
    if (WebUI.verifyEqual(t, t1)) {
        println('GOOD')
    } else {
        def write = WriteToExcel(String widget = 'Виджет: Отклонение котлового полезного отпуска', String dzo = 'Россети Тюмень', 
            String togle = 'Переключатель на : МЛН')
    }
    
    if (WebUI.verifyEqual(u, u1)) {
        println('GOOD')
    } else {
        def write = WriteToExcel(String widget = 'Виджет: Отклонение котлового полезного отпуска', String dzo = 'Тываэнерго', 
            String togle = 'Переключатель на : МЛН')
    }
    
    if (WebUI.verifyEqual(y, y1)) {
        println('GOOD')
    } else {
        def write = WriteToExcel(String widget = 'Виджет: Отклонение котлового полезного отпуска', String dzo = 'Россети Юг(ГК)', 
            String togle = 'Переключатель на : МЛН')
    }
    
    if (WebUI.verifyEqual(z, z1)) {
        println('GOOD')
    } else {
        def write = WriteToExcel(String widget = 'Виджет: Отклонение котлового полезного отпуска', String dzo = 'Россети Урал(ГК)', 
            String togle = 'Переключатель на : МЛН')
    }
    
    if (WebUI.verifyEqual(x, x1)) {
        println('GOOD')
    } else {
        def write = WriteToExcel(String widget = 'Виджет: Отклонение котлового полезного отпуска', String dzo = 'Россети Волга', 
            String togle = 'Переключатель на : МЛН')
    }
    
    if (WebUI.verifyEqual(w, w1)) {
        println('GOOD')
    } else {
        def write = WriteToExcel(String widget = 'Виджет: Отклонение котлового полезного отпуска', String dzo = 'Россети Янтарь', 
            String togle = 'Переключатель на : МЛН')
    }
    
    WebUI.closeBrowser()
}

static def VyruchkaVseToggleProc4() {
    WebUI.openBrowser('')

    'БЛОК РУКОВОДИТЕЛЕЙ'
    WebUI.navigateToUrl(findTestData('Test Data').getValue(7, 1))

    WebUI.setText(findTestObject('Object Repository/КПО/input__password'), findTestData('Test Data').getValue(6, 1))

    WebUI.setText(findTestObject('Object Repository/КПО/input__username'), findTestData('Test Data').getValue(5, 1))

    WebUI.click(findTestObject('Object Repository/КПО/button_'))

    WebUI.click(findTestObject('КПО/тогл ПРОЦЕНТ'))

    WebUI.click(findTestObject('КПО/фильтр Выручка'))

    WebUI.click(findTestObject('КПО/снять выделения в фильтре Выручка'))

    WebUI.click(findTestObject('КПО/применить в фильтре Выручка'))

    WebUI.click(findTestObject('КПО/фильтр Дата'))

    WebUI.click(findTestObject('КПО/снять выделения в фильтре Дата'))

    WebUI.click(findTestObject('КПО/применить в фильтре Дата'))

    WebUI.click(findTestObject('КПО/фильтр Дата'))

    WebUI.scrollToElement(findTestObject('КПО/2024 год'), 30)

    WebUI.scrollToElement(findTestObject('КПО/скролл Фильтр дата'), 30)

    WebUI.click(findTestObject('КПО/2024 год'), FailureHandling.CONTINUE_ON_FAILURE)

    WebUI.scrollToElement(findTestObject('КПО/раскрыть 1 квартла 2024'), 30)

    WebUI.scrollToElement(findTestObject('КПО/скролл Фильтр дата'), 30)

    WebUI.click(findTestObject('КПО/раскрыть 1 квартла 2024'))

    WebUI.scrollToElement(findTestObject('КПО/Январь 2024'), 30)

    WebUI.scrollToElement(findTestObject('КПО/скролл Фильтр дата'), 30)

    WebUI.click(findTestObject('КПО/Январь 2024'))

    WebUI.click(findTestObject('КПО/Февраль 2024'))

    WebUI.click(findTestObject('КПО/применить в фильтре Дата'))

    WebUI.click(findTestObject('Object Repository/КПО/фильтр ДЗО'))

    WebUI.click(findTestObject('КПО/снять выделения в фильтре ДЗО'))

    WebUI.click(findTestObject('КПО/применить в фильтре ДЗО'))

    WebUI.delay(5)

    String a = WebUI.getText(findTestObject('КПО/Данные с виджета факт1'))

    String a1 = WebUI.getText(findTestObject('КПО/Данные с виджета ПЛАН1'))

    String c = WebUI.getText(findTestObject('КПО/Данные с виджета Отклонения - тогл процент/ПАО Россети'))

    String d = WebUI.getText(findTestObject('КПО/Данные с виджета Отклонения - тогл процент/CentIPriv'))

    String e = WebUI.getText(findTestObject('КПО/Данные с виджета Отклонения - тогл процент/Centr'))

    String f = WebUI.getText(findTestObject('КПО/Данные с виджета Отклонения - тогл процент/chech'))

    String g = WebUI.getText(findTestObject('КПО/Данные с виджета Отклонения - тогл процент/kub'))

    String m = WebUI.getText(findTestObject('КПО/Данные с виджета Отклонения - тогл процент/len'))

    String n = WebUI.getText(findTestObject('КПО/Данные с виджета Отклонения - тогл процент/mo'))

    String o = WebUI.getText(findTestObject('КПО/Данные с виджета Отклонения - тогл процент/sevKaz'))

    String p = WebUI.getText(findTestObject('КПО/Данные с виджета Отклонения - тогл процент/SevZap'))

    String r = WebUI.getText(findTestObject('КПО/Данные с виджета Отклонения - тогл процент/Sibir'))

    String s = WebUI.getText(findTestObject('КПО/Данные с виджета Отклонения - тогл процент/Tomsk'))

    String t = WebUI.getText(findTestObject('КПО/Данные с виджета Отклонения - тогл процент/Tumen'))

    String u = WebUI.getText(findTestObject('КПО/Данные с виджета Отклонения - тогл процент/tuv'))

    String y = WebUI.getText(findTestObject('КПО/Данные с виджета Отклонения - тогл процент/Ug'))

    String z = WebUI.getText(findTestObject('КПО/Данные с виджета Отклонения - тогл процент/Ural'))

    String x = WebUI.getText(findTestObject('КПО/Данные с виджета Отклонения - тогл процент/volg'))

    String w = WebUI.getText(findTestObject('КПО/Данные с виджета Отклонения - тогл процент/Yantar'))

    String q = WebUI.getText(findTestObject('КПО/Данные с виджета Отклонения - тогл процент/fsk'))

    println(a)

    println(a1)

    println(c)

    'ВЫРУЧКА'
    WebUI.navigateToUrl(findTestData('Test Data').getValue(7, 2))

    WebUI.delay(15)

    if (WebUI.verifyTextPresent('Просьба обратить внимание', true) == true) {
        WebUI.click(findTestObject('КПО для раздела Выручка/закрыть уведомление'))
    }
    
    WebUI.click(findTestObject('КПО для раздела Выручка/Фильтр ДЗО'))

    WebUI.click(findTestObject('КПО для раздела Выручка/снять выделения в ДЗО'))

    WebUI.click(findTestObject('КПО для раздела Выручка/применить в фильтре ДЗО'))

    WebUI.click(findTestObject('КПО для раздела Выручка/ТОГЛ НА ПРОЦЕНТ'))

    WebUI.delay(5)

    String b = WebUI.getText(findTestObject('КПО для раздела Выручка/Данные с виджета факт'))

    String b1 = WebUI.getText(findTestObject('КПО для раздела Выручка/Данные с виджета ПЛАН'))

    String с1 = WebUI.getText(findTestObject('КПО для раздела Выручка/Данные с виджета Отклонения - тогл процент/ПАО Россети'))

    String d1 = WebUI.getText(findTestObject('КПО для раздела Выручка/Данные с виджета Отклонения - тогл процент/CentIPriv'))

    String e1 = WebUI.getText(findTestObject('КПО для раздела Выручка/Данные с виджета Отклонения - тогл процент/Centr'))

    String f1 = WebUI.getText(findTestObject('КПО для раздела Выручка/Данные с виджета Отклонения - тогл процент/chech'))

    String g1 = WebUI.getText(findTestObject('КПО для раздела Выручка/Данные с виджета Отклонения - тогл процент/kub'))

    String m1 = WebUI.getText(findTestObject('КПО для раздела Выручка/Данные с виджета Отклонения - тогл процент/len'))

    String n1 = WebUI.getText(findTestObject('КПО для раздела Выручка/Данные с виджета Отклонения - тогл процент/mo'))

    String o1 = WebUI.getText(findTestObject('КПО для раздела Выручка/Данные с виджета Отклонения - тогл процент/sevKaz'))

    String p1 = WebUI.getText(findTestObject('КПО для раздела Выручка/Данные с виджета Отклонения - тогл процент/SevZap'))

    String r1 = WebUI.getText(findTestObject('КПО для раздела Выручка/Данные с виджета Отклонения - тогл процент/Sibir'))

    String s1 = WebUI.getText(findTestObject('КПО для раздела Выручка/Данные с виджета Отклонения - тогл процент/Tomsk'))

    String t1 = WebUI.getText(findTestObject('КПО для раздела Выручка/Данные с виджета Отклонения - тогл процент/Tumen'))

    String u1 = WebUI.getText(findTestObject('КПО для раздела Выручка/Данные с виджета Отклонения - тогл процент/tuv'))

    String y1 = WebUI.getText(findTestObject('КПО для раздела Выручка/Данные с виджета Отклонения - тогл процент/Ug'))

    String z1 = WebUI.getText(findTestObject('КПО для раздела Выручка/Данные с виджета Отклонения - тогл процент/Ural'))

    String x1 = WebUI.getText(findTestObject('КПО для раздела Выручка/Данные с виджета Отклонения - тогл процент/volg'))

    String w1 = WebUI.getText(findTestObject('КПО для раздела Выручка/Данные с виджета Отклонения - тогл процент/Yantar'))

    String q1 = WebUI.getText(findTestObject('КПО для раздела Выручка/Данные с виджета Отклонения - тогл процент/fsk vuruch'))

    println(b)

    println(b1)

    println(с1)

    if (WebUI.verifyEqual(a, b) && WebUI.verifyEqual(a1, b1)) {
        println('GOOD')
    } else {
        def write = WriteToExcel(String widget = 'Виджет: ВЫполнение плановых показателй', String dzo = 'ПАО Россети', String togle = 'Переключатель на : %')
    }
    
    if (WebUI.verifyEqual(c, с1)) {
        println('GOOD')
    } else {
        def write = WriteToExcel(String widget = 'Виджет: Отклонение котлового полезного отпуска', String dzo = 'ПАО Россети', 
            String togle = 'Переключатель на : %')
    }
    
    if (WebUI.verifyEqual(d, d1)) {
        println('GOOD')
    } else {
        def write = WriteToExcel(String widget = 'Виджет: Отклонение котлового полезного отпуска', String dzo = 'Центр и Приволжье', 
            String togle = 'Переключатель на : %')
    }
    
    if (WebUI.verifyEqual(e, e1)) {
        println('GOOD')
    } else {
        def write = WriteToExcel(String widget = 'Виджет: Отклонение котлового полезного отпуска', String dzo = 'Центр', 
            String togle = 'Переключатель на : %')
    }
    
    if (WebUI.verifyEqual(f, f1)) {
        println('GOOD')
    } else {
        def write = WriteToExcel(String widget = 'Виджет: Отклонение котлового полезного отпуска', String dzo = 'Чеченэнерго', 
            String togle = 'Переключатель на : %')
    }
    
    if (WebUI.verifyEqual(g, g1)) {
        println('GOOD')
    } else {
        def write = WriteToExcel(String widget = 'Виджет: Отклонение котлового полезного отпуска', String dzo = 'Россети Кубань', 
            String togle = 'Переключатель на : %')
    }
    
    if (WebUI.verifyEqual(m, m1)) {
        println('GOOD')
    } else {
        def write = WriteToExcel(String widget = 'Виджет: Отклонение котлового полезного отпуска', String dzo = 'Россети Ленэнерго(ГК)', 
            String togle = 'Переключатель на : %')
    }
    
    if (WebUI.verifyEqual(n, n1)) {
        println('GOOD')
    } else {
        def write = WriteToExcel(String widget = 'Виджет: Отклонение котлового полезного отпуска', String dzo = 'Россети Москвоский регион', 
            String togle = 'Переключатель на : %')
    }
    
    if (WebUI.verifyEqual(o, o1)) {
        println('GOOD')
    } else {
        def write = WriteToExcel(String widget = 'Виджет: Отклонение котлового полезного отпуска', String dzo = 'Россети Северный Кавказ(ГК)', 
            String togle = 'Переключатель на : %')
    }
    
    if (WebUI.verifyEqual(p, p1)) {
        println('GOOD')
    } else {
        def write = WriteToExcel(String widget = 'Виджет: Отклонение котлового полезного отпуска', String dzo = 'Россети Северо-Запад', 
            String togle = 'Переключатель на : %')
    }
    
    if (WebUI.verifyEqual(r, r1)) {
        println('GOOD')
    } else {
        def write = WriteToExcel(String widget = 'Виджет: Отклонение котлового полезного отпуска', String dzo = 'Россети Сибирь(ГК)', 
            String togle = 'Переключатель на : %')
    }
    
    if (WebUI.verifyEqual(s, s1)) {
        println('GOOD')
    } else {
        def write = WriteToExcel(String widget = 'Виджет: Отклонение котлового полезного отпуска', String dzo = 'Россети Томск', 
            String togle = 'Переключатель на : %')
    }
    
    if (WebUI.verifyEqual(t, t1)) {
        println('GOOD')
    } else {
        def write = WriteToExcel(String widget = 'Виджет: Отклонение котлового полезного отпуска', String dzo = 'Россети Тюмень', 
            String togle = 'Переключатель на : %')
    }
    
    if (WebUI.verifyEqual(u, u1)) {
        println('GOOD')
    } else {
        def write = WriteToExcel(String widget = 'Виджет: Отклонение котлового полезного отпуска', String dzo = 'Тываэнерго', 
            String togle = 'Переключатель на : %')
    }
    
    if (WebUI.verifyEqual(y, y1)) {
        println('GOOD')
    } else {
        def write = WriteToExcel(String widget = 'Виджет: Отклонение котлового полезного отпуска', String dzo = 'Россети Юг(ГК)', 
            String togle = 'Переключатель на : %')
    }
    
    if (WebUI.verifyEqual(z, z1)) {
        println('GOOD')
    } else {
        def write = WriteToExcel(String widget = 'Виджет: Отклонение котлового полезного отпуска', String dzo = 'Россети Урал(ГК)', 
            String togle = 'Переключатель на : %')
    }
    
    if (WebUI.verifyEqual(x, x1)) {
        println('GOOD')
    } else {
        def write = WriteToExcel(String widget = 'Виджет: Отклонение котлового полезного отпуска', String dzo = 'Россети Волга', 
            String togle = 'Переключатель на : %')
    }
    
    if (WebUI.verifyEqual(w, w1)) {
        println('GOOD')
    } else {
        def write = WriteToExcel(String widget = 'Виджет: Отклонение котлового полезного отпуска', String dzo = 'Россети Янтарь', 
            String togle = 'Переключатель на : %')
    }
    
    if (WebUI.verifyEqual(q, q1)) {
        println('GOOD')
    } else {
        def write = WriteToExcel(String widget = 'Виджет: Отклонение котлового полезного отпуска', String dzo = 'Россети ФСК ЕЭС', 
            String togle = 'Переключатель на : %')
    }
    
    WebUI.closeBrowser()
}

static def VyruchkaOneToggleProc5() {
    WebUI.openBrowser('')

    'БЛОК РУКОВОДИТЕЛЕЙ'
    WebUI.navigateToUrl(findTestData('Test Data').getValue(7, 1))

    WebUI.setText(findTestObject('Object Repository/КПО/input__password'), findTestData('Test Data').getValue(6, 1))

    WebUI.setText(findTestObject('Object Repository/КПО/input__username'), findTestData('Test Data').getValue(5, 1))

    WebUI.click(findTestObject('Object Repository/КПО/button_'))

    WebUI.click(findTestObject('КПО/тогл ПРОЦЕНТ'))

    WebUI.click(findTestObject('КПО/фильтр Выручка'))

    WebUI.click(findTestObject('КПО/снять выделения в фильтре Выручка'))

    WebUI.click(findTestObject('КПО/выбрать Объем Электроэнергии'))

    WebUI.click(findTestObject('КПО/применить в фильтре Выручка'))

    WebUI.click(findTestObject('КПО/фильтр Дата'))

    WebUI.click(findTestObject('КПО/снять выделения в фильтре Дата'))

    WebUI.click(findTestObject('КПО/применить в фильтре Дата'))

    WebUI.click(findTestObject('КПО/фильтр Дата'))

    WebUI.scrollToElement(findTestObject('КПО/2024 год'), 30)

    WebUI.scrollToElement(findTestObject('КПО/скролл Фильтр дата'), 30)

    WebUI.click(findTestObject('КПО/2024 год'), FailureHandling.CONTINUE_ON_FAILURE)

    WebUI.scrollToElement(findTestObject('КПО/раскрыть 1 квартла 2024'), 30)

    WebUI.scrollToElement(findTestObject('КПО/скролл Фильтр дата'), 30)

    WebUI.click(findTestObject('КПО/раскрыть 1 квартла 2024'))

    WebUI.scrollToElement(findTestObject('КПО/Январь 2024'), 30)

    WebUI.scrollToElement(findTestObject('КПО/скролл Фильтр дата'), 30)

    WebUI.click(findTestObject('КПО/Январь 2024'))

    WebUI.click(findTestObject('КПО/Февраль 2024'))

    WebUI.click(findTestObject('КПО/применить в фильтре Дата'))

    WebUI.click(findTestObject('Object Repository/КПО/фильтр ДЗО'))

    WebUI.click(findTestObject('КПО/снять выделения в фильтре ДЗО'))

    WebUI.click(findTestObject('КПО/применить в фильтре ДЗО'))

    WebUI.delay(5)

    String a = WebUI.getText(findTestObject('КПО/Данные с виджета факт1'))

    String a1 = WebUI.getText(findTestObject('КПО/Данные с виджета ПЛАН1'))

    String c = WebUI.getText(findTestObject('КПО/Данные с виджета Отклонения - тогл процент- Выручка Объем электроэнергии/PAO ruk'))

    String d = WebUI.getText(findTestObject('КПО/Данные с виджета Отклонения - тогл процент- Выручка Объем электроэнергии/tuv ruk'))

    String e = WebUI.getText(findTestObject('КПО/Данные с виджета Отклонения - тогл процент- Выручка Объем электроэнергии/chech ruk'))

    String f = WebUI.getText(findTestObject('КПО/Данные с виджета Отклонения - тогл процент- Выручка Объем электроэнергии/sevkaz ruk'))

    String g = WebUI.getText(findTestObject('КПО/Данные с виджета Отклонения - тогл процент- Выручка Объем электроэнергии/ug ruk'))

    println(a)

    println(a1)

    println(c)

    'ВЫРУЧКА'
    WebUI.navigateToUrl(findTestData('Test Data').getValue(7, 2))

    WebUI.delay(15)

    if (WebUI.verifyTextPresent('Просьба обратить внимание', true) == true) {
        WebUI.click(findTestObject('КПО для раздела Выручка/закрыть уведомление'))
    }
    
    WebUI.click(findTestObject('КПО для раздела Выручка/ТОГЛ НА ПРОЦЕНТ'))

    WebUI.click(findTestObject('КПО для раздела Выручка/Фильтр ДЗО'))

    WebUI.click(findTestObject('КПО для раздела Выручка/снять выделения в ДЗО'))

    WebUI.click(findTestObject('КПО для раздела Выручка/применить в фильтре ДЗО'))

    WebUI.click(findTestObject('КПО для раздела Выручка/Фильтр Выручка'))

    WebUI.click(findTestObject('КПО для раздела Выручка/dsNATb'))

    WebUI.click(findTestObject('КПО для раздела Выручка/OBEM'))

    WebUI.click(findTestObject('КПО для раздела Выручка/PRIM'))

    WebUI.delay(5)

    String b = WebUI.getText(findTestObject('КПО для раздела Выручка/Данные с виджета факт'))

    String b1 = WebUI.getText(findTestObject('КПО для раздела Выручка/Данные с виджета ПЛАН'))

    String с1 = WebUI.getText(findTestObject('КПО для раздела Выручка/Данные с виджета Отклонения - тогл процент- Выручка Объем электроэнергии/PAO vur'))

    String d1 = WebUI.getText(findTestObject('КПО для раздела Выручка/Данные с виджета Отклонения - тогл процент- Выручка Объем электроэнергии/tuv vur'))

    String e1 = WebUI.getText(findTestObject('КПО для раздела Выручка/Данные с виджета Отклонения - тогл процент- Выручка Объем электроэнергии/chech vur'))

    String f1 = WebUI.getText(findTestObject('КПО для раздела Выручка/Данные с виджета Отклонения - тогл млн - Выручка Объем электроэнергии/sevkaz vurch'))

    String g1 = WebUI.getText(findTestObject('КПО для раздела Выручка/Данные с виджета Отклонения - тогл процент- Выручка Объем электроэнергии/ug vur'))

    println(b)

    println(b1)

    println(с1)

    if (WebUI.verifyEqual(a, b) && WebUI.verifyEqual(a1, b1)) {
        println('GOOD')
    } else {
        def write = WriteToExcel(String widget = 'Виджет: ВЫполнение плановых показателй', String dzo = 'ПАО Россети', String togle = 'Переключатель на : %')
    }
    
    if (WebUI.verifyEqual(c, с1)) {
        println('GOOD')
    } else {
        def write = WriteToExcel(String widget = 'Виджет: Отклонение котлового полезного отпуска', String dzo = 'ПАО Россети', 
            String togle = 'Переключатель на : %')
    }
    
    if (WebUI.verifyEqual(d, d1)) {
        println('GOOD')
    } else {
        def write = WriteToExcel(String widget = 'Виджет: Отклонение котлового полезного отпуска', String dzo = 'АО Тываэнерго', 
            String togle = 'Переключатель на : %')
    }
    
    if (WebUI.verifyEqual(e, e1)) {
        println('GOOD')
    } else {
        def write = WriteToExcel(String widget = 'Виджет: Отклонение котлового полезного отпуска', String dzo = 'АО Чеченэнерго', 
            String togle = 'Переключатель на : %')
    }
    
    if (WebUI.verifyEqual(f, f1)) {
        println('GOOD')
    } else {
        def write = WriteToExcel(String widget = 'Виджет: Отклонение котлового полезного отпуска', String dzo = 'Россети Северный Кавказ(ГК)', 
            String togle = 'Переключатель на : %')
    }
    
    if (WebUI.verifyEqual(g, g1)) {
        println('GOOD')
    } else {
        def write = WriteToExcel(String widget = 'Виджет: Отклонение котлового полезного отпуска', String dzo = 'Россети Юг(ГК)', 
            String togle = 'Переключатель на : %')
    }
    
    WebUI.closeBrowser()
}

static def VyruchkaTwoToggleProc6() {
    WebUI.openBrowser('')

    'БЛОК РУКОВОДИТЕЛЕЙ'
    WebUI.navigateToUrl(findTestData('Test Data').getValue(7, 1))

    WebUI.setText(findTestObject('Object Repository/КПО/input__password'), findTestData('Test Data').getValue(6, 1))

    WebUI.setText(findTestObject('Object Repository/КПО/input__username'), findTestData('Test Data').getValue(5, 1))

    WebUI.click(findTestObject('Object Repository/КПО/button_'))

    WebUI.click(findTestObject('КПО/тогл ПРОЦЕНТ'))

    WebUI.click(findTestObject('КПО/фильтр Выручка'))

    WebUI.click(findTestObject('КПО/снять выделения в фильтре Выручка'))

    WebUI.click(findTestObject('КПО/выбрать Объем реалицаии передачи ээ'))

    WebUI.click(findTestObject('КПО/применить в фильтре Выручка'))

    WebUI.click(findTestObject('КПО/фильтр Дата'))

    WebUI.click(findTestObject('КПО/снять выделения в фильтре Дата'))

    WebUI.click(findTestObject('КПО/применить в фильтре Дата'))

    WebUI.click(findTestObject('КПО/фильтр Дата'))

    WebUI.scrollToElement(findTestObject('КПО/2024 год'), 30)

    WebUI.scrollToElement(findTestObject('КПО/скролл Фильтр дата'), 30)

    WebUI.click(findTestObject('КПО/2024 год'), FailureHandling.CONTINUE_ON_FAILURE)

    WebUI.scrollToElement(findTestObject('КПО/раскрыть 1 квартла 2024'), 30)

    WebUI.scrollToElement(findTestObject('КПО/скролл Фильтр дата'), 30)

    WebUI.click(findTestObject('КПО/раскрыть 1 квартла 2024'))

    WebUI.scrollToElement(findTestObject('КПО/Январь 2024'), 30)

    WebUI.scrollToElement(findTestObject('КПО/скролл Фильтр дата'), 30)

    WebUI.click(findTestObject('КПО/Январь 2024'))

    WebUI.click(findTestObject('КПО/Февраль 2024'))

    WebUI.click(findTestObject('КПО/применить в фильтре Дата'))

    WebUI.click(findTestObject('Object Repository/КПО/фильтр ДЗО'))

    WebUI.click(findTestObject('КПО/снять выделения в фильтре ДЗО'))

    WebUI.click(findTestObject('КПО/применить в фильтре ДЗО'))

    WebUI.delay(5)

    String a = WebUI.getText(findTestObject('КПО/Данные с виджета факт1'))

    String a1 = WebUI.getText(findTestObject('КПО/Данные с виджета ПЛАН1'))

    String c = WebUI.getText(findTestObject('КПО/Данные с виджета Отклонения - тогл процент/ПАО Россети'))

    String d = WebUI.getText(findTestObject('КПО/Данные с виджета Отклонения - тогл процент -Выручка Объем реализации подачи ээ/CentIPriv'))

    String e = WebUI.getText(findTestObject('КПО/Данные с виджета Отклонения - тогл процент -Выручка Объем реализации подачи ээ/Centr'))

    String f = WebUI.getText(findTestObject('КПО/Данные с виджета Отклонения - тогл процент -Выручка Объем реализации подачи ээ/chech'))

    String g = WebUI.getText(findTestObject('КПО/Данные с виджета Отклонения - тогл процент -Выручка Объем реализации подачи ээ/kub'))

    String m = WebUI.getText(findTestObject('КПО/Данные с виджета Отклонения - тогл процент -Выручка Объем реализации подачи ээ/len'))

    String n = WebUI.getText(findTestObject('КПО/Данные с виджета Отклонения - тогл процент -Выручка Объем реализации подачи ээ/mo'))

    String o = WebUI.getText(findTestObject('КПО/Данные с виджета Отклонения - тогл процент -Выручка Объем реализации подачи ээ/sevKaz'))

    String p = WebUI.getText(findTestObject('КПО/Данные с виджета Отклонения - тогл процент -Выручка Объем реализации подачи ээ/SevZap'))

    String r = WebUI.getText(findTestObject('КПО/Данные с виджета Отклонения - тогл процент -Выручка Объем реализации подачи ээ/Sibir'))

    String s = WebUI.getText(findTestObject('КПО/Данные с виджета Отклонения - тогл процент -Выручка Объем реализации подачи ээ/Tomsk'))

    String t = WebUI.getText(findTestObject('КПО/Данные с виджета Отклонения - тогл процент -Выручка Объем реализации подачи ээ/Tumen'))

    String u = WebUI.getText(findTestObject('КПО/Данные с виджета Отклонения - тогл процент -Выручка Объем реализации подачи ээ/tuv'))

    String y = WebUI.getText(findTestObject('КПО/Данные с виджета Отклонения - тогл процент -Выручка Объем реализации подачи ээ/Ug'))

    String z = WebUI.getText(findTestObject('КПО/Данные с виджета Отклонения - тогл процент -Выручка Объем реализации подачи ээ/Ural'))

    String x = WebUI.getText(findTestObject('КПО/Данные с виджета Отклонения - тогл процент -Выручка Объем реализации подачи ээ/volg'))

    String w = WebUI.getText(findTestObject('КПО/Данные с виджета Отклонения - тогл процент -Выручка Объем реализации подачи ээ/Yantar'))

    String q = WebUI.getText(findTestObject('КПО/Данные с виджета Отклонения - тогл процент -Выручка Объем реализации подачи ээ/fsk'))

    println(a)

    println(a1)

    println(c)

    'ВЫРУЧКА'
    WebUI.navigateToUrl(findTestData('Test Data').getValue(7, 2))

    WebUI.delay(15)

    if (WebUI.verifyTextPresent('Просьба обратить внимание', true) == true) {
        WebUI.click(findTestObject('КПО для раздела Выручка/закрыть уведомление'))
    }
    
    WebUI.click(findTestObject('КПО для раздела Выручка/Фильтр ДЗО'))

    WebUI.click(findTestObject('КПО для раздела Выручка/снять выделения в ДЗО'))

    WebUI.click(findTestObject('КПО для раздела Выручка/применить в фильтре ДЗО'))

    WebUI.click(findTestObject('КПО для раздела Выручка/ТОГЛ НА ПРОЦЕНТ'))

    WebUI.click(findTestObject('КПО для раздела Выручка/Фильтр Выручка'))

    WebUI.click(findTestObject('КПО для раздела Выручка/dsNATb'))

    WebUI.click(findTestObject('КПО для раздела Выручка/div_EE'))

    WebUI.click(findTestObject('КПО для раздела Выручка/PRIM'))

    WebUI.delay(5)

    String b = WebUI.getText(findTestObject('КПО для раздела Выручка/Данные с виджета факт'))

    String b1 = WebUI.getText(findTestObject('КПО для раздела Выручка/Данные с виджета ПЛАН'))

    String с1 = WebUI.getText(findTestObject('КПО для раздела Выручка/Данные с виджета Отклонения - тогл процент/ПАО Россети'))

    String d1 = WebUI.getText(findTestObject('КПО для раздела Выручка/Данные с виджета Отклонения - тогл процент - ВыручкаОбъем реализации подачи ээ/CentIPriv'))

    String e1 = WebUI.getText(findTestObject('КПО для раздела Выручка/Данные с виджета Отклонения - тогл процент - ВыручкаОбъем реализации подачи ээ/Centr'))

    String f1 = WebUI.getText(findTestObject('КПО для раздела Выручка/Данные с виджета Отклонения - тогл процент - ВыручкаОбъем реализации подачи ээ/chech'))

    String g1 = WebUI.getText(findTestObject('КПО для раздела Выручка/Данные с виджета Отклонения - тогл процент - ВыручкаОбъем реализации подачи ээ/kub'))

    String m1 = WebUI.getText(findTestObject('КПО для раздела Выручка/Данные с виджета Отклонения - тогл процент - ВыручкаОбъем реализации подачи ээ/len'))

    String n1 = WebUI.getText(findTestObject('КПО для раздела Выручка/Данные с виджета Отклонения - тогл процент - ВыручкаОбъем реализации подачи ээ/mo'))

    String o1 = WebUI.getText(findTestObject('КПО для раздела Выручка/Данные с виджета Отклонения - тогл процент - ВыручкаОбъем реализации подачи ээ/sevKaz'))

    String p1 = WebUI.getText(findTestObject('КПО для раздела Выручка/Данные с виджета Отклонения - тогл процент - ВыручкаОбъем реализации подачи ээ/SevZap'))

    String r1 = WebUI.getText(findTestObject('КПО для раздела Выручка/Данные с виджета Отклонения - тогл процент - ВыручкаОбъем реализации подачи ээ/Sibir'))

    String s1 = WebUI.getText(findTestObject('КПО для раздела Выручка/Данные с виджета Отклонения - тогл процент - ВыручкаОбъем реализации подачи ээ/Tomsk'))

    String t1 = WebUI.getText(findTestObject('КПО для раздела Выручка/Данные с виджета Отклонения - тогл процент - ВыручкаОбъем реализации подачи ээ/Tumen'))

    String u1 = WebUI.getText(findTestObject('КПО для раздела Выручка/Данные с виджета Отклонения - тогл процент - ВыручкаОбъем реализации подачи ээ/tuv'))

    String y1 = WebUI.getText(findTestObject('КПО для раздела Выручка/Данные с виджета Отклонения - тогл процент - ВыручкаОбъем реализации подачи ээ/Ug'))

    String z1 = WebUI.getText(findTestObject('КПО для раздела Выручка/Данные с виджета Отклонения - тогл процент - ВыручкаОбъем реализации подачи ээ/Ural'))

    String x1 = WebUI.getText(findTestObject('КПО для раздела Выручка/Данные с виджета Отклонения - тогл процент - ВыручкаОбъем реализации подачи ээ/volg'))

    String w1 = WebUI.getText(findTestObject('КПО для раздела Выручка/Данные с виджета Отклонения - тогл процент - ВыручкаОбъем реализации подачи ээ/Yantar'))

    String q1 = WebUI.getText(findTestObject('КПО для раздела Выручка/Данные с виджета Отклонения - тогл процент - ВыручкаОбъем реализации подачи ээ/fsk vuruch'))

    if (WebUI.verifyEqual(a, b) && WebUI.verifyEqual(a1, b1)) {
        println('GOOD')
    } else {
        def write = WriteToExcel(String widget = 'Виджет: ВЫполнение плановых показателй', String dzo = 'ПАО Россети', String togle = 'Переключатель на : %')
    }
    
    if (WebUI.verifyEqual(c, с1)) {
        println('GOOD')
    } else {
        def write = WriteToExcel(String widget = 'Виджет: Отклонение котлового полезного отпуска', String dzo = 'ПАО Россети', 
            String togle = 'Переключатель на : %')
    }
    
    if (WebUI.verifyEqual(d, d1)) {
        println('GOOD')
    } else {
        def write = WriteToExcel(String widget = 'Виджет: Отклонение котлового полезного отпуска', String dzo = 'Центр и Приволжье', 
            String togle = 'Переключатель на : %')
    }
    
    if (WebUI.verifyEqual(e, e1)) {
        println('GOOD')
    } else {
        def write = WriteToExcel(String widget = 'Виджет: Отклонение котлового полезного отпуска', String dzo = 'Центр', 
            String togle = 'Переключатель на : %')
    }
    
    if (WebUI.verifyEqual(f, f1)) {
        println('GOOD')
    } else {
        def write = WriteToExcel(String widget = 'Виджет: Отклонение котлового полезного отпуска', String dzo = 'Чеченэнерго', 
            String togle = 'Переключатель на : %')
    }
    
    if (WebUI.verifyEqual(g, g1)) {
        println('GOOD')
    } else {
        def write = WriteToExcel(String widget = 'Виджет: Отклонение котлового полезного отпуска', String dzo = 'Россети Кубань', 
            String togle = 'Переключатель на : %')
    }
    
    if (WebUI.verifyEqual(m, m1)) {
        println('GOOD')
    } else {
        def write = WriteToExcel(String widget = 'Виджет: Отклонение котлового полезного отпуска', String dzo = 'Россети Ленэнерго(ГК)', 
            String togle = 'Переключатель на : %')
    }
    
    if (WebUI.verifyEqual(n, n1)) {
        println('GOOD')
    } else {
        def write = WriteToExcel(String widget = 'Виджет: Отклонение котлового полезного отпуска', String dzo = 'Россети Москвоский регион', 
            String togle = 'Переключатель на : %')
    }
    
    if (WebUI.verifyEqual(o, o1)) {
        println('GOOD')
    } else {
        def write = WriteToExcel(String widget = 'Виджет: Отклонение котлового полезного отпуска', String dzo = 'Россети Северный Кавказ(ГК)', 
            String togle = 'Переключатель на : %')
    }
    
    if (WebUI.verifyEqual(p, p1)) {
        println('GOOD')
    } else {
        def write = WriteToExcel(String widget = 'Виджет: Отклонение котлового полезного отпуска', String dzo = 'Россети Северо-Запад', 
            String togle = 'Переключатель на : %')
    }
    
    if (WebUI.verifyEqual(r, r1)) {
        println('GOOD')
    } else {
        def write = WriteToExcel(String widget = 'Виджет: Отклонение котлового полезного отпуска', String dzo = 'Россети Сибирь(ГК)', 
            String togle = 'Переключатель на : %')
    }
    
    if (WebUI.verifyEqual(s, s1)) {
        println('GOOD')
    } else {
        def write = WriteToExcel(String widget = 'Виджет: Отклонение котлового полезного отпуска', String dzo = 'Россети Томск', 
            String togle = 'Переключатель на : %')
    }
    
    if (WebUI.verifyEqual(t, t1)) {
        println('GOOD')
    } else {
        def write = WriteToExcel(String widget = 'Виджет: Отклонение котлового полезного отпуска', String dzo = 'Россети Тюмень', 
            String togle = 'Переключатель на : %')
    }
    
    if (WebUI.verifyEqual(u, u1)) {
        println('GOOD')
    } else {
        def write = WriteToExcel(String widget = 'Виджет: Отклонение котлового полезного отпуска', String dzo = 'Тываэнерго', 
            String togle = 'Переключатель на : %')
    }
    
    if (WebUI.verifyEqual(y, y1)) {
        println('GOOD')
    } else {
        def write = WriteToExcel(String widget = 'Виджет: Отклонение котлового полезного отпуска', String dzo = 'Россети Юг(ГК)', 
            String togle = 'Переключатель на : %')
    }
    
    if (WebUI.verifyEqual(z, z1)) {
        println('GOOD')
    } else {
        def write = WriteToExcel(String widget = 'Виджет: Отклонение котлового полезного отпуска', String dzo = 'Россети Урал(ГК)', 
            String togle = 'Переключатель на : %')
    }
    
    if (WebUI.verifyEqual(x, x1)) {
        println('GOOD')
    } else {
        def write = WriteToExcel(String widget = 'Виджет: Отклонение котлового полезного отпуска', String dzo = 'Россети Волга', 
            String togle = 'Переключатель на : %')
    }
    
    if (WebUI.verifyEqual(w, w1)) {
        println('GOOD')
    } else {
        def write = WriteToExcel(String widget = 'Виджет: Отклонение котлового полезного отпуска', String dzo = 'Россети Янтарь', 
            String togle = 'Переключатель на : %')
    }
    
    if (WebUI.verifyEqual(q, q1)) {
        println('GOOD')
    } else {
        def write = WriteToExcel(String widget = 'Виджет: Отклонение котлового полезного отпуска', String dzo = 'Россети ФСК ЕЭС', 
            String togle = 'Переключатель на : %')
    }
    
    WebUI.closeBrowser()
}

static def WriteToExcel(def widget, def dzo, def togle) {
    String sheetName = 'List1'

    def data = findTestData('Test Data')

    int n = data.getRowNumbers() + 1

    String dZO = WebUI.getText(findTestObject('КПО для раздела Выручка/Фильтр ДЗО'))

    String year = WebUI.getText(findTestObject('КПО для раздела Выручка/фильтр Дата'))

    println(year)

    String dashboardName = 'Котловой полезный отпуск'

    String page = 'Данные не  совпадают'

    Date d = new Date()

    SimpleDateFormat format1

    format1 = new SimpleDateFormat('dd.MM.yyyy')

    String date = format1.format(d)

    println(date)

    String fltr = WebUI.getText(findTestObject('КПО для раздела Выручка/Фильтр Выручка'))

    println(fltr)

    String filtrVrch = 'Фильтр Выручка: ' + fltr

    println(filtrVrch)

    def workbook01 = ExcelKeywords.getWorkbook(GlobalVariable.excelFilePath)

    def sheet01 = ExcelKeywords.getExcelSheet(workbook01, sheetName)

    ExcelKeywords.setValueToCellByIndex(sheet01, n, 0, dashboardName)

    ExcelKeywords.setValueToCellByIndex(sheet01, n, 1, dZO)

    ExcelKeywords.setValueToCellByIndex(sheet01, n, 2, dzo)

    ExcelKeywords.setValueToCellByIndex(sheet01, n, 3, filtrVrch)

    ExcelKeywords.setValueToCellByIndex(sheet01, n, 4, year)

    ExcelKeywords.setValueToCellByIndex(sheet01, n, 5, page)

    ExcelKeywords.setValueToCellByIndex(sheet01, n, 6, widget)

    ExcelKeywords.setValueToCellByIndex(sheet01, n, 7, togle)

    ExcelKeywords.setValueToCellByIndex(sheet01, n, 8, date)

    n = (n + 1)

    ExcelKeywords.saveWorkbook(GlobalVariable.excelFilePath, workbook01)

    WebUI.acceptAlert()
}

