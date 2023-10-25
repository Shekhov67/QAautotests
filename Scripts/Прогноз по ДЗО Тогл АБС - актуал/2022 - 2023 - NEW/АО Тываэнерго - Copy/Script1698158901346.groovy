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
import java.util.HashMap as HashMap
import java.util.Map as Map
import java.util.Date as Date
import java.text.SimpleDateFormat as SimpleDateFormat

def openBrowser = OpenBrowser()

WebUI.click(findTestObject('Прогноз по ДЗО/тогл на МЛН'))

WebUI.click(findTestObject('Прогноз по ДЗО/фильтр ДЗО'))

WebUI.click(findTestObject('Прогноз по ДЗО/снять выделенные в ФИЛЬТРЕ дзо'))

WebUI.click(findTestObject('Прогноз по ДЗО/применить в фильтре ДЗО'))

WebUI.click(findTestObject('Прогноз по ДЗО/фильтр ДЗО'))

WebUI.doubleClick(findTestObject('Прогноз по ДЗО/ПАО Россети'))

WebUI.click(findTestObject('Прогноз по ДЗО/РаспредКомплекс'))

WebUI.click(findTestObject('Прогноз по ДЗО/выбрать АО Тываэнерго'))

WebUI.click(findTestObject('Прогноз по ДЗО/применить в фильтре ДЗО'))

def selectDate = SelectDate()

WebUI.closeBrowser()

openBrowser = OpenBrowser()

WebUI.click(findTestObject('Прогноз по ДЗО/тогл на МЛН'))

WebUI.click(findTestObject('Прогноз по ДЗО/фильтр ДЗО'))

WebUI.click(findTestObject('Прогноз по ДЗО/снять выделенные в ФИЛЬТРЕ дзо'))

WebUI.click(findTestObject('Прогноз по ДЗО/применить в фильтре ДЗО'))

WebUI.click(findTestObject('Прогноз по ДЗО/фильтр ДЗО'))

WebUI.doubleClick(findTestObject('Прогноз по ДЗО/ПАО Россети'))

WebUI.click(findTestObject('Прогноз по ДЗО/РаспредКомплекс'))

WebUI.click(findTestObject('Прогноз по ДЗО/АО Тываэнерго'))

WebUI.click(findTestObject('Прогноз по ДЗО/АО Тываэнерго(2 уровень)'))

WebUI.click(findTestObject('Прогноз по ДЗО/применить в фильтре ДЗО'))

selectDate = SelectDate()

WebUI.closeBrowser( //////////////Запись даты
    ///////////////
    )

static def SelectDate() {
    WebUI.click(findTestObject('Прогноз по ДЗО/фильтр Год'))

    WebUI.click(findTestObject('Прогноз по ДЗО/выбрать 2022'))

    WebUI.callTestCase(findTestCase('Тестовые тесты/Фильтр дата в прогнозе по ДЗО'), [:], FailureHandling.CONTINUE_ON_FAILURE)

    WebUI.click(findTestObject('Прогноз по ДЗО/фильтр Год'))

    WebUI.click(findTestObject('Прогноз по ДЗО/выбрать 2023'))

    WebUI.callTestCase(findTestCase('Тестовые тесты/Фильтр дата в прогнозе по ДЗО'), [:], FailureHandling.CONTINUE_ON_FAILURE)
}

static def OpenBrowser() {
    WebUI.openBrowser('')

    WebUI.navigateToUrl(findTestData('Test Data').getValue(3, 1))

    WebUI.setText(findTestObject('Object Repository/Прогноз по ДЗО/input__password'), findTestData('Test Data').getValue(
            2, 1))

    WebUI.setText(findTestObject('Object Repository/Прогноз по ДЗО/input__username'), findTestData('Test Data').getValue(
            1, 1))

    WebUI.click(findTestObject('Object Repository/Прогноз по ДЗО/button_'))
}

