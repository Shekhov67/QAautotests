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

def selectDate = SelectDate()

static def SelectDate() {
    WebUI.click(findTestObject('Отпуск в сеть(виджеты)/фильтр Дата'))

    WebUI.click(findTestObject('Отпуск в сеть(виджеты)/снять выделения в фильтре Дата'))

    WebUI.click(findTestObject('Отпуск в сеть(виджеты)/применить в фильтре Дата'))

    WebUI.click(findTestObject('Отпуск в сеть(виджеты)/фильтр Дата'))

    WebUI.scrollToElement(findTestObject('Отпуск в сеть сверка/2024 год'), 30)

    WebUI.scrollToElement(findTestObject('Отпуск в сеть(виджеты)/скрол'), 30)

    WebUI.click(findTestObject('Отпуск в сеть сверка/2024 год'))

    WebUI.scrollToElement(findTestObject('Отпуск в сеть сверка/1 квартал'), 30)

    WebUI.scrollToElement(findTestObject('Отпуск в сеть(виджеты)/скрол'), 30)

    WebUI.click(findTestObject('Отпуск в сеть сверка/выбрать 1 квартал 2024'), FailureHandling.CONTINUE_ON_FAILURE)

    WebUI.click(findTestObject('Отпуск в сеть сверка/выбрать 2 квартал 2024'), FailureHandling.CONTINUE_ON_FAILURE)

    WebUI.scrollToElement(findTestObject('Отпуск в сеть сверка/3 квартал 2024 раскрыть'), 30)

    WebUI.scrollToElement(findTestObject('Отпуск в сеть(виджеты)/скрол'), 30)

    WebUI.click(findTestObject('Отпуск в сеть(виджеты)/выбрать 3 квартал 2024'), FailureHandling.CONTINUE_ON_FAILURE)

    WebUI.click(findTestObject('Отпуск в сеть(виджеты)/4 квартал 2024'), FailureHandling.CONTINUE_ON_FAILURE)

    WebUI.scrollToElement(findTestObject('Отпуск в сеть(виджеты)/Октябрь'), 30)

    WebUI.scrollToElement(findTestObject('Отпуск в сеть(виджеты)/скрол'), 30)

    WebUI.click(findTestObject('Отпуск в сеть(виджеты)/Октябрь'), FailureHandling.CONTINUE_ON_FAILURE)

    WebUI.click(findTestObject('Отпуск в сеть(виджеты)/применить в фильтре Дата'))
}

