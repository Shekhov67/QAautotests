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
    WebUI.click(findTestObject('Общие в сеть/Фильтр Дата'))

    WebUI.click(findTestObject('Общие в сеть/Снять выделения в фильтре Дата'))

    WebUI.click(findTestObject('Общие в сеть/Применить в фильтре Дата'))

    WebUI.click(findTestObject('Общие в сеть/Фильтр Дата'))

    WebUI.scrollToElement(findTestObject('Отпуск в сеть сверка/2023 год'), 30)

    WebUI.scrollToElement(findTestObject('Общие в сеть/скролл до фильтра дата'), 30)

    WebUI.click(findTestObject('Отпуск в сеть сверка/2023 год'))

    WebUI.scrollToElement(findTestObject('Отпуск в сеть сверка/4 квартал 2023'), 30)

    WebUI.scrollToElement(findTestObject('Общие в сеть/скролл до фильтра дата'), 30)

    WebUI.click(findTestObject('Отпуск в сеть сверка/выбрать 1 квартал'))

    WebUI.click(findTestObject('Отпуск в сеть(виджеты)/выбрать 2 квартал 2023'))

    WebUI.click(findTestObject('Отпуск в сеть сверка/раскрыть 3 квартал 2023'))

    WebUI.click(findTestObject('Отпуск в сеть сверка/выбрать июль'))

    WebUI.scrollToElement(findTestObject('Отпуск в сеть(виджеты)/Август'), 30)

    WebUI.scrollToElement(findTestObject('Отпуск в сеть(виджеты)/скрол'), 30)

    WebUI.click(findTestObject('Отпуск в сеть(виджеты)/Август'), FailureHandling.CONTINUE_ON_FAILURE)

    WebUI.scrollToElement(findTestObject('Отпуск в сеть(виджеты)/Сентябрь'), 30)

    WebUI.scrollToElement(findTestObject('Отпуск в сеть(виджеты)/скрол'), 30)

    WebUI.click(findTestObject('Отпуск в сеть(виджеты)/Сентябрь'), FailureHandling.CONTINUE_ON_FAILURE)

    WebUI.click(findTestObject('Общие в сеть/Применить в фильтре Дата'))
}

