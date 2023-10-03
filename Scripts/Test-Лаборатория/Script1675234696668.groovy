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

WebUI.openBrowser('')

WebUI.navigateToUrl(findTestData('Test Data').getValue(3, 1))

WebUI.setText(findTestObject('Object Repository/Прогноз по ДЗО/input__password'), findTestData('Test Data').getValue(2, 
        1))

WebUI.setText(findTestObject('Object Repository/Прогноз по ДЗО/input__password'), findTestData('Test Data').getValue(2, 
        1))

WebUI.setText(findTestObject('Object Repository/Прогноз по ДЗО/input__username'), findTestData('Test Data').getValue(1, 
        1))

WebUI.click(findTestObject('Object Repository/Прогноз по ДЗО/button_'))

WebUI.setText(findTestObject('Object Repository/Прогноз по ДЗО/input__password'), findTestData('Test Data').getValue(2, 
        1))

WebUI.click(findTestObject('Прогноз по ДЗО/Тюмень Исполнительный аппарат'))

WebUI.click(findTestObject('Прогноз по ДЗО/выбрать 2023'))

/*def togle1 = WebUI.getCSSValue(findTestObject('Прогноз по ДЗО/тогл процента'), 'background')

println(togle1)

def togle2 = WebUI.getCSSValue(findTestObject('Прогноз по ДЗО/тогл абс'), 'background')

println(togle2)*/
'sqwded22222222222222222222qwe'
if (WebUI.getCSSValue(findTestObject('Прогноз по ДЗО/тогл процента'), 'background') == 'rgba(0, 0, 0, 0) url("https://bi.rosseti.digital/corelogic/api/query/image?fileGuid=96a494b6a9c4452397eb2c87d1c586b9&access_token=NoAuth") no-repeat scroll 50% 50% / contain padding-box border-box') {
    WebUI.click(findTestObject('Прогноз по ДЗО/тогл процента'))

    String typeOfTogle = 'преключение на абс'

    println(typeOfTogle)
} else if (WebUI.getCSSValue(findTestObject('Прогноз по ДЗО/тогл абс'), 'background') == 'rgba(0, 0, 0, 0) url("https://bi.rosseti.digital/corelogic/api/query/image?fileGuid=be8bb87696b04b9b83060a398b6bd4df&access_token=NoAuth") no-repeat scroll 50% 50% / contain padding-box border-box') {
    WebUI.click(findTestObject('Прогноз по ДЗО/тогл абс'))

    String typeOfTogle = 'преключение на проценты.'

    println(typeOfTogle)
}

