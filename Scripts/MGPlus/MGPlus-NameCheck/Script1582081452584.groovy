import org.apache.poi.xssf.usermodel.XSSFSheet as XSSFSheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook as XSSFWorkbook
import org.openqa.selenium.By
import org.openqa.selenium.WebDriver

import com.kms.katalon.core.configuration.RunConfiguration as RunConfiguration
import com.kms.katalon.core.logging.KeywordLogger as KeywordLogger
import com.kms.katalon.core.testdata.reader.ExcelFactory as ExcelFactory
import com.kms.katalon.core.testobject.TestObject as TestObject
import com.kms.katalon.core.webui.driver.DriverFactory
import com.kms.katalon.core.webui.keyword.WebUiBuiltInKeywords as WebUI

String dirpath = RunConfiguration.getProjectDir()

KeywordLogger log = new KeywordLogger()

//third parameter means if you want the first row as your header or column name.
//In your case, it should be true.
FileInputStream file = new FileInputStream(new File(dirpath + '//Data//MGPlus-Data.xlsx'))

XSSFWorkbook workbook = new XSSFWorkbook(file)

XSSFSheet sheet = workbook.getSheetAt(0)

Object excelFile = ExcelFactory.getExcelDataWithDefaultSheet(dirpath + '//Data//MGPlus-Data.xlsx', 'sheet1', true)

//println(excelFile.getRowNumbers()) //Get total rows of the test data
//https://docs.katalon.com/javadoc/com/kms/katalon/core/testdata/reader/SheetPOI.html
WebUI.openBrowser('https://rg-play.com/SlotGame/MGPlus')

String path = "//span[@class='txt-gameName']"
TestObject product = new TestObject()
product = WebUI.modifyObjectProperty(product, 'xpath', 'equals', path, true)
WebUI.waitForElementVisible(product, 1000)

int success = 0
int failure = 0

WebDriver Driver = DriverFactory.getWebDriver()
println(sheet.size())
for (int i = 0; i < (sheet.size() - 1); i++) {
	try{
		if(Driver.findElement(By.xpath("//*[text()='"+ excelFile.internallyGetValue(0, i) + "']")).isDisplayed()){
		success = success+1
		}
		
	   }
	catch(org.openqa.selenium.NoSuchElementException e){
		failure = failure+1
		log.logFailed(excelFile.internallyGetValue(0, i) + ' is not present')
		}

}

println((('成功上架Success:' + success) + ' 應上架未上架Failure:') + failure)

//關閉EXCEL********
file.close()

FileOutputStream outFile = new FileOutputStream(new File(dirpath + '//Data//MGPlus-Data.xlsx'))

workbook.write(outFile)

outFile.close()

