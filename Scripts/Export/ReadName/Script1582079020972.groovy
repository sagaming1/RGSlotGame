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
import com.kms.katalon.core.testobject.TestObject as TestObject
import com.kms.katalon.core.webservice.keyword.WSBuiltInKeywords as WS
import com.kms.katalon.core.webui.keyword.WebUiBuiltInKeywords as WebUI
import com.kms.katalon.core.windows.keyword.WindowsBuiltinKeywords as Windows
import internal.GlobalVariable as GlobalVariable
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.lang.String
import com.kms.katalon.core.testdata.reader.ExcelFactory

//WebUI.openBrowser('https://rg-play.com/SlotGame/PNG')

//println(WebUI.getText(findTestObject('New Test Object')))

String path = "//span[@class='txt-gameName']"
TestObject product = new TestObject()

product = WebUI.modifyObjectProperty(product, 'xpath', 'equals', path, true)

WebUI.openBrowser('https://rg-play.com/SlotGame/MGPlus')

println(WebUI.getText(product))

List<String> cs = CustomKeywords.'FindElementsByXPath.getElementContentsAsList'(path)
List<String> vs = []

//println(cs.size())
//println(cs)
FileInputStream file = new FileInputStream (new File("D:\\Workspace\\2020\\0218\\MGPlus-Data.xlsx"))
XSSFWorkbook workbook = new XSSFWorkbook(file);
XSSFSheet sheet = workbook.getSheetAt(0);
Object excelFile = ExcelFactory.getExcelDataWithDefaultSheet("D:\\Workspace\\2020\\0218\\MGPlus-Data.xlsx","sheet1", true)

for (i =0 ; i < cs.size() ; i++){
//	println(cs[i])
	sheet.createRow(i+1)
	sheet.getRow(i+1).createCell(0).setCellValue(cs[i])
}

file.close();
FileOutputStream outFile =new FileOutputStream(new File("D:\\Workspace\\2020\\0218\\MGPlus-Data.xlsx"));
workbook.write(outFile);
outFile.close();