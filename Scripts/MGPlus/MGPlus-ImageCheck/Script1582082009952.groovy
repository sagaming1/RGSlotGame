import org.openqa.selenium.By
import org.openqa.selenium.WebDriver
import org.openqa.selenium.WebElement
import com.kms.katalon.core.annotation.Keyword
import com.kms.katalon.core.webui.driver.DriverFactory

import com.kms.katalon.core.logging.KeywordLogger as KeywordLogger
import com.kms.katalon.core.testobject.TestObject as TestObject
import com.kms.katalon.core.webui.keyword.WebUiBuiltInKeywords as WebUI

KeywordLogger log = new KeywordLogger()

WebUI.openBrowser('https://rg-play.com/SlotGame/MGPlus')

String path = '//img[@class=\'img-gameSlot\']'

TestObject product = new TestObject()

product = WebUI.modifyObjectProperty(product, 'xpath', 'equals', path, true)

WebUI.waitForElementVisible(product, 1000)

List<String> vs = getWebElementsAsList(path)

int success = 0

int failure = 0

for (i = 0; i < vs.size(); i++) {
    //	println(vs[i].getSize().getWidth())
    if ((vs[i]).getSize().getWidth() < 50) {
        
		String path2 = "(//span[@class='txt-gameName'])" + "[" + (i+1) + "]"
		
		TestObject product2 = new TestObject()
		
		product2 = WebUI.modifyObjectProperty(product, 'xpath', 'equals', path2, true)
				
        log.logFailed((WebUI.getText(product2) + 'Image' + ' load failed'))

        failure = (failure + 1)
    } else {
        success = (success + 1)
    }
}

println((('上架遊戲圖片正常顯示Success:' + success) + ' 上架遊戲圖片未正常顯示Failure:') + failure)


List<WebElement> getWebElementsAsList(String xpath4elements) {
	WebDriver webDriver = DriverFactory.getWebDriver()
	List<WebElement> elements = webDriver.findElements(By.xpath(xpath4elements))
	return elements
}

