package pageobjects;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;

public class RegisterPage {
	public static WebElement Gender(WebDriver driver) {
		return driver.findElement(By.xpath("//*[@id=\"gender-female\"]"));
		
	}

}
