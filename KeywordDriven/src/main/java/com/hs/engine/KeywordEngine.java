package com.hs.engine;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Properties;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;

import com.hs.base.Base;

public class KeywordEngine extends Base{
	
	public static Workbook workbook;
	public static Sheet sheet;
	
	public static WebDriver driver;
	public static Properties properties = Base.intiProperties();
	
	public final static String scenarioSheet = System.getProperty("user.dir")+"\\src\\main\\java\\com\\hs\\scenarios\\Login.xlsx";
	
	public static void startExecution(String sheetName)  {
		String locatorName = null;
		String locatorValue = null;
		
		FileInputStream fis = null;
		
		try {
			fis = new FileInputStream(scenarioSheet);
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		}
		
		try {
			workbook = WorkbookFactory.create(fis);
		} catch (EncryptedDocumentException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		int k = 0;
		sheet = workbook.getSheet(sheetName);
		
		for(int i=0;i<sheet.getLastRowNum();i++) {
			String locatorColumnValue = sheet.getRow(i+1).getCell(k+1).toString().trim();  
			if(!locatorColumnValue.equals("NA")) {
				locatorName = locatorColumnValue.split("=")[0].trim();
				locatorValue = locatorColumnValue.split("=")[1].trim();
			}
			String action = sheet.getRow(i+1).getCell(k+2).toString().trim();  
			String value = sheet.getRow(i+1).getCell(k+3).toString().trim();  
			
			switch (action) {
			
			case "open browser":
				if(value.isEmpty() || value.equals("NA")) {
					driver = Base.initDriver(properties.getProperty("browser"));
				}else {
					driver = Base.initDriver(value);
				}
				break;
				
			case "launch url":
				if(value.isEmpty() || value.equals("NA")) {
					driver.get(properties.getProperty("url"));
				}else {
					driver.get(value);
				}
				break;
			
			case "quit":
				driver.quit();
				break;
			
			default:
				break;
			}
			
			WebElement element ;
			
			switch (locatorName) {
			case "id":
				element = driver.findElement(By.id(locatorValue));
				if(action.equalsIgnoreCase("sendkeys")) {
					element.clear();
					element.sendKeys(value);
				}else if(action.equalsIgnoreCase("click"))
					element.click();
				locatorName = null;
				break;
			case "name":
				element = driver.findElement(By.name(locatorValue));
				if(action.equalsIgnoreCase("sendkeys")) {
					element.clear();
					element.sendKeys(value);
				}else if(action.equalsIgnoreCase("click"))
					element.click();
				locatorName = null;
				break;

			default:
				break;
			}
		}
	}

}
