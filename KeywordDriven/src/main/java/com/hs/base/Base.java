package com.hs.base;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Properties;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.edge.EdgeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;

public class Base {
	public static WebDriver driver;
	public static Properties properties;
	public static String configFilePath = System.getProperty("user.dir")+"\\src\\main\\java\\com\\hs\\config\\config.properties";
	
	
	
	public static WebDriver initDriver(String driverName) {
		switch (driverName) {
		case "chrome":
			ChromeOptions options = new ChromeOptions();
			if(properties.getProperty("headless").equals("yes")) {
				options.addArguments("--headless");
			}
			driver = new ChromeDriver(options);
			break;
		case "firefox":
			driver = new FirefoxDriver();
			break;
		case "edge":
			driver = new EdgeDriver();
			break;
		default:
			driver = null;
			break;
		}
		
		return driver;
	}
	
	public static Properties intiProperties() {
		properties = new Properties();
		try {
			FileInputStream fileInputStream = new FileInputStream(configFilePath);
			try {
				properties.load(fileInputStream);
			} catch (IOException e) {
				e.printStackTrace();
			}
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		}
		return properties;
	}

}
