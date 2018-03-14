package com.pack.base;

import java.util.concurrent.TimeUnit;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
//import org.openqa.selenium.ie.InternetExplorerDriver;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.testng.annotations.AfterClass;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.Parameters;

public class base 
{

	public WebDriver driver;
	public static String driverPath = "./BrowserDrivers/chromedriver.exe";
	public static String driverPathIE = "./BrowserDrivers/IEDriverServer.exe";

	public WebDriver getDriver() 
	{
		return driver;
	}

	private void setDriver(String browserType, String appURL) 
	
	{
		switch (browserType) {
		case "chrome":
			driver = initChromeDriver(appURL);
			break;
		case "IE":
			driver = initInternetExplorerDriver(appURL);
			break;
		case "firefox":
			driver = initFirefoxDriver(appURL);
			break;
		default:
			System.out.println("browser : " + browserType
					+ " is invalid, Launching Firefox as browser of choice..");
			driver = initFirefoxDriver(appURL);
			driver.manage().timeouts().implicitlyWait(500, TimeUnit.SECONDS);
		}
	}
	private static WebDriver initChromeDriver(String appURL) {
		System.out.println("Launching google chrome with new profile..");
		System.setProperty("webdriver.chrome.driver", driverPath);
		WebDriver driver = new ChromeDriver();
		driver.manage().window().maximize();
		driver.manage().timeouts().implicitlyWait(400, TimeUnit.SECONDS);
		driver.navigate().to(appURL);
		return driver;
	}
	private static WebDriver initInternetExplorerDriver(String appURL) {
		System.out.println("Launching Internet Explorer with new profile..");
		System.setProperty("webdriver.ie.driver", driverPathIE);
		WebDriver driver = new InternetExplorerDriver();
		driver.manage().window().maximize();
		driver.manage().timeouts().implicitlyWait(400, TimeUnit.SECONDS);
		driver.navigate().to(appURL);
		return driver;
	}

	private static WebDriver initFirefoxDriver(String appURL) {
		System.out.println("Launching Firefox browser..");
		WebDriver driver = new FirefoxDriver();
		driver.manage().window().maximize();
		driver.manage().timeouts().implicitlyWait(400, TimeUnit.SECONDS);
		driver.navigate().to(appURL);
		return driver;
	}

	@Parameters({ "browserType", "appURL" })
	@BeforeClass
	public void initializeTestBaseSetup(String browserType, String appURL) 
	{
		try 
		{
			setDriver(browserType, appURL);

		} 
		catch (Exception e)
		{
			System.out.println("Error....." + e.getStackTrace());
		}
	}
	
	@AfterClass
	public void tearDown() throws InterruptedException 
	{
		Thread.sleep(900000);
		driver.quit();
	}
}
	