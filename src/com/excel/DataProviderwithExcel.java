package com.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.concurrent.TimeUnit;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

public class DataProviderwithExcel {

	static WebDriver driver;

	@BeforeClass
	public void browserLaunch() {
		System.setProperty("webdriver.chrome.driver", "C:/Users/sarav/Downloads/chromedriver_win32/chromedriver.exe");
		driver = new ChromeDriver();
		driver.manage().window().maximize();
		driver.get("http://leaftaps.com/opentaps/control/main");
		driver.manage().timeouts().implicitlyWait(50, TimeUnit.SECONDS);
	}

	@DataProvider
	public Object[][] getLoginData() throws IOException {

		File loc = new File("C:\\Users\\sarav\\Downloads\\DataSheet.xlsx");

		FileInputStream stream = new FileInputStream(loc);
		Workbook w = new XSSFWorkbook(stream);
		Sheet s = w.getSheet("Data");

		Object[][] data = new Object[s.getLastRowNum()][s.getRow(0).getLastCellNum()];

		for (int i = 0; i < s.getLastRowNum(); i++) {
			for (int j = 0; j < s.getRow(0).getLastCellNum(); j++) {

				data[i][j] = s.getRow(i + 1).getCell(j).toString();

			}

		}
		return data;

	}

	@Test(dataProvider = "getLoginData")
	public static void getLoginTest(String username, String password) {

		driver.findElement(By.id("username")).sendKeys(username);
		driver.findElement(By.id("password")).sendKeys(password);
		
		driver.findElement(By.xpath("//*[@type='submit']")).click(); 

	}
}
