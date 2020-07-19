package com.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.concurrent.TimeUnit;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;

import atu.testrecorder.ATUTestRecorder;
import atu.testrecorder.exceptions.ATUTestRecorderException;

public class ExcelInput {
	public static void main(String[] args) throws Throwable {
		
		System.setProperty("webdriver.chrome.driver",
				"C:/Users/sarav/Downloads/chromedriver_win32/chromedriver.exe");
		WebDriver driver = new ChromeDriver();
		
		SimpleDateFormat simple=new SimpleDateFormat("yy-MM-dd HH-mm-ss");
		Date dat=new Date();
		ATUTestRecorder recorder=new ATUTestRecorder("C:\\Users\\sarav\\OneDrive\\Documents\\ATU Record files","Testvideo"+simple.format(dat), false);
		recorder.start();
		
		driver.manage().window().maximize();
		driver.get("http://www.adactin.com/HotelApp/index.php");
		driver.manage().timeouts().pageLoadTimeout(10, TimeUnit.SECONDS);
		driver.manage().timeouts().implicitlyWait(15, TimeUnit.SECONDS);

		File loc= new File("C:\\Users\\sarav\\Downloads\\DataSheet.xlsx");
		   
		   FileInputStream stream =new FileInputStream(loc);
		   Workbook w =new XSSFWorkbook(stream);
		   Sheet s=w.getSheet("Data");
		   
		for (int i = 1; i < s.getPhysicalNumberOfRows(); i++) {
			Row r = s.getRow(i);
			for (int j = 0; j < r.getPhysicalNumberOfCells(); j++) {
				Cell c = r.getCell(j);

				int type = c.getCellType();
				if (type == 1) {
					String stringCellValue = c.getStringCellValue();
					System.out.println(stringCellValue);
					if(j == 0) {
						driver.findElement(By.id("username")).sendKeys(stringCellValue);
					} else {
						driver.findElement(By.id("password")).sendKeys(stringCellValue);
					}
				}
				if (type == 0) {
					if (DateUtil.isCellDateFormatted(c)) {
						Date dateCell = c.getDateCellValue();
						SimpleDateFormat fr = new SimpleDateFormat("DD-MM-YYYY");
						String date = fr.format(dateCell);
						System.out.println(date);
						if(j == 0) {
							driver.findElement(By.id("username")).sendKeys(date);
						} else {
							driver.findElement(By.id("password")).sendKeys(date);
						}

					} else {
						double d = c.getNumericCellValue();
						long l = (long) d;
						String numeric = String.valueOf(l);
						System.out.println(numeric);
						
						if(j == 0) {
							driver.findElement(By.id("username")).sendKeys(numeric);
						} else {
							driver.findElement(By.id("password")).sendKeys(numeric);
						}
					}
				}

			}
			
			driver.findElement(By.id("login")).click();
				

		}
		driver.quit();
		recorder.stop();
		
	}

}
