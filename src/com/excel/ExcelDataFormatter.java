package com.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.concurrent.TimeUnit;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

import atu.testrecorder.ATUTestRecorder;
import atu.testrecorder.exceptions.ATUTestRecorderException;

public class ExcelDataFormatter {

	public static void main(String[] args) throws IOException, ATUTestRecorderException {

		System.setProperty("webdriver.chrome.driver",
				"C:/Users/sarav/Downloads/chromedriver_win32/chromedriver.exe");
		WebDriver driver = new ChromeDriver();

		SimpleDateFormat simple = new SimpleDateFormat("yy-MM-dd HH-mm-ss");
		Date dat = new Date();
		ATUTestRecorder recorder = new ATUTestRecorder("C:\\Users\\sarav\\OneDrive\\Documents\\ATU Record files",
				"Testvideo" + simple.format(dat), false);
		recorder.start();

		driver.manage().window().maximize();
		driver.get("http://www.adactin.com/HotelApp/index.php");
		driver.manage().timeouts().implicitlyWait(50, TimeUnit.SECONDS);

		File loc = new File("C:\\Users\\sarav\\Downloads\\DataSheet.xlsx");

		FileInputStream stream = new FileInputStream(loc);
		Workbook w = new XSSFWorkbook(stream);
		Sheet s = w.getSheet("Data");

		for (int i = 1; i < s.getPhysicalNumberOfRows(); i++) {
			Row r = s.getRow(i);
			for (int j = 0; j < r.getPhysicalNumberOfCells(); j++) {
				Cell c = r.getCell(j);

				DataFormatter formatter = new DataFormatter();
				String cellData = formatter.formatCellValue(c);
				System.out.println(cellData);
				if (j == 0) {
					driver.findElement(By.id("username")).sendKeys(cellData);
				} else {
					driver.findElement(By.id("password")).sendKeys(cellData);
				}
			}

			driver.findElement(By.id("login")).click();
			
			WebElement error = driver.findElement(By.xpath("//*[contains(text(),'Invalid Login details')]"));
			String text = error.getText();
			Cell c = r.createCell(2);
			c.setCellValue(text);
			FileOutputStream out = new FileOutputStream(loc);
			w.write(out);

		}
		driver.quit();
		recorder.stop();

	}

}