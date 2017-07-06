/**
 * 
 */
package com.ExtractDocument;

import java.awt.AWTException;
import java.awt.Robot;
import java.awt.event.KeyEvent;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.concurrent.TimeUnit;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

/**
 * @author Admin
 *
 */
public class ExtractDocument 
{
	WebDriver driver;
	WebDriverWait wait;
	XSSFWorkbook workbook;
	XSSFSheet sheet;
	XSSFCell cell;
	
	@BeforeTest(alwaysRun = true)
	public void TestSetup()
	{
		driver = new FirefoxDriver();
		wait = new WebDriverWait(driver,100);
		driver.manage().window().maximize();
		driver.manage().timeouts().implicitlyWait(30, TimeUnit.MINUTES);
		//driver.get("https://www.google.com");
	}
	@Test(priority = 0)
	public void TestRun() throws IOException, AWTException
	{
		File src = new File("C:\\Users\\Admin\\workspace\\Practise\\Excel Sheet.xlsx");
		FileInputStream fin = new FileInputStream(src);
		workbook = new XSSFWorkbook(fin);
		sheet = workbook.getSheetAt(0);
		for(int i = 1; i<=sheet.getLastRowNum(); i++)
		{
			// Element one.
			cell = sheet.getRow(i).getCell(1);
			//cell.setCellType(Cell.CELL_TYPE_STRING);
			// Adjust date format.
			DateFormat dateformat = new SimpleDateFormat("dd-MM-yyyy");
			Date date = cell.getDateCellValue();
			String date1 = dateformat.format(date);
			System.out.println(date1);
			//driver.findElement(By.id("ID number one")).sendKeys(cell.getStringCellValue());
			
			// Element one.
			/*cell = sheet.getRow(i).getCell(1);
			cell.setCellType(Cell.CELL_TYPE_STRING);
			//driver.findElement(By.id("ID number one")).sendKeys(cell.getStringCellValue());
						
			// Element one.
			cell = sheet.getRow(i).getCell(1);
			cell.setCellType(Cell.CELL_TYPE_STRING);
			//driver.findElement(By.id("ID number one")).sendKeys(cell.getStringCellValue());
						
			// Element one.
			cell = sheet.getRow(i).getCell(1);
			cell.setCellType(Cell.CELL_TYPE_STRING);
			//driver.findElement(By.id("ID number one")).sendKeys(cell.getStringCellValue());
						
			// Element one.
			cell = sheet.getRow(i).getCell(1);
			cell.setCellType(Cell.CELL_TYPE_STRING);
			//driver.findElement(By.id("ID number one")).sendKeys(cell.getStringCellValue());
						
			// Element one.
			cell = sheet.getRow(i).getCell(1);
			cell.setCellType(Cell.CELL_TYPE_STRING);
			//driver.findElement(By.id("ID number one")).sendKeys(cell.getStringCellValue());
						
			// Element one.
			cell = sheet.getRow(i).getCell(1);
			cell.setCellType(Cell.CELL_TYPE_STRING);
			//driver.findElement(By.id("ID number one")).sendKeys(cell.getStringCellValue());
				*/	
			
			// Code to download File.
			Robot robo  = new Robot();
			robo.keyPress(KeyEvent.VK_DOWN);
		}
	}

}
