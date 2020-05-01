package org.tcs.testing;

import java.awt.AWTException;
import java.awt.Robot;
import java.awt.event.KeyEvent;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Date;
import java.text.SimpleDateFormat;

import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Actions;

public class BaseClass {

	public static WebDriver driver;
	public static Robot r;

	public static void launchBrowser() {
		System.setProperty("webdriver.chrome.driver", "D:\\Selenium Kannan\\Baseclass\\Driver\\chromedriver.exe");
		driver = new ChromeDriver();
	}

	public static void pageUrl() {
		// TODO Auto-generated method stub
		String url = driver.getCurrentUrl();
		System.out.println(url);
	}

	public static void windowmax() {
		driver.manage().window().maximize();
	}

	public static void pageTitle() {
		String title = driver.getTitle();
		System.out.println(title);

	}

	public static void loadUrl(String url) {
		driver.get(url);
	}

	public static void fill(WebElement e, String s) {
		e.sendKeys(s);

	}

	public static void btnClick(WebElement e) {
		e.click();

	}

	public static void getTextEntered(WebElement e) {
		String attribute = e.getAttribute("value");
		System.out.println(attribute);
	}

	public static void getTheText(WebElement e) {
		String text = e.getText();
		System.out.println(text);

	}

	public static void windowClose() {
		driver.close();
	}

	public static void rightClick(WebElement e) {
		Actions a = new Actions(driver);
		a.contextClick(e).perform();
	}

	public static void doubleClick(WebElement e) {
		Actions a = new Actions(driver);
		a.doubleClick(e).perform();

	}

	public static void mouseHover(WebElement e) {
		Actions a = new Actions(driver);
		a.moveToElement(e).perform();

	}

	public static void draganddrop(WebElement s, WebElement d) {
		Actions a = new Actions(driver);
		a.dragAndDrop(s, d).perform();
	}

	public static void copy() throws AWTException {
		r = new Robot();
		r.keyPress(KeyEvent.VK_CONTROL);
		r.keyPress(KeyEvent.VK_C);
		r.keyRelease(KeyEvent.VK_CONTROL);
		r.keyRelease(KeyEvent.VK_C);

	}

	public static void paste() throws AWTException {
		r = new Robot();
		r.keyPress(KeyEvent.VK_CONTROL);
		r.keyPress(KeyEvent.VK_V);
		r.keyRelease(KeyEvent.VK_CONTROL);
		r.keyRelease(KeyEvent.VK_V);

	}

	public static void excelRead(String path, String sheetname, int row, int cell) throws IOException {
		File loc = new File(path);
		FileInputStream stream = new FileInputStream(loc);
		Workbook w = new XSSFWorkbook(stream);
		Sheet s = w.getSheet(sheetname);
		Row r = s.getRow(row);
		Cell c = r.getCell(cell);
		System.out.println(c);
	}

	public static void excelUpdateStringValue(String path, String sheetName, int row, int cell, String exName,
			String newName) throws IOException {
		File loc = new File(path);
		FileInputStream stream = new FileInputStream(loc);
		Workbook w = new XSSFWorkbook(stream);
		Sheet s = w.getSheet(sheetName);
		Row r = s.getRow(row);
		Cell c = r.getCell(cell);
		int cellType = c.getCellType();
		if (cellType == 1) {
			String s1 = c.getStringCellValue();
			if (s1.equals(exName)) {
				c.setCellValue(newName);
			}

		}

		FileOutputStream f = new FileOutputStream(loc);
		w.write(f);
		System.out.println(c);
		System.out.println("Update done. Please check your excelSheet");
	}
	
	public static void excelUpdateDateValue(String path, String sheetName, int row, int cell,String oldDate, String newDate) throws IOException {
		File loc = new File(path);
		FileInputStream stream = new FileInputStream(loc);
		Workbook w = new XSSFWorkbook(stream);
		Sheet s = w.getSheet(sheetName);
		Row r = s.getRow(row);
		Cell c = r.getCell(cell);
		int cellType = c.getCellType();
		if(cellType==0) {
			java.util.Date dateCellValue = c.getDateCellValue();
			SimpleDateFormat s1 = new SimpleDateFormat("dd-MM-yyyy");
			String format = s1.format(dateCellValue);
			if (format.equals(oldDate)) {
				c.setCellValue(newDate);
			}
			FileOutputStream f = new FileOutputStream(loc);
			w.write(f);
			System.out.println(c);
			System.out.println("Date Update Done. Please check your sheet");
			
						
		}
		
		
	}
	public static void screenshot(String path) throws IOException {
		TakesScreenshot tk=(TakesScreenshot)driver;
		File s = tk.getScreenshotAs(OutputType.FILE);
		File d = new File(path);
		FileUtils.copyFile(s, d);
		
	}

}
