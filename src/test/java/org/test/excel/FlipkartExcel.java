package org.test.excel;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Set;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

public class FlipkartExcel {
	public static void main(String[] args) throws Throwable {
		int i=0;
		System.setProperty("webdriver.chrome.driver", "C:\\Users\\Rajesh\\eclipse-workspace\\Mavenn\\drivers\\chromedriver.exe");
		WebDriver driver=new ChromeDriver();
		driver.manage().window().maximize();
		driver.get("https://www.flipkart.com/");
		Thread.sleep(2000);
		WebElement close = driver.findElement(By.xpath("(//button[contains(@class,'AkmmA' )])[1]"));
		close.click();
		WebElement  srchTxt=driver.findElement(By.name("q"));
	    srchTxt.sendKeys("redmi mobiles");	
	    WebElement search = driver.findElement(By.xpath("//button[@type='submit']"));
		search.click();
		Thread.sleep(5000);
		List<WebElement> results = driver.findElements(By.xpath("//div[contains(@class,'wU')]"));

		File f = new File("C:\\Users\\Rajesh\\eclipse-workspace\\Mavenn\\target\\new.Xlsx");
		Workbook w = new XSSFWorkbook();
		Sheet s = w.createSheet("Write");
		for (WebElement x : results) {
			String text = x.getText();

			Row r = s.createRow(i);
			Cell cell = r.createCell(0);
			cell.setCellValue(text);
			i++;

		}

		FileOutputStream f1 = new FileOutputStream(f);
		w.write(f1);
		System.out.println("Done");

		driver.findElement(By.xpath("//div[contains(@class,'wU')]")).click();
		String Parentid = driver.getWindowHandle();
		Set<String> allWinid = driver.getWindowHandles();
		List<String> emp = new ArrayList<String>();
		emp.addAll(allWinid);
		String st = emp.get(1);
		driver.switchTo().window(st);
		Thread.sleep(4000);
		String resulttext = driver.findElement(By.xpath("//span[contains(@class,'KyD')]")).getText();
		System.out.println("String retrieved from wep page" + resulttext);

		File r = new File("C:\\Users\\Rajesh\\eclipse-workspace\\Mavenn\\target\\new.Xlsx");
		FileInputStream f3 = new FileInputStream(r);
		Workbook fli = new XSSFWorkbook(f3);
		Sheet sh = w.getSheet("Flip");

		Row ro = s.getRow(3);

		Cell c = ro.getCell(0);
		int e = c.getCellType();
		if (e == 1) {
			String q = c.getStringCellValue();
			System.out.println("String retrened from Excel" + q);
			if (q.contains("Sapphire")) {
				System.out.println("Process passed" + q);
			} else {
				System.out.println("Process Failed");
			}
		}

	}

	}


