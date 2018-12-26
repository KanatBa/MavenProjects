package com.cyber.zohopublic;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;
import java.util.concurrent.TimeUnit;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

import io.github.bonigarcia.wdm.WebDriverManager;

public class googleSearchTest {
		WebDriver driver;
	  	  
	@BeforeTest
	  public void beforeTest() throws IOException {
		  WebDriverManager.chromedriver().setup();
			driver=new ChromeDriver();
			driver.get("https://www.google.com/");
			driver.manage().window().maximize();
			driver.manage().timeouts().implicitlyWait(10,TimeUnit.SECONDS);
			
			  
	  }
	@Test
	public void searchResults() throws IOException {
		  File src=new File("C:\\Users\\kbaialiev\\Desktop\\testing-maven\\excelsearch\\TestData\\googleSearch.xlsx");
		  FileInputStream fls=new FileInputStream(src);
		  XSSFWorkbook wb=new XSSFWorkbook(fls);
		  XSSFSheet sheet1= wb.getSheetAt(0);
		  System.out.println(sheet1.getRow(1).getCell(0).getStringCellValue());
		  int colNum=sheet1.getRow(1).getPhysicalNumberOfCells();
		  System.out.println(colNum);
		  for(int i=1; i<=sheet1.getLastRowNum(); i++) {
			  
		  	driver.findElement(By.name("q")).sendKeys(sheet1.getRow(i).getCell(0).getStringCellValue());
			driver.findElement(By.name("q")).sendKeys(Keys.ENTER);
			String result=driver.findElement(By.id("resultStats")).getText();
			result=result.replaceAll("About ", "");
			result=result.replaceAll("results","");
			if(sheet1.getRow(i).getCell(colNum-1).getStringCellValue()!="") {
				sheet1.getRow(i).createCell(colNum).setCellValue(result);
				Date date = new Date();  
				sheet1.getRow(0).createCell(colNum).setCellValue("SearchResult "+date);
			}
				
			
			driver.findElement(By.name("q")).clear();
			
			
		  }
		  
		  FileOutputStream fos=new FileOutputStream("C:\\Users\\kbaialiev\\Desktop\\testing-maven\\excelsearch\\TestData\\googleSearch.xlsx");
		  wb.write(fos);
		  fos.close();
		  
		
	}
	
	@AfterTest
	  public void afterTest() {
		    driver.close();
			driver.quit();
	  }
	
	 
		    
		 
	
}
