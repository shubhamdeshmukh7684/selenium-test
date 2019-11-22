package google;

import org.testng.annotations.Test;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.DataProvider;

import java.io.File;
import java.io.IOException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.Assert;
import org.testng.annotations.AfterTest;

public class NewTest {
	  private WebDriver driver;
	  File inFile;
	  XSSFWorkbook wb;
	  XSSFSheet sheet;
  @Test (priority=1)	
  public void openGoogle() {
	  //driver.navigate().to("www.javatpoint.com");  
	  driver.get("http://google.com"); 
	  //Assert.assertEquals(""+driver.getTitle(), "Google");
	  System.out.println("Page Loaded : "+driver.getTitle());
	  driver.findElement(By.linkText("Gmail")).click();
  }
  @Test (dataProvider="testdata")	
  public void openGmail(String uname,String password1) {
	 
	  System.out.println(uname+"      "+password1);
	  Assert.assertEquals(""+driver.getTitle(), "Gmail");
	  System.out.print("\n"+driver.getTitle());
	  

  }
  @BeforeTest
  public void beforeTest() {    
	  System.setProperty("webdriver.chrome.driver", "C:\\Users\\Sush\\Downloads\\chromedriver_win32\\chromedriver.exe");
	  driver = new ChromeDriver(); 
	  
  }

  @AfterTest
  public void afterTest() {
	  driver.quit();
  }
  
  @DataProvider(name="testdata")
  public Object[][] TestDataFeed() throws IOException, InvalidFormatException
  {
	  Object data[][]=new Object[2][2];
	  inFile= new File("C:\\Users\\Sush\\Desktop\\Test.xlsx");
	  wb= new XSSFWorkbook(inFile);
	  sheet=wb.getSheetAt(0);
	  
	  data[0][0]=sheet.getRow(0).getCell(0).getStringCellValue();
	  data[0][1]=sheet.getRow(0).getCell(1).getStringCellValue();
	  data[1][0]=sheet.getRow(1).getCell(0).getStringCellValue();
	  data[1][1]=sheet.getRow(2).getCell(0).getStringCellValue();
	  
	  System.out.println(data);
	  return data;
	  
  }

}
