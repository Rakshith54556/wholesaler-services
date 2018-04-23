package cubevalues;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Properties;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.ie.InternetExplorerDriver;

public class ReportValues {
	
	public static WebDriver driver;
	
	static String a;
	static File filename=new File("C:\\Users\\RBS\\Desktop\\workbook1.xlsx");
	
//	static String driverpath="C:\\Users\\RBS\\Desktop\\Selenium";

	public static void main(String[] args) throws IOException, InterruptedException {
		
		Properties prop=new Properties();
		
		FileInputStream obj1=new FileInputStream("C:\\Users\\RBS\\git\\wholesalerservices\\cubevalues\\config.properties");
		prop.load(obj1);
		System.setProperty("webdriver.ie.driver", "C:\\Users\\RBS\\Desktop\\Selenium\\IEDriverServer.exe");
		driver= new InternetExplorerDriver();
		
		driver.get(prop.getProperty("URL"));
		driver.manage().window().maximize();
		
		
	    driver.findElement(By.name("username")).sendKeys(prop.getProperty("username"));
	   driver.findElement(By.name("password")).sendKeys(prop.getProperty("password"));
	  driver.findElement(By.name("SubmitCreds")).click();
	  
	  driver.findElement(By.xpath(prop.getProperty("EXECXPATH"))).click();
	  Thread.sleep(10000);
	  
	 a= driver.findElement(By.xpath("*//div[contains(text(),'56,9')] ")).getText();
	 
	 System.out.println(a);
	 
	 FileInputStream fis= new FileInputStream(filename);
	 XSSFWorkbook workbook = new XSSFWorkbook(fis);
	 XSSFSheet worksheet= workbook.createSheet("report calculation");
	 
	 XSSFRow row1=null;
	 XSSFCell cell=null;
	 
	 row1=worksheet.createRow(0);
	 worksheet.getRow(0);
	 cell= row1.createCell(1);
	 cell.setCellValue("Report values");
	 
	 for (int i=0; i<=5 ;i++){
	 row1=worksheet.createRow(i+1);
	 worksheet.getRow(i+1);
	 cell= row1.createCell(1);
	 cell.setCellValue(a);
	 FileOutputStream fos =new FileOutputStream(filename);
	 workbook.write(fos);
	  
	  driver.close();
	  
	  
		
	 }
		

	}

}
