package Sdet_Excel;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.concurrent.TimeUnit;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;



public class Example {

	public static void main(String[] args) throws IOException {
	
		FileInputStream fi = new FileInputStream("TestData.xlsx");
		Workbook book = new XSSFWorkbook(fi);
		Sheet sheet = book.getSheet("AdminLoginValidData");
		
		int rowcount = sheet.getLastRowNum();
		int colcount=sheet.getRow(0).getLastCellNum();
		
		for(int i=0;i<rowcount;i++)
		{
			Row row = sheet.getRow(i);
			for(int j=0;j<colcount;j++)
			{
				String data = row.getCell(j).toString();//to raed values from string
				System.out.print("    "+data);
			}
			System.out.println();
		}
		
				
		book.close();
		fi.close();


	
	
	//------------------------------------------------
	System.setProperty("webdriver.chrome.driver", "chromedriver.exe");		
		
		WebDriver driver = new ChromeDriver();		
		driver.manage().deleteAllCookies();
		driver.manage().window().maximize();		
		driver.manage().timeouts().implicitlyWait(20, TimeUnit.SECONDS);
		
		driver.get("https://facebook.com");
	FileInputStream fi1 = new FileInputStream("TestData.xlsx");
	Workbook book1 = new XSSFWorkbook(fi1);
	Sheet sheet1 = book1.getSheet("AdminLoginValidData");
	
	int rowcount1 = sheet1.getLastRowNum();
     for(int i=1;i<=rowcount;i++)
	{
		
		String empfname = sheet1.getRow(i).getCell(1).getStringCellValue();
		double sal =sheet1.getRow(i).getCell(2).getNumericCellValue();
		boolean status1 = sheet1.getRow(i).getCell(3).getBooleanCellValue();
		String emplname = sheet1.getRow(i).getCell(4).getStringCellValue();
		String empdesig = sheet1.getRow(i).getCell(5).getStringCellValue();
		
		
		driver.findElement(By.linkText("Register")).click();
	     //entering information
	     driver.findElement(By.id("firstname")).sendKeys(empfname);
	     driver.findElement(By.id("lastname")).sendKeys(emplname);
	     driver.findElement(By.id("empdesigntion")).sendKeys(empdesig);
	
	}
 	book1.close();
 	fi1.close();
    
	// writing data into excel
 	FileOutputStream fo = new FileOutputStream("C://Users//Admin//OneDrive//Documents//Book32.xlsx");
	Workbook book2 = new XSSFWorkbook();
	Sheet sheet2 = book2.createSheet("happy");
	
     for(int i=0;i<=5;i++)
	{

    	 Row row = sheet2.createRow(i);
    	 for(int j=0;j<3;j++)
    	 {
    		 row.createCell(j).setCellValue("xyz");
    	 }
	}
     book2.write(fo);
     book2.close();
     fo.close();

}
}
