package Sdet_Excel;

import java.io.IOException;
import java.util.concurrent.TimeUnit;

import org.junit.AfterClass;
import org.junit.BeforeClass;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.Assert;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

public class DataDrivenTesting_DataProvider {

	public  WebDriver driver;
	@BeforeTest
	public void setUp()
	{
		    System.setProperty("webdriver.chrome.driver", "chromedriver.exe");		
		    driver = new ChromeDriver();		
			driver.manage().window().maximize();		
			driver.manage().timeouts().implicitlyWait(20, TimeUnit.SECONDS);
			
		    
	}
	@Test(dataProvider = "LoginCredentials")
	public void loginTest(String uname, String pwd,String result) throws InterruptedException
	{
		driver.get("http://orangehrm.qedgetech.com");
		driver.findElement(By.id("txtUsername")).sendKeys(uname);
		driver.findElement(By.id("txtPassword")).sendKeys(pwd);
		driver.findElement(By.id("btnLogin")).click();
		String exp_url="http://orangehrm.qedgetech.com/symfony/web/index.php/dashboard";
		String act_url=driver.getCurrentUrl();
		if(result.equals("valid"))
		{
			if(exp_url.equals(act_url))
			{
			driver.findElement(By.partialLinkText("Welcome")).click();
			driver.findElement(By.linkText("Logout")).click();
			Thread.sleep(5000);
			Assert.assertTrue(true);
			}else
			{
				Assert.assertTrue(false);
			}
			
		}else if(result.equals("invalid"))
		{
			if(exp_url.equals(act_url))
			{
			driver.findElement(By.partialLinkText("Welcome")).click();
			driver.findElement(By.linkText("Logout")).click();
			Assert.assertTrue(false);
			}else
			{
				Assert.assertTrue(true);
			}
			
		}
		
		
		
	}
	@DataProvider(name="LoginCredentials")
	public Object[][] getLoginData() throws IOException
	{
		//we are sending hardcoded Object array values to loginTest without sending values from excel sheet
		/*
		Object[][] loginData= {{"Admin", "Qedge123!@#","valid"},{"abcvjss","Qedge123!@#","invalid"},{"Admin","dgsdgsg","invalid"},{"xcvxvvz","sdsdgsgsg","invalid"}};
		return loginData;
		*/
		
		//sending values from excelsheet to loginTest using Xlutils class
		String xlfile="Book1.xlsx";
		String xlsheet="Data";
		int rows = Xlutils.getRowCount(xlfile, xlsheet);
		int cells=Xlutils.getCellCount(xlfile, xlsheet, 1);
		Object[][] loginData=new Object[rows][cells];
		for(int i=1;i<=rows;i++)
		{
			for(int j=0;j<cells;j++)
			{
				loginData[i-1][j]=Xlutils.getCellData(xlfile, xlsheet, i, j);
			}
		}
		return loginData;
	}
	@AfterTest
	public void tearDown()
	{
		
		driver.quit();
	}
	
}
