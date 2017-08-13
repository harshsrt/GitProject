package WriteExcel;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.HttpURLConnection;
import java.net.URL;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.TimeUnit;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

public class Brokenlink_main {

	static int i;
	public static void main(String[] args) throws IllegalStateException, InvalidFormatException, IOException {
		
		System.setProperty("webdriver.chrome.driver","C:\\Selenium\\chromedriver_win32\\chromedriver.exe");
   		WebDriver driver = new ChromeDriver();
   		
		driver.manage().window().maximize();
		driver.get("https://products.office.com/en-us/office-365-personal");
		driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
	
		JavascriptExecutor jse = (JavascriptExecutor)driver;
		jse.executeScript("scroll(0, 1000);");
		
		driver.findElement(By.id("feedbackSectionCloseBtn")).click();
		
		driver.findElement(By.xpath("//span[contains(text(), 'Show all')]")).click();
		
		WebElement main= driver.findElement(By.xpath("//div[@role='main']"));
	       // System.out.println(main.findElements(By.tagName("a")).size()) ; 
	        List<WebElement> links =main.findElements(By.tagName("a"));
		
		System.out.println("Total links are "+links.size());
		
		for(i=0;i<links.size();i++)
			
		{
			
			{
			WebElement ele= links.get(i);
			if(ele.getText().length()!=0)
			{
			String url=ele.getAttribute("href");
	
			verifyLinkActive(url);
			
			
			}
			}
		}
	
			
		
			
			
		
	}
	
	public static void verifyLinkActive(String linkUrl) throws EncryptedDocumentException, InvalidFormatException, IOException
	{
		FileInputStream fis=new FileInputStream("C:/Selenium/Testdata/Testdata.xlsx");
		Workbook wb=WorkbookFactory.create(fis);
		Sheet s=wb.getSheet("sheet3");
		s.createRow(0);
		s.getRow(0).createCell(0).setCellValue("Link Status");
		
		
        try 
        {
        	
           URL url = new URL(linkUrl);
           
           HttpURLConnection httpURLConnect=(HttpURLConnection)url.openConnection();
           
           httpURLConnect.setConnectTimeout(500);
           
           httpURLConnect.connect();
           
           if(httpURLConnect.getResponseCode()==200)
          {
        
   			String linktext=(linkUrl+" - "+httpURLConnect.getResponseMessage()+"  "+httpURLConnect.getResponseCode());
   			s.createRow(i);
   			s.getRow(i+2).createCell(0).setCellValue(linktext);
			 System.out.println(linkUrl+" - "+httpURLConnect.getResponseMessage()+"  "+httpURLConnect.getResponseCode());
   			
          }
	
           
          if(httpURLConnect.getResponseCode()==HttpURLConnection.HTTP_NOT_FOUND)  
           {
        	  
              String linktext=linkUrl+" - "+httpURLConnect.getResponseMessage() + " - "+ HttpURLConnection.HTTP_NOT_FOUND+"  "+httpURLConnect.getResponseCode();
              s.createRow(i);
              s.getRow(i+2).createCell(0).setCellValue(linktext);
             
              System.out.println(linkUrl+" - "+httpURLConnect.getResponseMessage()+"  "+httpURLConnect.getResponseCode());
            }
         
          FileOutputStream fos=new FileOutputStream("C:/Selenium/Testdata/Testdata.xlsx");
  		wb.write(fos);
          
        } catch (Exception e) {
        	
        }
    } 

	}



