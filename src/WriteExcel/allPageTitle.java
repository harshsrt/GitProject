package WriteExcel;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
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

public class allPageTitle {
	 
	public static void main(String[] args) throws EncryptedDocumentException, InvalidFormatException, IOException {
		
		FileInputStream fis=new FileInputStream("C:/Selenium/Testdata/Testdata.xlsx");
		Workbook wb=WorkbookFactory.create(fis);
		Sheet s=wb.getSheet("sheet3");
		s.createRow(0);
		System.setProperty("webdriver.chrome.driver","C:\\Selenium\\chromedriver_win32\\chromedriver.exe");
		WebDriver driver = new ChromeDriver();
		driver.manage().window().maximize();
		driver.get("https://products.office.com/en-US");
		driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
	
	
		WebElement main= driver.findElement(By.xpath("//div[@role='main']"));
       // System.out.println(main.findElements(By.tagName("a")).size()) ; 
        List<WebElement> alllinks =main.findElements(By.tagName("a"));
        
		 int count = alllinks.size();
		System.out.println(count); 
		
		List<String> text = new ArrayList<String>();
		
		for (WebElement a : alllinks)
		{
			if (a.getText().length()!=0) 
			{
				
				text.add(a.getText());
			} 
		}
		List<String> href=new ArrayList<String>();
		
		for (WebElement c : alllinks)
		{
			
			if (c.getText().length()!=0) 
			{
				
			href.add(c.getAttribute("href"));
				
			}
		List<String> pagetitle = new ArrayList<String>();
		
		for(String hrefs : href)
		{
			driver.navigate().to(hrefs);
			System.out.println(driver.getTitle());
			pagetitle.add(driver.getTitle());
			  driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
		
		}
		
			s.getRow(0).createCell(0).setCellValue("Link text");
			 s.getRow(0).createCell(2).setCellValue("Page Title");
			 
				for (int i = 0; i < text.size(); i++) {
					String ss = text.get(i);
					System.out.println(ss);
				//	s.createRow(i).createCell(0).setCellValue(ss);
					s.createRow(i+2);
					s.getRow(i+2).createCell(0).setCellValue(ss);
				}
				
				for (int i = 0; i < pagetitle.size(); i++) {
					String qq = pagetitle.get(i);
					//s.createRow(i).createCell(4).setCellValue(tt);
					s.createRow(100);
					s.getRow(i+2).createCell(2).setCellValue(qq);
				}
				FileOutputStream fos=new FileOutputStream("C:/Selenium/Testdata/Testdata.xlsx");
				wb.write(fos);
				
				System.out.println("prog ends");
			}

			}
	}
	