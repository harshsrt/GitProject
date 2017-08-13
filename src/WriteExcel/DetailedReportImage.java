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

public class DetailedReportImage {

	
	public static void main(String[] args) throws EncryptedDocumentException, InvalidFormatException, IOException {
	
		FileInputStream fis=new FileInputStream("C:/Selenium/Testdata/Testdata.xlsx");
		Workbook wb=WorkbookFactory.create(fis);
		Sheet s=wb.getSheet("sheet5");
		s.createRow(0);
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
        List<WebElement> alllinks =main.findElements(By.tagName("img"));
        
		 int count = alllinks.size();
		System.out.println(count); 
		List<String> href = new ArrayList<String>();
		
		for(WebElement x:alllinks)
		{
			
			href.add(x.getAttribute("src"));
		
			
		}
		
		
		List<String> role = new ArrayList<String>();
		
		for (WebElement y : alllinks) 
		{

					role.add(y.getAttribute("role"));
			
			
				
			 
		}
List<String> alt = new ArrayList<String>();
		
		for (WebElement b : alllinks)
		{
			if(b.getAttribute("alt").length()!=0)
			{
				alt.add(b.getAttribute("alt"));
			}
			else if(b.getAttribute("alt").length()==0)
				{
					alt.add("Empty alt");
				}
					
						
		}
					
				
	
List<String> id = new ArrayList<String>();
		
		for (WebElement z : alllinks)
		{
			
				if(z.getAttribute("id").length()!=0)
				id.add(z.getAttribute("id"));
				else
				{
					id.add("No ID Found");
			
					
				} 
		}
		
List<String> mscom = new ArrayList<String>();
		
		for (WebElement a : alllinks) {
			//if (a.getText().length()!=0)
			//{
			if(a.getAttribute("class").contains("mscom"))
			mscom.add("Yes");
			else{
			mscom.add("No");
			//}
			} 
		}
		
		s.createRow(0);
		s.getRow(0).createCell(0).setCellValue("Href");
		
		
		s.getRow(0).createCell(2).setCellValue("Role");
		s.getRow(0).createCell(4).setCellValue("Alt");
		
		s.getRow(0).createCell(6).setCellValue("ID");
		
		
		s.getRow(0).createCell(8).setCellValue("Class Control(mscom) ");
		
	//	s.getRow(0).createCell(6).setCellValue("Internal/External ");
		//s.getRow(0).createCell(7).setCellValue("Https on Hover");
		//s.getRow(0).createCell(8).setCellValue("Link text");
		
		
		
		System.out.println("started typing href");
		
		for (int i = 0; i < href.size(); i++) {
			String ss = href.get(i);
			System.out.println(ss);
		//	s.createRow(i).createCell(0).setCellValue(ss);
			s.createRow(i+2);
			s.getRow(i+2).createCell(0).setCellValue(ss);
		}
		for (int i = 0; i < role.size(); i++) {
			String pp = role.get(i);
			s.createRow(100);
		s.getRow(i+2).createCell(2).setCellValue(pp);
	//	s.createRow(i).createCell(2).setCellValue(pp);
		}
		for (int i = 0; i < alt.size(); i++) {
			String qq = alt.get(i);
		//	s.createRow(i).createCell(4).setCellValue(tt);
			s.createRow(100);
			s.getRow(i+2).createCell(4).setCellValue(qq);
		
	}
		
	
		for (int i = 0; i < id.size(); i++) {
			String tt = id.get(i);
			//s.createRow(i).createCell(4).setCellValue(tt);
			s.createRow(100);
			s.getRow(i+2).createCell(6).setCellValue(tt);
		
	}
		for (int i = 0; i < mscom.size(); i++) {
			String qq = mscom.get(i);
			//s.createRow(i).createCell(4).setCellValue(tt);
			s.createRow(100);
			s.getRow(i+2).createCell(8).setCellValue(qq);
		
	}
	
		
		
		FileOutputStream fos=new FileOutputStream("C:/Selenium/Testdata/Testdata.xlsx");
		wb.write(fos);
		
		System.out.println("prog ends");
	}

	}


	






