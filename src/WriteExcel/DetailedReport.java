package WriteExcel;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.TimeUnit;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;



public class DetailedReport {
	
	
	public static void main(String[] args) throws EncryptedDocumentException, InvalidFormatException, IOException {
	
		FileInputStream fis=new FileInputStream("C:/Selenium/Testdata/Testdata.xlsx");
		Workbook wb=WorkbookFactory.create(fis);
		Sheet s=wb.getSheet("sheet4");
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
        List<WebElement> alllinks =main.findElements(By.tagName("a"));
        
		 int count = alllinks.size();
		System.out.println(count); 
		List<String> href = new ArrayList<String>();
		
		for(WebElement x:alllinks)
		{
			if (x.getText().length()!=0) 
			{
			href.add(x.getAttribute("href"));
		
			} 
		}
		
		
		List<String> title = new ArrayList<String>();
		
		for (WebElement y : alllinks) {
			if (y.getText().length()!=0)
			{
				title.add(y.getAttribute("ms.title"));
			} 
		}
List<String> id = new ArrayList<String>();
		
		for (WebElement z : alllinks) {
			if (z.getText().length()!=0) 
			{
				if(z.getAttribute("id").length()!=0)
				id.add(z.getAttribute("id"));
				else{
					id.add("No Id Found");
				}
				
					
			} 
		}
		
List<String> mscom = new ArrayList<String>();
		
		for (WebElement a : alllinks) {
			if (a.getText().length()!=0)
			{
			if(a.getAttribute("class").contains("mscom"))
			mscom.add("Yes");
			else{
			mscom.add("No");
			}
			} 
		}
		
List<String> text = new ArrayList<String>();
		
		for (WebElement b : alllinks)
		{
			if (b.getText().length()!=0) 
			{
				text.add(b.getText());
			} 
		}
List<String> external = new ArrayList<String>();
		
		for (WebElement c : alllinks)
		{
			if (c.getText().length()!=0) 
			{
			 
					if(c.getAttribute("target").contains("_self"))
				{
				external.add("Internal Link");
				}
					else if(c.getAttribute("target").contains("_blank"))
					{
						external.add("External Link");
						
					}else{
						external.add("Inpage Link");
					}
				
				
				
			
			}
		}
		List<String> https = new ArrayList<String>();
		
		for (WebElement d : alllinks) {
			if (d.getText().length()!=0)
			{
			if(d.getAttribute("href").contains("https://"))
			https.add("Yes");
			else{
			https.add("No");
			}
			} 
		}
		
		   
		s.getRow(0).createCell(0).setCellValue("Href");
		
		
		//s.getRow(0).createCell(2).setCellValue("Title");
		s.getRow(0).createCell(2).setCellValue("Link text");
		
		s.getRow(0).createCell(4).setCellValue("Id");
		
		
		s.getRow(0).createCell(5).setCellValue("Class Control(mscom) ");
		
		s.getRow(0).createCell(6).setCellValue("Internal/External ");
		s.getRow(0).createCell(7).setCellValue("Https on Hover");
		//s.getRow(0).createCell(8).setCellValue("Link text");
		
		
		
		System.out.println("started typing href");
		
		for (int i = 0; i < href.size(); i++) {
			String ss = href.get(i);
			System.out.println(ss);
		//	s.createRow(i).createCell(0).setCellValue(ss);
			s.createRow(i+2);
			s.getRow(i+2).createCell(0).setCellValue(ss);
		}
		
		for (int i = 0; i < text.size(); i++) {
			String qq = text.get(i);
			//s.createRow(i).createCell(4).setCellValue(tt);
			s.createRow(100);
			s.getRow(i+2).createCell(2).setCellValue(qq);
		
	}
		//for (int i = 0; i < title.size(); i++) {
		//	String pp = title.get(i);
		//	s.createRow(100);
		//s.getRow(i+1).createCell(2).setCellValue(pp);
		//s.createRow(i).createCell(2).setCellValue(pp);
		//}
	
		for (int i = 0; i < id.size(); i++) {
			String tt = id.get(i);
			//s.createRow(i).createCell(4).setCellValue(tt);
			s.createRow(100);
			s.getRow(i+2).createCell(4).setCellValue(tt);
		
	}
		for (int i = 0; i < mscom.size(); i++) {
			String qq = mscom.get(i);
			//s.createRow(i).createCell(4).setCellValue(tt);
			s.createRow(100);
			s.getRow(i+2).createCell(5).setCellValue(qq);
		
	}
		for (int i = 0; i < external.size(); i++) {
			String gg = external.get(i);
			//s.createRow(i).createCell(4).setCellValue(tt);
			//s.createRow(i+2);
			s.getRow(i+2).createCell(6).setCellValue(gg);
		
	}
		for (int i = 0; i < https.size(); i++) {
			String ff = https.get(i);
			//s.createRow(i).createCell(4).setCellValue(tt);
			//s.createRow(i+2);
			s.getRow(i+2).createCell(7).setCellValue(ff);
		
	}
		
		
		FileOutputStream fos=new FileOutputStream("C:/Selenium/Testdata/Testdata.xlsx");
		wb.write(fos);
		
		System.out.println("prog ends");
	}

	

	}


	

