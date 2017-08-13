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
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

public class detailedreportmain {

	public static void main(String[] args) throws EncryptedDocumentException, InvalidFormatException, IOException {
	
		FileInputStream fis=new FileInputStream("C:/Selenium/Testdata/Testdata.xlsx");
		Workbook wb=WorkbookFactory.create(fis);
		Sheet s=wb.getSheet("Sheet4");
		s.createRow(0);
		
		
		
		System.setProperty("webdriver.chrome.driver","C:\\Selenium\\chromedriver_win32\\chromedriver.exe");
		WebDriver driver = new ChromeDriver();
		driver.manage().window().maximize();
		driver.get("https://products.office.com/en-US");
		//driver.findElement(By.id("need_close")).click();
		//driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
		//driver.switchTo().alert().dismiss();
		//driver.findElement(By.className("accordion-expand-all")).click();
	
		//driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
		WebElement main= driver.findElement(By.xpath("//div[@role='main']"));
        System.out.println(main.findElements(By.tagName("a")).size()) ; 
        List<WebElement> alllinks = main.findElements(By.tagName("a"));
		 int count = alllinks.size();
		 
		List<String> href = new ArrayList<String>();
		
		for(int i=1;i<count;i++)
        {
			if(alllinks.get(i).getText() == null | (alllinks.get(i).getText()).equalsIgnoreCase("")|alllinks.get(i).getAttribute("href") == null);
			
			else if((alllinks.get(i).getAttribute("href")==null))
		          href.add("No href");
		          else
		        	  href.add(alllinks.get(i).getAttribute("id"));
		
		
        }
		
		List<String> title = new ArrayList<String>();
		for(int i=1;i<count;i++)
        {
			if(alllinks.get(i).getText() == null | (alllinks.get(i).getText()).equalsIgnoreCase("")|alllinks.get(i).getAttribute("href") == null);
			
			else if((alllinks.get(i).getText()==null))
		          title.add("No title");
		          else
		        	  title.add(alllinks.get(i).getAttribute("id"));
                  
        }
        	
List<String> id = new ArrayList<String>();
for(int i=1;i<count;i++)
{
	if(alllinks.get(i).getText() == null | (alllinks.get(i).getText()).equalsIgnoreCase("")|alllinks.get(i).getAttribute("href") == null);
	
          else if((alllinks.get(i).getAttribute("id")==null))
          id.add("No Id");
          else
        	  id.add(alllinks.get(i).getAttribute("id"));
          
		
		
		}
		
List<String> mscom = new ArrayList<String>();

for(int i=1;i<count;i++)
{
	if(alllinks.get(i).getText() == null | (alllinks.get(i).getText()).equalsIgnoreCase("")|alllinks.get(i).getAttribute("href") == null);
	
	else if((alllinks.get(i).getAttribute("class")==null))
        id.add("No Class");
        else
      	  id.add(alllinks.get(i).getAttribute("class"));
}
List<String> linkcontrol = new ArrayList<String>();
for(int i=0;i<mscom.size();i++)
{
	
if(mscom.get(i).contains("mscom")){
linkcontrol.add("Yes");
}
else{
	linkcontrol.add("No");
}
	
}
		s.getRow(0).createCell(1).setCellValue("S No");

		s.getRow(0).createCell(1).setCellValue("Href");
		
		s.getRow(0).createCell(2).setCellValue("Title");
		
		s.getRow(0).createCell(4).setCellValue("Id");
		
		s.getRow(0).createCell(6).setCellValue("Class");
		
		s.getRow(0).createCell(8).setCellValue("Link Control");
		
		System.out.println("started typing href");
	
		for (int i = 1; i < href.size(); i++) {
			
			
			s.getRow(i).createCell(0).setCellValue(i);
		}
		for (int i = 1; i < href.size(); i++) {
			String ss = href.get(i);
		//	s.createRow(i).createCell(0).setCellValue(ss);
			s.createRow(1);
			s.getRow(i).createCell(1).setCellValue(ss);
		}
		
		
		for (int i = 1; i < title.size(); i++) {
			String pp = title.get(i);
			s.createRow(1);
		s.getRow(i).createCell(2).setCellValue(pp);
		//s.createRow(i).createCell(2).setCellValue(pp);
		}
		for (int i = 1; i < id.size(); i++) {
			String tt = id.get(i);
			//s.createRow(i).createCell(4).setCellValue(tt);
			s.createRow(i);
			s.getRow(i).createCell(4).setCellValue(tt);
		
	}
	
		for (int i = 1; i < mscom.size(); i++) {
			String qq = mscom.get(i);
			//s.createRow(i).createCell(4).setCellValue(tt);
			s.createRow(1);
			s.getRow(i).createCell(6).setCellValue(qq);
		
	}
		for (int i = 1; i < linkcontrol.size(); i++) {
			String bb = linkcontrol.get(i);
			//s.createRow(i).createCell(4).setCellValue(tt);
			s.createRow(1);
			s.getRow(i).createCell(8).setCellValue(bb);
		
	}
		
		FileOutputStream fos=new FileOutputStream("C:/Selenium/Testdata/Testdata.xlsx");
		wb.write(fos);
		
		System.out.println("prog ends");
	}

	}




