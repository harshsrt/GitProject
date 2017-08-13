package WriteExcel;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

public class WriteExcel1 {

	public static void main(String[] args) throws EncryptedDocumentException, InvalidFormatException, IOException {
	
		FileInputStream fis=new FileInputStream("C:/Selenium/Testdata/Testdata.xlsx");
		Workbook wb=WorkbookFactory.create(fis);
		Sheet s=wb.getSheet("sheet1");
		System.setProperty("webdriver.chrome.driver","C:\\Selenium\\chromedriver_win32\\chromedriver.exe");
		WebDriver driver = new ChromeDriver();
		driver.manage().window().maximize();
		driver.get("https://www.google.com");
		List<WebElement> alllinks= driver.findElements(By.tagName("a"));
		Row r=s.createRow(0);
		for(int i =0;i<alllinks.size();i++)
		{
			
			String linktext=alllinks.get(i).getAttribute("href");
			s.createRow(i).createCell(0).setCellValue(linktext);
		}
		FileOutputStream fos=new FileOutputStream("C:/Selenium/Testdata/Testdata.xlsx");
		wb.write(fos);
	}
	}


