package WriteExcel ;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

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


public class visiblelink {

       public static void main(String[] args) throws EncryptedDocumentException, InvalidFormatException, IOException {
    	   System.setProperty("webdriver.chrome.driver","C:\\Selenium\\chromedriver_win32\\chromedriver.exe");
   		WebDriver driver = new ChromeDriver();
   	   
              //Open the URL (Website)
        driver.get("https://products.office.com/en-in/office-365-personal");
        FileInputStream fis=new FileInputStream("C:/Selenium/Testdata/Testdata.xlsx");
        Workbook wb=WorkbookFactory.create(fis);
		//create a sheet
		Sheet s=wb.getSheet("sheet3");
        CellStyle cellStyle = s.getWorkbook().createCellStyle();
	    Font font = s.getWorkbook().createFont();
	    font.setBold(true);
	    font.setFontHeightInPoints((short) 16);
	    cellStyle.setFont(font);
	   
	   
	    Row row = s.createRow(0);
	    Cell cellTitle = row.createCell(0);
	    	
	    cellTitle.setCellStyle(cellStyle);
	    cellTitle.setCellValue("Detailed Report");
	   
	    
       
       
	   
        WebElement main= driver.findElement(By.xpath("//div[@role='main']"));
        System.out.println(main.findElements(By.tagName("a")).size()) ; 
        List<WebElement> mainlink = main.findElements(By.tagName("a"));
       
        // Count the total Link list on Web Page 
        int linkListCount = mainlink.size();
               
        //Print the total count of links on webpage
            
        System.out.println("Total Number of link on webpage = "  + linkListCount);
    
        for(int i=1;i<linkListCount;i++)
        {
        	if(mainlink.get(i).getText()== null | (mainlink.get(i).getText()).equalsIgnoreCase("")|mainlink.get(i).getAttribute("href") == null);                 
                  else
                  
                	  System.out.println(mainlink.get(i).getText()+" $ "+mainlink.get(i).getAttribute("id")+" $ "+mainlink.get(i).getAttribute("ms.title")+" $ "+mainlink.get(i).getAttribute("class"));
        	String ss=mainlink.get(i).getText()+" $ "+mainlink.get(i).getAttribute("id")+" $ "+mainlink.get(i).getAttribute("ms.title")+" $ "+mainlink.get(i).getAttribute("class");
        	s.createRow(i).createCell(0).setCellValue(ss);
                  
                  
        	FileOutputStream fos=new FileOutputStream("C:/Selenium/Testdata/Testdata.xlsx");
			wb.write(fos);
                  
        
        
                  } }}


       
        
        
        
       
        
       
       
        
     
       
       


