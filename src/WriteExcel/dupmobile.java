package WriteExcel;

import java.util.List;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.openqa.selenium.By;
import org.openqa.selenium.Dimension;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.Set;
import java.util.concurrent.TimeUnit;


public class dupmobile {

       public static void main(String[] args) throws EncryptedDocumentException, InvalidFormatException, IOException 
       {
    	   System.setProperty("webdriver.chrome.driver","C:\\Selenium\\chromedriver_win32\\chromedriver.exe");
   		WebDriver driver = new ChromeDriver();
   		Dimension d = new Dimension(420,600);
		//Resize the current window to the given dimension
		driver.manage().window().setSize(d);
              //Open the URL (Website)
        driver.get("https://products.office.com/en-us/compare-all-microsoft-office-products?tab=1");
       
        
		driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
        FileInputStream fis=new FileInputStream("C:/Selenium/Testdata/Testdata.xlsx");
        //create a workbook
		Workbook wb=WorkbookFactory.create(fis);
		//create a sheet
		Sheet s=wb.getSheet("sheet1");
		List<String> alldupId = new ArrayList<String>();

		// Identify the all the web element having id attribute
		// List
		List<WebElement> allLinkElements = driver.findElements(By.xpath("//*[@id]"));

		// Count the total Link list on Web Page
		int linkListCount = allLinkElements.size();
		System.out.println("No of web element with ID  " + linkListCount);
		//create a list to store all the ids of the web element
		List<String> allId = new ArrayList<String>();

		for (WebElement x : allLinkElements) {
			//System.out.println(x.getAttribute("id"));
			allId.add(x.getAttribute("id"));

		}
		//create a set in order to find the duplicate ids
		Set<String> store = new HashSet<>();
	
		for (String ids : allId) {
			//Set will allow only unique ids
			if (store.add(ids) == false) {
				System.out.println("Found an element having duplicate id as  :-  " + ids);
			//store all the duplicate ids to list	
			alldupId.add(ids);
			}
			//print the list to get the duplicat ids
				for(int i=0;i<alldupId.size();i++)
				{
					String dupIds=alldupId.get(i);
					s.createRow(i+1).createCell(0).setCellValue(dupIds);	
				}
				}
		//if size of the list is null , No duplicate ids found
		if(alldupId.size()==0)
		{
			 System.out.println("Hurray !! No duplicate id found");

		}else
		//if size of list is not null, duplicate ids found
		{
			System.out.println("Duplicate id found ! Please report a bug");
		}
		   			
				
				FileOutputStream fos=new FileOutputStream("C:/Selenium/Testdata/Testdata.xlsx");
				wb.write(fos);
			
			}
				
			}

		
	


     
       
        
        
        
       
        
       
       
        
     
       
       



