package readExcel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;
import java.util.Set;
import java.util.concurrent.TimeUnit;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.openqa.selenium.Alert;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Actions;
public class Naukri_resume {

	public static void main(String[] args) throws IllegalStateException, InvalidFormatException, IOException {
		 System.setProperty("webdriver.chrome.driver","C:\\Selenium\\chromedriver_win32\\chromedriver.exe");
	   		WebDriver driver = new ChromeDriver();
	   		File src= new File("C:\\Selenium\\Testdata\\TestData.xlsx");
			FileInputStream fis=new FileInputStream(src);
			Workbook wb=WorkbookFactory.create(fis);
			Sheet s=wb.getSheet("Sheet1");
			String p =s.getRow(0).getCell(0).getStringCellValue();
			String q =s.getRow(0).getCell(1).getStringCellValue();
	   		
		       
		       // driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
		        driver.get("http://www.naukri.com/");
		       // driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
		      
		        String homepage=driver.getWindowHandle();
		        System.out.println(homepage);
		        Set<String> windows=driver.getWindowHandles();
		        System.out.println(windows.size());
		        
		        Iterator iterator=windows.iterator();
		        String currentWindowId;
		        
		        while(iterator.hasNext()){
		        	currentWindowId=iterator.next().toString();
		        	System.out.println(currentWindowId);
		        	
		        	if(!currentWindowId.equals(homepage)){
		        		driver.switchTo().window(currentWindowId);
		        		driver.close();
		        		driver.switchTo().window(homepage);
		        	}	
		        	}
		        driver.manage().window().maximize();	
		        
		        driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);

		        driver.findElement(By.xpath(".//*[@id='login_Layer']/span")).click();
		        driver.findElement(By.xpath(".//*[@id='eLogin']")).sendKeys(p);
		        driver.findElement(By.xpath(".//*[@id='pLogin']")).sendKeys(q);
		        
		        driver.findElement(By.xpath(".//*[@id='lgnFrm']/div[8]/button")).click();
		        driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
		        driver.findElement(By.xpath(".//*[@id='leftNav_updateProfile']/ul/li/a[4]")).click();
		        driver.findElement(By.xpath(".//*[@id='uploadLink']")).click();
		        WebElement upload=driver.findElement(By.xpath(".//*[@id='attachCV']"));
		       // upload.click();
		        upload.sendKeys("C:\\Resume\\Harsh_Resume.doc");
		        driver.findElement(By.xpath(".//*[@id='editForm']/div[7]/button")).click();
		       
		        WebElement element = driver.findElement(By.xpath(".//*[@id='mainHeader']/div/div/ul[2]/li[2]/a/div[1]/span"));
		        
		        Actions action = new Actions(driver);
		 
		        action.moveToElement(element).build().perform();
		 
		        driver.findElement(By.linkText("Log Out")).click();
		 
		        
		      //  driver.findElement(By.xpath("//a[text()='Post Resume']")).click(); //click on post resume
		        
		      //  driver.switchTo().frame(driver.findElement(By.id("frmUpload"))); //browse button is under frame so 1st switch control to frame
		      //  driver.findElement(By.id("browsecv")).sendKeys("path of the file which u want to upload");
		       }

	

}

