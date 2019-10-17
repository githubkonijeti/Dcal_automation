package com.dcal;

import java.io.FileInputStream;
import java.util.concurrent.TimeUnit;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import com.microsoft.schemas.office.visio.x2012.main.CellType;
import junit.framework.Test;
import java.io.IOException;
import org.openqa.selenium.Keys;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFCell;
//import org.apache.poi.sl.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.ss.usermodel.Row.MissingCellPolicy;
import org.apache.poi.ss.usermodel.Sheet;
import org.openqa.selenium.By;

//import org.openqa.selenium.firefox.FirefoxDriver;
public class DosingCal {
	//private static final String XSSFCell = null;
	//private static final CellAddress B1 = null;
	WebDriver driver;
	private XSSFWorkbook wb2;
	private XSSFWorkbook wb;

	public void invokeBrowser() {
		System.setProperty("webdriver.chrome.driver", "C:\\Users\\Ramu\\chromedriver_win32\\chromedriver.exe");
		driver = new ChromeDriver();
		driver.manage().deleteAllCookies();
		driver.manage().window().maximize();
		driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
		driver.manage().timeouts().pageLoadTimeout(10, TimeUnit.SECONDS);
		driver.get("https://dosingcalculator-qa.nxstage.com");
	}

	public void Login() throws IOException, InterruptedException {
		try {
		driver.findElement(By.xpath("//*[@name=\"ctl00$MainContent$Continue\"]")).click();
		FileInputStream fis = new FileInputStream("D:\\Test_cred.xlsx");
		wb = new XSSFWorkbook(fis);
		XSSFSheet sheet1 = wb.getSheetAt(0);
		int rowCount = sheet1.getLastRowNum();
		// System.out.println(rowCount);
		for (int row = 0; row <= 0; row++) {
			 
			  int i = row;
			  if(i == 0) {
				  String Username = sheet1.getRow(i).getCell(1).getStringCellValue();
				  driver.findElement(By.id("UserName")).sendKeys(Username);
				  System.out.println(Username);		
				  Thread.sleep(500);
			  }
			  
	
		}
		for (int row = 1; row <= rowCount; row++) {
			int i = row;
			if(i == 1) {
				String Password = sheet1.getRow(i).getCell(1).getStringCellValue();
				driver.findElement(By.id("Password")).sendKeys(Password);
				System.out.println(Password);
			}
	
			//driver.findElement(By.id("MainContent_LoginButton")).click();
			//Thread.sleep(500);
		}
		driver.findElement(By.id("MainContent_LoginButton")).click();
		}
		
		catch (Exception e) {
            e.printStackTrace();
        }
	}
		
	  public void Inputs() throws IOException, InterruptedException {
		  //System.out.println("test");
		  
	  FileInputStream fis2 = new FileInputStream("D:\\Test_cred.xlsx");
	  wb2 = new XSSFWorkbook(fis2); 
	  XSSFSheet sheet2 = wb2.getSheetAt(1);
	  int rowCount = sheet2.getLastRowNum();
	  
	  for (int row=0;row<=rowCount; row++) { 
		  int i = row;
		  
	  if(i == 0) {
		  
		  double Age = Double.parseDouble(sheet2.getRow(i).getCell(1).toString());
	      int Age1 = (int)Age;
	      String str = String.valueOf(Age1);  
		  //double Age = sheet2.getRow(i).getCell(1).getNumericCellValue();
		  driver.findElement(By.id("txtAge")).sendKeys(""+Age1);
	  }
	  if(i == 1) {
		  double Weight = Double.parseDouble(sheet2.getRow(i).getCell(1).toString());
		  int Weight1 = (int)Weight;
	      String str = String.valueOf(Weight1);
		  driver.findElement(By.id("txtWeight")).sendKeys(""+Weight1);
	  }
	  if(i == 2) {
		  double Height = Double.parseDouble(sheet2.getRow(i).getCell(1).toString());
		  int Height1 = (int)Height;
	      String str = String.valueOf(Height1);  
		  driver.findElement(By.id("txtHeight")).sendKeys(""+Height1);
	  }
	  if(i == 3) {
		  double Blood_flow_rate = sheet2.getRow(i).getCell(1).getNumericCellValue();
		  driver.findElement(By.id("txtBloodFlowRate")).sendKeys(""+Blood_flow_rate);
	  }
	  if(i == 4) {
		  double Residual_renal_function = Double.parseDouble(sheet2.getRow(i).getCell(1).toString());
		  int Residual_renal_function1 = (int)Residual_renal_function;
	      String str = String.valueOf(Residual_renal_function1);  
		  driver.findElement(By.id("txtResidualRenal")).sendKeys(""+Residual_renal_function1);
	  }
	  if(i == 5) {
		  double Treatment_maximum_UF_rate = sheet2.getRow(i).getCell(1).getNumericCellValue();
		  driver.findElement(By.id("txtMaxUFRate")).sendKeys(""+Treatment_maximum_UF_rate);
	  }	

	  if(i == 6) {
		  double Weekly_UF_volume = sheet2.getRow(i).getCell(1).getNumericCellValue();
		  driver.findElement(By.id("txtWeeklyUFvolume")).sendKeys(""+Weekly_UF_volume);
	  }	
	  
	  if(i == 7) {
		  double Minimum_hours_week = sheet2.getRow(i).getCell(1).getNumericCellValue();
		  driver.findElement(By.id("txtMinHoursWeek")).sendKeys(""+Minimum_hours_week);
	  }	

	 if(i == 8) {
		
		double Target_weekly_stdKt_V = sheet2.getRow(i).getCell(1).getNumericCellValue();
		driver.findElement(By.id("txtTargetWeeklystdKt")).sendKeys(""+Target_weekly_stdKt_V);
	 }
	 
	  }
	  driver.findElement(By.id("btnCalculate")).click();
	  }
	  
	  
	  public void InputDirectEntry() throws IOException, InterruptedException {
		  //System.out.println("test");
		  
	  FileInputStream fis2 = new FileInputStream("D:\\Test_cred.xlsx");
	  wb2 = new XSSFWorkbook(fis2); 
	  XSSFSheet sheet2 = wb2.getSheetAt(2);
	  int rowCount = sheet2.getLastRowNum();
	  driver.findElement(By.id("rbDirEnt")).click();
	  
	  for (int row=0;row<=rowCount; row++) { 
		  int i = row;
		  
	  if(i == 0) {
		  
		  double TBW = Double.parseDouble(sheet2.getRow(i).getCell(1).toString());
	      int TBW1 = (int)TBW;
	      String str = String.valueOf(TBW1);  
		  driver.findElement(By.id("txtTBW")).sendKeys(""+TBW1);
	  }
	  
	  if(i == 1) {
		  double Blood_flow_rate = sheet2.getRow(i).getCell(1).getNumericCellValue();
		  driver.findElement(By.id("txtBloodFlowRate")).sendKeys(""+Blood_flow_rate);
	  }
	  
	  if(i == 2) {
		  double Residual_renal_function = Double.parseDouble(sheet2.getRow(i).getCell(1).toString());
		  int Residual_renal_function1 = (int)Residual_renal_function;
	      String str = String.valueOf(Residual_renal_function1);  
		  driver.findElement(By.id("txtResidualRenal")).sendKeys(""+Residual_renal_function1);
	  }
	  
	  if(i == 3) {
		  double Cycler_maximum_UF_rate = sheet2.getRow(i).getCell(1).getNumericCellValue();
		  driver.findElement(By.id("txtCyclerSettingMaxUFRate")).sendKeys(""+Cycler_maximum_UF_rate);
	  }	
 
	  if(i == 4) {
		  double Weekly_UF_volume = sheet2.getRow(i).getCell(1).getNumericCellValue();
		  driver.findElement(By.id("txtWeeklyUFvolume")).sendKeys(""+Weekly_UF_volume);
	  }	
	  
	  if(i == 5) {
		  double Minimum_hours_week = sheet2.getRow(i).getCell(1).getNumericCellValue();
		  driver.findElement(By.id("txtMinHoursWeek")).sendKeys(""+Minimum_hours_week);
	  }	

	 if(i == 6) {
		
		double Target_weekly_stdKt_V = sheet2.getRow(i).getCell(1).getNumericCellValue();
		driver.findElement(By.id("txtTargetWeeklystdKt")).sendKeys(""+Target_weekly_stdKt_V);
	 }
	 
	  }
	  driver.findElement(By.id("btnCalculate")).click();
	  }
	  
	  public void Logout() throws InterruptedException { 
		   driver.findElement(By.id("HeadLoginView_HeadLoginStatus")).click(); 
		  Thread.sleep(2000);
		  driver.close(); 
		  } 
	  	  
	 public static void main(String[] args) throws IOException, InterruptedException {
		DosingCal obj = new DosingCal();
		obj.invokeBrowser();
		obj.Login();
		//obj.Inputs();
		//obj.InputDirectEntry();
		obj.Logout();
	}

}
