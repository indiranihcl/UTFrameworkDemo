package main.community;

import java.awt.AWTException;
import java.awt.HeadlessException;
import java.awt.Rectangle;
import java.awt.Robot;
import java.awt.Toolkit;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.List;

import javax.imageio.ImageIO;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

import com.relevantcodes.extentreports.ExtentReports;
import com.relevantcodes.extentreports.ExtentTest;
import com.relevantcodes.extentreports.LogStatus;

public class componentsDevelopment  {
	 static ExtentReports report;
	 static ExtentTest logger; 
	 static WebDriver driver;
	 static ExtentReports extent;
	 public static File source;
	 public static HashMap<String,String> orMap= new HashMap<String,String>();
	 public static HashMap<String,String> orMapTestdata=new HashMap<String,String>();
     public static ThreadLocal<ExtentTest> testwss = new ThreadLocal<ExtentTest>();
     public static String idxpath;
     public static String namexpath;
     public static String classxpath;
     public static String classxpathval;
     public static String idxpathval;
     public static String namexpathval;
     public static String xpathval;
     public static String browserURl;
     public String currClass;
     public static int pageKey;
     
     public static String browser_type;
     public static String UrlValue;
     
	@BeforeTest
	public void startReport(){
	 extent = new ExtentReports (System.getProperty("user.dir") +"/test-output/AutomationReport.html", true);
	 extent.addSystemInfo("Host Name", "Component Framework").addSystemInfo("Environment", "Automation Testing").addSystemInfo("User Name", "Indirani");
	 extent.loadConfig(new File(System.getProperty("user.dir")+"\\extent-config.xml"));
	 }

	@Test(priority=1)
	 public void FindControls() throws Exception
	 {
	 //logger=extent.startTest("Finding Components ");
	 InitiateURL();
	 browserURl = driver.getCurrentUrl();
	 pageKey=1;
	 List<WebElement> el = driver.findElements(By.cssSelector("*"));
	  for ( WebElement e : el ) {
	    if(e.getTagName().equals("input")) {
	    //System.out.println("%%%%%%%%%%%" + e.getAttribute("type"));
	    	String controlName;
	    switch(e.getAttribute("type"))
	    {
	    case "text":
	    	idxpath=e.getAttribute("id");
	    	namexpath=e.getAttribute("name");
	    	classxpath=e.getAttribute("class");
	    	if((!(namexpath==null)) || (!(classxpath==null)) || (!(idxpath==null)))
	    	{
	    	idxpathval="//input[@id='"+idxpath+"']";
	    	namexpathval="//input[@name='"+namexpath+"']";
	    	classxpathval="//input[@class='"+classxpath+"']";
	    	xpathval=idxpathval;
	    	}
	    	controlName="TextBox";
	    	writeResult(pageKey,browserURl, controlName,idxpathval,namexpathval,classxpathval,xpathval);
	    	break;
	    	
	    case "checkbox":
	    	idxpath=e.getAttribute("id");
	    	namexpath=e.getAttribute("name");
	    	classxpath=e.getAttribute("class");
	    	if((!(namexpath==null)) || (!(classxpath==null)) || (!(idxpath==null)))
	    	{
	    	idxpathval="//input[@id='"+idxpath+"']";
	    	namexpathval="//input[@name='"+namexpath+"']";
	    	classxpathval="//input[@class='"+classxpath+"']";
	    	xpathval=idxpathval;
	    	}
	    	controlName="Checkbox";
	    	writeResult(pageKey,browserURl,controlName,idxpathval,namexpathval,classxpathval,xpathval);
	    	break;
	    	
	    case "password":
	    	idxpath=e.getAttribute("id");
	    	namexpath=e.getAttribute("name");
	    	classxpath=e.getAttribute("class");
	    	if((!(namexpath==null)) || (!(classxpath==null)) || (!(idxpath==null)))
	    	{
	    	idxpathval="//input[@id='"+idxpath+"']";
	    	namexpathval="//input[@name='"+namexpath+"']";
	    	classxpathval="//input[@class='"+classxpath+"']";
	    	xpathval=idxpathval;
	    	}
	    	controlName="password TextBox";
	    	writeResult(pageKey,browserURl, controlName,idxpathval,namexpathval,classxpathval,xpathval);
	    	break;
	    	
	    case "submit":
	    	idxpath=e.getAttribute("id");
	    	namexpath=e.getAttribute("name");
	    	classxpath=e.getAttribute("class");
	    	if((!(namexpath==null)) || (!(classxpath==null)) || (!(idxpath==null)))
	    	{
	    	idxpathval="//input[@id='"+idxpath+"']";
	    	namexpathval="//input[@name='"+namexpath+"']";
	    	classxpathval="//input[@class='"+classxpath+"']";
	    	xpathval=idxpathval;
	    	}
	    	controlName="Button";
	    	writeResult(pageKey,browserURl,controlName,idxpathval,namexpathval,classxpathval,xpathval);
	    	break;
	    	
	    case "hidden":
	    	break;
	  }
	  }
	   else if(e.getTagName().equals("select")) {
		  // System.out.println(" select");
	    }  
	  }
	 }
	
	 @Test(priority=3)
	 public void keywordTest() throws Exception
	 {
		String filename = "W:/MphasisQA/Indirani/AG_Workspace/UTFramework/src/test/resources/AutoMateDOMTesting.xls";
		Workbook wb;
		try {

		    File fileName = new File(filename);
		    Workbook Wb = WorkbookFactory.create(fileName);
			Sheet sheet = Wb.getSheet("Configure");
			rowloop: for (Row row : sheet) {
				Cell cell0 = row.getCell(2);
				Cell cells = row.getCell(0);
				if (cell0 == null) {
					continue;
				}
				String cellText = cells.getStringCellValue();
				System.out.println("*****************: "+ cellText);
				switch (cellText.toLowerCase()) {
				case "text box":
					currClass = row.getCell(2).getStringCellValue();
					System.out.println("&&&&&&&&&&&&&&&&&: "+ row.getCell(2).getStringCellValue());
					if (row.getCell(2).getStringCellValue().equalsIgnoreCase("YES")) {
							//addTextBoxMethods(cellText);
							if(addTextBoxMethods(cellText))
							{
								boolean status=true;
								System.out.println("Row value ********"+ row.getRowNum());
								System.out.println("Row value  alone ********"+ row);
								int sval = row.getRowNum();
								writeStatusResult(status,sval);
							}
					}
					break;
				case "button":
					currClass = row.getCell(2).getStringCellValue();
					if (row.getCell(2).getStringCellValue().equalsIgnoreCase("YES")) {
						if(addMethods(cellText))
						{
							boolean status=true;
							System.out.println("Row value ********"+ row.getRowNum());
							System.out.println("Row value  alone ********"+ row);
							int sval = row.getRowNum();
							writeStatusResult(status,sval);
						}
					}
					break;
				default:
					break;
				}
			}
		} catch (EncryptedDocumentException | InvalidFormatException
				| IOException e) {
			e.printStackTrace();
		}

	}
	
	 public static void InitiateURL() throws Exception
	 {
		 logger=extent.startTest("Initiate Browser");
		 orReaderTestData();
	     String browserName = browser_type;
		 String url=UrlValue;
		 if(browserName.equalsIgnoreCase("chrome")){
			   File ieFile = new File(System.getProperty("user.dir") + "\\Drivers\\chromedriver_75.exe");
			   System.setProperty("webdriver.chrome.driver", ieFile.getAbsolutePath());
			   driver = new ChromeDriver();
			   driver.get(url);
             }
         else if(browserName.equalsIgnoreCase("ie")){
      	   File ieFile = new File(System.getProperty("user.dir") + "\\Drivers\\IEDriverServer.exe");
      	   System.setProperty("webdriver.ie.driver", ieFile.getAbsolutePath());
      	   driver = new InternetExplorerDriver();
      	   driver.get(url);
         }
		 
		 driver.manage().window().maximize(); 
		 logger.log(LogStatus.INFO, "Browser Initiated ");

		 //driver.get(url);
		 Thread.sleep(5000);
		 logger.log(LogStatus.PASS, "<b>"+"Application is up and running"+ "</b>");
		 logger.log(LogStatus.INFO, "<a href=" + takeScreenshot("InitiateBrowser") +"> <img width='100' height='100' src=" + takeScreenshot("InitiateBrowser") + "> </a>");
	 }

	public static boolean addTextBoxMethods(String controlname) throws Exception
	{
		System.out.println("Control Name : " + controlname);
		boolean statuscheck=false;
		logger.log(LogStatus.INFO, "To validate TextBox basic validation in UI");
		orReader();
		WebElement ele;
		ele = driver.findElement(By.xpath(orMap.get("TextBox")));
		if(ele.isDisplayed())
		{
		logger.log(LogStatus.PASS, "<b>"+"Text Box is Present"+ "</b>");
		statuscheck=true;
		}
		if(ele.isEnabled())
		{
			statuscheck=true;
			logger.log(LogStatus.PASS, "<b>"+"Text Box is enabled"+ "</b>");
			ele.sendKeys("2367919");
			logger.log(LogStatus.PASS, "Able to Enter value in Textbox");
			logger.log(LogStatus.INFO, " <a href=" + takeScreenshot("TextBox") +"> <img width='100' height='100' src=" + takeScreenshot("TextBox") + ">");
		}
		else{
			logger.log(LogStatus.FAIL, "<b>"+"Text Box is not enabled"+ "</b>");
			statuscheck=false;
		}
		return statuscheck;
	}

	public static boolean addMethods(String controlname) throws Exception
	{
		
		boolean statuscheck=false;
		System.out.println("Control Name : " + controlname);
		logger.log(LogStatus.INFO, "To Verify Button Visibility in UI");
		
		orReader();
		//String btnName=orMap.get("Button");
		boolean dflag= driver.findElement(By.xpath(orMap.get("Button"))).isDisplayed();
		boolean eflag=driver.findElement(By.xpath(orMap.get("Button"))).isEnabled();
		String value = "Log On";
		String expvalue=driver.findElement(By.xpath(orMap.get("Button"))).getAttribute("value");
		if(expvalue.equalsIgnoreCase(value))
		{
			logger.log(LogStatus.PASS, "<b>"+"Button Name as "+expvalue+"</b>");
			logger.log(LogStatus.INFO, " <a href=" + takeScreenshot("Button") +"> <img width='100' height='100' src=" + takeScreenshot("Button") + ">");
			statuscheck=true;
			
		}
		else
		{
			logger.log(LogStatus.FAIL, "<b>"+"Button Name as "+expvalue+"</b>");
			statuscheck=false;
		}
		
		if(dflag)
		{
			logger.log(LogStatus.PASS, "<b>"+"Button is displayed "+"</b>");
			statuscheck=true;
		}
		else
		{
			logger.log(LogStatus.FAIL, "<b>"+"Button is not displayed "+"</b>");
			statuscheck=false;
		}
		if(eflag)
		{
			logger.log(LogStatus.PASS, "<b>"+"Button is Enabled "+"</b>");
			statuscheck=true;
		}
		else
		{
			logger.log(LogStatus.FAIL, "<b>"+"Button is not Enabled "+"</b>");
			statuscheck=false;
		}
		return statuscheck;
		
	}
	
    public static HashMap<String,String> orReader()
    {
          String folderName = System.getProperty("user.dir") +"\\src\\test\\resources";
          String filename = folderName + "\\AutoMateDOMTesting.xls";
		   String elementName="";
          String elementValue="";
          try{
          Workbook dl= WorkbookFactory.create(new File(filename));
          Sheet ws=dl.getSheet("CompListing");
          for(Row rw: ws)
          {
                elementName=rw.getCell(3).getStringCellValue();
                elementValue=rw.getCell(7).getStringCellValue();
                System.out.println("OR Map"+elementName + elementValue);
                orMap.put(elementName, elementValue);
          }
          }
          catch(Exception e)
          {
          e.printStackTrace();
          }
          return orMap;
    }
    
    public static void orReaderTestData() throws InvalidFormatException, IOException
    {
    	String folderName = System.getProperty("user.dir") +"\\src\\test\\resources";
        String filename = folderName + "\\AutoMateDOMTesting.xls";
  	    Workbook dl= WorkbookFactory.create(new File(filename));
         Sheet ws=dl.getSheet("Main");
         boolean firstRow = true;
         
         for(Row rw: ws) {
       	if (firstRow) {
             firstRow = false;
             continue;
           }   
     	   DataFormatter formatter = new DataFormatter();
     	   Cell cell = rw.getCell(2);
     	   browser_type = formatter.formatCellValue(cell);
     	   System.out.println(browser_type);
     	   Cell url_value = rw.getCell(3);
     	   UrlValue = formatter.formatCellValue(url_value);
     	   System.out.println(UrlValue);
         }
    }
    
	public static void writeResult(int pageKey,String browserURl, String controlName, String idval, String namexpathval, String classxpathval,String xpathval) throws Exception {
		
	    String fileLocation = System.getProperty("user.dir") + "//src//test//resources//";
	    String fileNameForm = fileLocation + "AutoMateDOMTesting.xls";
	    //int rowCount;
	    File f = new File(fileNameForm);
	    if (f.exists()) {
	      FileInputStream fileOut = new FileInputStream(new File(fileNameForm));
	      HSSFWorkbook workbook = new HSSFWorkbook(fileOut);
	      HSSFSheet worksheet = workbook.getSheet("CompListing");
	      int row1 = worksheet.getPhysicalNumberOfRows();
	      int rs=row1-1;
	      int rowval = rs + 1;
	      
	      Row row = worksheet.createRow(rowval);
	      HSSFCell cellpagekey = (HSSFCell) row.createCell(0);
	      cellpagekey.setCellType(Cell.CELL_TYPE_STRING);
	      cellpagekey.setCellValue(pageKey);
	      
	      HSSFCell cellbrowser = (HSSFCell) row.createCell(1);
	      cellbrowser.setCellType(Cell.CELL_TYPE_STRING);
	      cellbrowser.setCellValue(browserURl);
	     
	      HSSFCell cellctrl = (HSSFCell) row.createCell(3);
	      cellctrl.setCellType(Cell.CELL_TYPE_STRING);
	      cellctrl.setCellValue(controlName);
	      
	      HSSFCell cellid = (HSSFCell) row.createCell(4);
	      cellid.setCellType(Cell.CELL_TYPE_STRING);
	      cellid.setCellValue(idval);
	
	      HSSFCell cellname = (HSSFCell) row.createCell(5);
	      cellname.setCellType(Cell.CELL_TYPE_STRING);
	      cellname.setCellValue(namexpathval);
	      
	      HSSFCell cellclass = (HSSFCell) row.createCell(6);
	      cellclass.setCellType(Cell.CELL_TYPE_STRING);
	      cellclass.setCellValue(classxpathval);
	      
	      HSSFCell cellxpath = (HSSFCell) row.createCell(7);
	      cellxpath.setCellType(Cell.CELL_TYPE_STRING);
	      cellxpath.setCellValue(xpathval);
	      
	      fileOut.close();
	      FileOutputStream outFile = new FileOutputStream(new File(fileNameForm));
	      workbook.write(outFile);
	      workbook.close();
	      outFile.close();
	    }
	  }
	
	
	public static void writeStatusResult(boolean status,int sval) throws Exception {
	    String fileLocation = System.getProperty("user.dir") + "//src//test//resources//";
	    String fileNameForm = fileLocation + "AutoMateDOMTesting.xls";
	    //int rowCount;
	    File f = new File(fileNameForm);
	    if (f.exists()) {
	      FileInputStream fileOut = new FileInputStream(new File(fileNameForm));
	      HSSFWorkbook workbook = new HSSFWorkbook(fileOut);
	      HSSFSheet worksheet = workbook.getSheet("Configure");
	      //int row1 = worksheet.getPhysicalNumberOfRows();
	     System.out.println("Write Status Result : "+ sval  +"####"+ status);
	      if(status==true)
	      {
	      Row row = worksheet.getRow(sval);
	      HSSFCell cell8 = (HSSFCell) row.createCell(3);
	      cell8.setCellType(Cell.CELL_TYPE_STRING);
	      cell8.setCellValue("PASS");
	      }
	      else
	      {
	      Row row = worksheet.getRow(sval);
	      HSSFCell cell8 = (HSSFCell) row.createCell(3);
	      cell8.setCellValue("FAIL");
	      }
	      
	      fileOut.close();
	      FileOutputStream outFile = new FileOutputStream(new File(fileNameForm));
	      workbook.write(outFile);
	      workbook.close();
	      outFile.close();
	    }
	  }
	
	public static String takeScreenshot(String screenName) throws HeadlessException, AWTException, IOException
	{
		String screenshotFolderPath= System.getProperty("user.dir")+"\\elementScreen";
		BufferedImage image = new Robot().createScreenCapture(new Rectangle(Toolkit.getDefaultToolkit().getScreenSize()));
		ImageIO.write(image, "png", new File(screenshotFolderPath+"\\"+screenName+".png")); 
		String strscreen =screenshotFolderPath+"\\"+screenName+".png";
		return strscreen;
	}		
	 
	 @AfterTest
	 public void endReport() throws Exception {
		 //tearDown();
	     extent.endTest(logger);
	     extent.flush();
	 }
/*	 public void tearDown() throws Exception {
		 driver.close();
		 logger.log(LogStatus.INFO, "Browser Closed");
		 driver.quit();
	 }*/
	}

