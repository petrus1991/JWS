

import java.io.File;
import java.io.IOException;
import java.util.List;
import java.util.Random;
import java.awt.AWTException;
import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.remote.DesiredCapabilities;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.ExpectedCondition;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

/*----------------------------------------------------------------------------------
 * File name 		: Register.java									
 * Description 		: For Jakarta Web Solutions
 * Created On		: 02-Jan-2021
 * Created By 		: Petrus
 * ---------------------------------------------------------------------------------
 * No. 	Date 			By			History 			
 * ---------------------------------------------------------------------------------
 * 1. 	02-Jan-2021		Petrus	 	Script Automation Created
--------------------------------------------------------------------------------------*/

public class Register  {
	public static void main(String[] args) throws IOException, InterruptedException, AWTException{
        // FIREFOX DRIVER
		System.setProperty("webdriver.gecko.driver", "C:/Users/Petrus/Downloads/Compressed/JAVA_IDE/geckodriver.exe");
    	DesiredCapabilities capabilities = new DesiredCapabilities();
    	capabilities.setCapability("acceptInsecureCerts", true);
    	WebDriver driver = new FirefoxDriver(capabilities);

        // NAVIGATE TO LINK
    	driver.navigate().to("http://automationpractice.com");

        //CREATE USER
        Register.CreateUser(driver);
        System.out.println("-----------------FINISH-----------------");
	}
	
	  	
	public static void CreateUser(WebDriver driver) throws InterruptedException, IOException, AWTException{
		String path = "C:/Users/Petrus/Downloads/Compressed/JAVA_IDE/Register.xls";
		File inputWorkbook = new File(path);
		Workbook w;
	
  	    try {
  	    	w = Workbook.getWorkbook(inputWorkbook);
  	    	WebDriverWait wait = new WebDriverWait(driver, 300);
  	      
	  	    //SHEET
	  	    Sheet sheet = w.getSheet(0); //DATA USER AMBIL DARI EXCEL SHEET PERTAMA
	  	    int nRow = sheet.getRows();
	  	    int nColumn = sheet.getColumns();
	  	    System.out.println("Jumlah Kolom : " +nColumn);
	  	    System.out.println("Jumlah Baris : " +nRow);
	  	    System.out.println("-----------------START-----------------");
	  	    
	  	    //WAITING SIGN IN LINK/BUTTON
	  	    wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//a[contains(text(),'Sign in')]")));
	  	    
	  	    //WAITING PAGE FULLY LOADED
	  	    wait.until(new ExpectedCondition<Boolean>() {
	  	    	public Boolean apply(WebDriver driver) {
	  	    		 return ((JavascriptExecutor)driver).executeScript("return document.readyState").equals("complete");
	  	    	}
	  	    });
	  	    
	  	    //CLICK SIGN IN LINK/BUTTON
	  	    driver.findElement(By.xpath("//a[contains(text(),'Sign in')]")).click();
	  		
	  	    for(int n=2; n<nRow; n++){ //INI AKAN LOOPING SAMPAI DATA ROW TERAKHIR DI EXCEL
		  	    //WAITING FIELD
	  	    	wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("email_create")));
		  	  
		  	    //WAITING PAGE FULLY LOADED
		  	    wait.until(new ExpectedCondition<Boolean>() {
		  	    	public Boolean apply(WebDriver driver) {
		  	    		 return ((JavascriptExecutor)driver).executeScript("return document.readyState").equals("complete");
		  	    	}
		  	    });
		  	    
		  	    System.out.println("================================== DATA KE-"+(n-1)+" ==================================");
		  	    
		  	    //EMAIL
	  	    	WebElement email_create = driver.findElement(By.id("email_create"));
	  	    	Cell cellemail_create = sheet.getCell(0,n);
	  	    	//JIKA EMAIL DIISI DI EXCEL MAKA DIAMBIL DARI EXCEL
	  	    	if(cellemail_create.getContents() != null && !cellemail_create.getContents().equals("")){
		  	    	System.out.println("Email : "+cellemail_create.getContents());
		  	    	email_create.sendKeys(Keys.chord(Keys.CONTROL,"a"));
		  	    	email_create.sendKeys(cellemail_create.getContents());	
	  	    	}
	  	    	//JIKA EMAIL TIDAK DIISI, MAKA AKAN GENERATE RANDOMLY UNTUK EMAILNYA.
	  	    	else{
	  	    		String sRandomEmail = generateRandomChars(15);
	  	    		System.out.println("Random Email : " +sRandomEmail);
		  	    	email_create.sendKeys(Keys.chord(Keys.CONTROL,"a"));
		  	    	email_create.sendKeys(sRandomEmail);	
	  	    	}
		  	    
	  	    	
	  	    	//WAITING BUTTON CREATE ACCOUNT
	  	    	wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//button[@id='SubmitCreate']/span")));
	  	    	
	  	    	//SLEEP
	  	    	Thread.sleep(500);
	  	    	
	  	    	//CLICK BUTTON CREATE ACCOUNT
	  	    	driver.findElement(By.xpath("//button[@id='SubmitCreate']/span")).click();
	  	    	
	  	    	//WAITING FIELD 
	  	    	wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("customer_firstname")));
			  	  
		  	    //WAITING PAGE FULLY LOADED
		  	    wait.until(new ExpectedCondition<Boolean>() {
		  	    	public Boolean apply(WebDriver driver) {
		  	    		 return ((JavascriptExecutor)driver).executeScript("return document.readyState").equals("complete");
		  	    	}
		  	    });	
		  	    
		  	    //TITLE
	  	    	List<WebElement> id_gender = driver.findElements(By.name("id_gender"));
	  	    	Cell cellid_gender = sheet.getCell(1,n);
	  	    	if(cellid_gender.getContents() != null && !cellid_gender.getContents().equals("")){
		  	    	System.out.println("Title : "+cellid_gender.getContents());
		  	    	String sTitle;
		  	    	if(cellid_gender.getContents().equalsIgnoreCase("Mr")){
		  	    		sTitle = "1";
		  	    	}
		  	    	else{
		  	    		sTitle = "2";
		  	    	}
		  	    	System.out.println("sTitle : "+sTitle);
		  	    	int iGenderSize = id_gender.size();
		  	    	for(int a=0; a < iGenderSize ; a++ ){
		  	    		String sValue = id_gender.get(a).getAttribute("value");
		  	    		if (sValue.equalsIgnoreCase(sTitle)){
		  	    			id_gender.get(a).click();
		  	    			break;
		  	    		}
		  	    	}
	  	    	}
		  	    
		  	    //FIRST NAME
	  	    	WebElement customer_firstname = driver.findElement(By.id("customer_firstname"));
	  	    	Cell cellcustomer_firstname = sheet.getCell(2,n);
	  	    	if(cellcustomer_firstname.getContents() != null && !cellcustomer_firstname.getContents().equals("")){
		  	    	System.out.println("First Name : "+cellcustomer_firstname.getContents());
		  	    	customer_firstname.sendKeys(Keys.chord(Keys.CONTROL,"a"));
		  	    	customer_firstname.sendKeys(cellcustomer_firstname.getContents());	
	  	    	}
		  	    
		  	    //LAST NAME
	  	    	WebElement customer_lastname = driver.findElement(By.id("customer_lastname"));
	  	    	Cell cellcustomer_lastname = sheet.getCell(3,n);
	  	    	if(cellcustomer_lastname.getContents() != null && !cellcustomer_lastname.getContents().equals("")){
		  	    	System.out.println("Last Name : "+cellcustomer_lastname.getContents());
		  	    	customer_lastname.sendKeys(Keys.chord(Keys.CONTROL,"a"));
		  	    	customer_lastname.sendKeys(cellcustomer_lastname.getContents());	
	  	    	}
	  	    	
		  	    //PASSWORD
	  	    	WebElement passwd = driver.findElement(By.id("passwd"));
	  	    	Cell cellpasswd = sheet.getCell(4,n);
	  	    	if(cellpasswd.getContents() != null && !cellpasswd.getContents().equals("")){
		  	    	System.out.println("Password : "+cellpasswd.getContents());
		  	    	passwd.sendKeys(Keys.chord(Keys.CONTROL,"a"));
		  	    	passwd.sendKeys(cellpasswd.getContents());	
	  	    	}
	  	    	
		  	    //DATE OF BIRTH
		  	    Cell cellDOB = sheet.getCell(5,n);
		  	    if(cellDOB.getContents() != null && !cellDOB.getContents().equals("")){
		  	    	String cellYear = cellDOB.getContents().substring(6,10);
		  	    	String cellMonth = cellDOB.getContents().substring(3,5);
		  	    	String cellDays = cellDOB.getContents().substring(0,2);
		  	    	System.out.println("Date Of Birth : "+cellDays+"-"+cellMonth+"-"+cellYear);
		  	    	
		  	    	//ISI TANGGAL 
		  	    	cellDays = DateSelect(cellDays);
			  	    Select days = new Select(driver.findElement(By.id("days")));
			  	    days.selectByValue(cellDays);
		  	    	
		  	    	//ISI BULAN
		  	    	cellMonth = MonthSelect(cellMonth);
		  	    	Select months = new Select(driver.findElement(By.id("months")));
		  	    	months.selectByVisibleText(cellMonth);
			  	    
		  	    	//ISI TAHUN
			  	    Select years = new Select(driver.findElement(By.id("years")));
			  	    years.selectByValue(cellYear);

		  	    }
		  	    
		  	    //SIGN UP FOR OUR NEWS LETTER?
		  	    String isNewsletter = "N";
		  	    WebElement newsletter = driver.findElement(By.id("newsletter"));
		  	    Cell cellnewsletter = sheet.getCell(6,n);
		  	    if(cellnewsletter.getContents() != null && !cellnewsletter.getContents().equals("")){
			  	    System.out.println("News Letter? : "+cellnewsletter.getContents());
		  	    	if(cellnewsletter.getContents().equalsIgnoreCase("Y")){
		  	    		isNewsletter = "Y";
			  	    	if (!newsletter.isSelected()){
			  	    		newsletter.click();
						}
			  	    }
			  	    else{
			  	    	if (newsletter.isSelected()){
			  	    		newsletter.click();	
						}
		  	    	}
		  	    }
	  	    	
		  	    //RECEIVE SPECIAL OFFERS FROM OUR PARTNER
		  	    String isOptin = "N";
		  	    WebElement optin = driver.findElement(By.id("optin"));
		  	    Cell celloptin = sheet.getCell(7,n);
		  	    if(celloptin.getContents() != null && !celloptin.getContents().equals("")){
			  	    System.out.println("Special offers? : "+celloptin.getContents());
		  	    	if(celloptin.getContents().equalsIgnoreCase("Y")){
		  	    		isOptin = "Y";
			  	    	if (!optin.isSelected()){
			  	    		optin.click();
						}
			  	    }
			  	    else{
			  	    	if (optin.isSelected()){
			  	    		optin.click();	
						}
		  	    	}
		  	    }
		  	    
		  	    //ADDRESS FIRST NAME
	  	    	WebElement firstname = driver.findElement(By.id("firstname"));
	  	    	Cell cellfirstname = sheet.getCell(8,n);
	  	    	if(cellfirstname.getContents() != null && !cellfirstname.getContents().equals("")){
		  	    	System.out.println("Address First Name : "+cellfirstname.getContents());
		  	    	firstname.sendKeys(Keys.chord(Keys.CONTROL,"a"));
		  	    	firstname.sendKeys(cellfirstname.getContents());	
	  	    	}
		  	    
		  	    //ADDRESS LAST NAME
	  	    	WebElement lastname = driver.findElement(By.id("lastname"));
	  	    	Cell celllastname = sheet.getCell(9,n);
	  	    	if(celllastname.getContents() != null && !celllastname.getContents().equals("")){
		  	    	System.out.println("Address Last Name : "+celllastname.getContents());
		  	    	lastname.sendKeys(Keys.chord(Keys.CONTROL,"a"));
		  	    	lastname.sendKeys(celllastname.getContents());	
	  	    	}
	  	    	
		  	    //COMPANY
	  	    	WebElement company = driver.findElement(By.id("company"));
	  	    	Cell cellcompany = sheet.getCell(10,n);
	  	    	if(cellcompany.getContents() != null && !cellcompany.getContents().equals("")){
		  	    	System.out.println("Company : "+cellcompany.getContents());
		  	    	company.sendKeys(Keys.chord(Keys.CONTROL,"a"));
		  	    	company.sendKeys(cellcompany.getContents());	
	  	    	}
	  	    	
	  	    	//ADDRESS1
	  	    	WebElement address1 = driver.findElement(By.id("address1"));
	  	    	Cell celladdress1 = sheet.getCell(11,n);
	  	    	if(celladdress1.getContents() != null && !celladdress1.getContents().equals("")){
		  	    	System.out.println("Address 1 : "+celladdress1.getContents());
		  	    	address1.sendKeys(Keys.chord(Keys.CONTROL,"a"));
		  	    	address1.sendKeys(celladdress1.getContents());	
	  	    	}
	  	    	
	  	    	//ADDRESS2
	  	    	WebElement address2 = driver.findElement(By.id("address2"));
	  	    	Cell celladdress2 = sheet.getCell(12,n);
	  	    	if(celladdress2.getContents() != null && !celladdress2.getContents().equals("")){
		  	    	System.out.println("Address 2 : "+celladdress2.getContents());
		  	    	address2.sendKeys(Keys.chord(Keys.CONTROL,"a"));
		  	    	address2.sendKeys(celladdress2.getContents());	
	  	    	}
	  	    	
	  	    	//CITY
	  	    	WebElement city = driver.findElement(By.id("city"));
	  	    	Cell cellcity = sheet.getCell(13,n);
	  	    	if(cellcity.getContents() != null && !cellcity.getContents().equals("")){
		  	    	System.out.println("City : "+cellcity.getContents());
		  	    	city.sendKeys(Keys.chord(Keys.CONTROL,"a"));
		  	    	city.sendKeys(cellcity.getContents());	
	  	    	}
	  	    	
	  	    	//COUNTRY
	  	    	boolean bCountry = false;
		  	    Select id_country = new Select(driver.findElement(By.id("id_country")));
		  	    Cell cellid_country = sheet.getCell(14,n);
		  	    if(cellid_country.getContents() != null && !cellid_country.getContents().equals("")){
		  	    	System.out.println("id_country : "+cellid_country.getContents());
		  	    	id_country.selectByVisibleText(cellid_country.getContents());
		  	    	if(cellid_country.getContents().equalsIgnoreCase("UNITED STATES")){
		  	    		bCountry = true;
		  	    	}
		  	    }else{
		  	    	bCountry = true;
		  	    }
		  	    
		  	    //JIKA COUNTRY TIDAK DIINPUT / DIINPUT UNITED STATES MAKA HARUS INPUT STATE SAMA POSTAL CODE
		  	    if(bCountry){
		  	    	System.out.println("INPUT STATE AND POSTAL CODE");
		  	    	
		  	    	//SLEEP
		  	    	Thread.sleep(500);
		  	    	
		  	    	//STATE
			  	    Select id_state = new Select(driver.findElement(By.id("id_state")));
			  	    Cell cellid_state = sheet.getCell(15,n);
			  	    if(cellid_state.getContents() != null && !cellid_state.getContents().equals("")){
			  	    	System.out.println("id_state : "+cellid_state.getContents());
			  	    	id_state.selectByVisibleText(cellid_state.getContents());
			  	    }
			  	    
			  	    //POSTAL CODE
		  	    	WebElement postcode = driver.findElement(By.id("postcode"));
		  	    	Cell cellpostcode = sheet.getCell(16,n);
		  	    	if(cellpostcode.getContents() != null && !cellpostcode.getContents().equals("")){
			  	    	System.out.println("postcode : "+cellpostcode.getContents());
			  	    	postcode.sendKeys(Keys.chord(Keys.CONTROL,"a"));
			  	    	postcode.sendKeys(cellpostcode.getContents());	
		  	    	}
		  	    }
		  	    
		  	    //ADDITIONAL INFORMATION
	  	    	WebElement other = driver.findElement(By.id("other"));
	  	    	Cell cellother = sheet.getCell(17,n);
	  	    	if(cellother.getContents() != null && !cellother.getContents().equals("")){
		  	    	System.out.println("other : "+cellother.getContents());
		  	    	other.sendKeys(Keys.chord(Keys.CONTROL,"a"));
		  	    	other.sendKeys(cellother.getContents());	
	  	    	}
	  	    	
	  	    	//HOME PHONE
	  	    	WebElement phone = driver.findElement(By.id("phone"));
	  	    	Cell cellphone = sheet.getCell(18,n);
	  	    	if(cellphone.getContents() != null && !cellphone.getContents().equals("")){
		  	    	System.out.println("phone : "+cellphone.getContents());
		  	    	phone.sendKeys(Keys.chord(Keys.CONTROL,"a"));
		  	    	phone.sendKeys(cellphone.getContents());	
	  	    	}
	  	    	
	  	    	//MOBILE PHONE
	  	    	WebElement phone_mobile = driver.findElement(By.id("phone_mobile"));
	  	    	Cell cellphone_mobile = sheet.getCell(19,n);
	  	    	if(cellphone_mobile.getContents() != null && !cellphone_mobile.getContents().equals("")){
		  	    	System.out.println("phone_mobile : "+cellphone_mobile.getContents());
		  	    	phone_mobile.sendKeys(Keys.chord(Keys.CONTROL,"a"));
		  	    	phone_mobile.sendKeys(cellphone_mobile.getContents());	
	  	    	}
	  	    	
	  	    	//ASSIGN AN ADRESS ALIAS FOR FUTURE REFERENCE
	  	    	WebElement alias = driver.findElement(By.id("alias"));
	  	    	Cell cellalias = sheet.getCell(20,n);
	  	    	if(cellalias.getContents() != null && !cellalias.getContents().equals("")){
		  	    	System.out.println("alias : "+cellalias.getContents());
		  	    	alias.sendKeys(Keys.chord(Keys.CONTROL,"a"));
		  	    	alias.sendKeys(cellalias.getContents());	
	  	    	}
	  	    	
	  	    	//WAITING BUTTON REGISTER
	  	    	wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//button[@id='submitAccount']/span")));
	  	    	
	  	    	//SLEEP
	  	    	Thread.sleep(1000);
	  	    	
	  	    	//CLICK BUTTON REGISTER
	  	    	driver.findElement(By.xpath("//button[@id='submitAccount']/span")).click();
	  	    	
		  	    //WAITING SIGN OUT LINK/BUTTON
		  	    wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//a[contains(text(),'Sign out')]")));
			  	  
		  	    //WAITING PAGE FULLY LOADED
		  	    wait.until(new ExpectedCondition<Boolean>() {
		  	    	public Boolean apply(WebDriver driver) {
		  	    		 return ((JavascriptExecutor)driver).executeScript("return document.readyState").equals("complete");
		  	    	}
		  	    });	
	  	    	
	  	    	//SLEEP
	  	    	Thread.sleep(1000);
	  	    	
		  	    //CLICK SIGN OUT LINK/BUTTON
		  	    driver.findElement(By.xpath("//a[contains(text(),'Sign out')]")).click();
	  	    	
		    	//WAITING BACKGROUND PROCESS
		  	    if(n+1 != nRow){
		  	    	System.out.println("Redirecting new row..");
		  	    	Thread.sleep(1000);
		  	    }
	  	    }
  	    }
  	    catch (BiffException e) {
  	    	e.printStackTrace();
  	    }		  	
	
	}
	
	public static String MonthSelect(String Month){
		String rtn = "";
		switch (Month){
			case "01": rtn = "January "; break;
			case "02": rtn = "Febuary "; break;
			case "03": rtn = "March "; break;
			case "04": rtn = "April "; break;
			case "05": rtn = "May "; break;
			case "06": rtn = "June "; break;
			case "07": rtn = "July "; break;
			case "08": rtn = "August "; break;
			case "09": rtn = "September "; break;
			case "10": rtn = "October "; break;
			case "11": rtn = "November "; break;
			case "12": rtn = "December "; break;
		}
		return rtn;
	}
	  
	public static String DateSelect(String Date){
		String rtn = "";
		switch (Date){
			case "01": rtn = "1"; break;
			case "02": rtn = "2"; break;
			case "03": rtn = "3"; break;
			case "04": rtn = "4"; break;
			case "05": rtn = "5"; break;
			case "06": rtn = "6"; break;
			case "07": rtn = "7"; break;
			case "08": rtn = "8"; break;
			case "09": rtn = "9"; break;
			default: rtn = Date;
		}
		return rtn;
	}
	
	
	public static String generateRandomChars(int length) {
	    StringBuilder sb = new StringBuilder();
	    Random random = new Random();
	    String candidateChars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz1234567890";
	    for (int i=0; i<length; i++) {
	        sb.append(candidateChars.charAt(random.nextInt(candidateChars.length())));
	    }
	    sb.append("@gmail.com");
	    
	    return sb.toString();
	}
}
