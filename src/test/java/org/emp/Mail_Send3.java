package org.emp;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.ListIterator;
import java.util.Properties;
import java.util.Set;

import javax.activation.DataHandler;
import javax.activation.DataSource;
import javax.activation.FileDataSource;
import javax.mail.MessagingException;
import javax.mail.Multipart;
import javax.mail.Session;
import javax.mail.Transport;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.Test;

import io.github.bonigarcia.wdm.WebDriverManager;

public class Mail_Send3 {
	
	static WebDriver driver;
    static Row row1;
    static int size;
	static Row r;

//	public static void main(String[] args) throws InterruptedException, Exception {
		
	@Test
	public void BrowserLaunch() throws InterruptedException, IOException{
	WebDriverManager.chromedriver().setup();
	WebDriver driver=new ChromeDriver();
	driver.get("https://kedge.srinsofttech.com/#/kn-pages/auth/login");
	driver.manage().window().maximize();
	
	
	WebElement uname = driver.findElement(By.xpath("//input[@placeholder='Employee id']"));
	uname.sendKeys("SS1407");
	
	WebElement Password = driver.findElement(By.xpath("//input[@placeholder='Password']"));
	Password.sendKeys("Jul@2022");
	
	WebElement Login = driver.findElement(By.xpath("//button[@type='submit']"));
	Login.click();
	
	WebDriverWait waits=new WebDriverWait(driver,30);
	waits.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//span[text()='LMS']")));
	WebElement LMS = driver.findElement(By.xpath("//span[text()='LMS']"));
	LMS.click();
	
	String handle = driver.getWindowHandle();	
    
    Set<String> allwindow = driver.getWindowHandles();

    for (String eachwindow : allwindow) {
    if (!eachwindow.equals(handle)) {
    	driver.switchTo().window(eachwindow);
		
	}
		
	} 
	WebElement Punch = driver.findElement(By.xpath("(//span[text()='Punch Details'])[1]"));
	Punch.click();	   
	
	// List<WebElement>allwindow2 = driver.findElements(By.xpath("//table[@class='table']//tr//th"));
	 List<WebElement> Headers = driver.findElements(By.xpath("//table[@class='table']//tr//th"));
			Workbook w =new XSSFWorkbook();
		    Sheet a=w.createSheet();
			Row row = a.createRow(0);
			
			a.setDisplayGridlines(true);
		    
			Row rows = a.createRow(0);
		    Cell cell=rows.createCell(0);
	        cell.setCellValue("Emp ID");
		    Cell cell2=rows.createCell(1);
	        cell2.setCellValue("Emp Name");
	        	    
			CellStyle style=  w.createCellStyle();
			style.setFillBackgroundColor(IndexedColors.YELLOW.getIndex());
			style.setFillPattern(FillPatternType.FINE_DOTS);
			cell.setCellStyle(style);
			cell2.setCellStyle(style);

//			style.setAlignment(CellStyle.ALIGN_CENTER);
			Font font= w.createFont();
//		    font.setBold(true);
			style.setWrapText(true);
	        style.setFont(font);
	        font.setBold(true);
			
	        ListIterator<WebElement> li = Headers.listIterator();
		     
			 for (int i = 0; i <Headers.size(); i++) {
				 Cell createCell = rows.createCell(i);
				 createCell.setCellStyle(style);
				 a.setColumnWidth(i, 5000);
				 createCell.setCellValue(li.next().getText());
			 }
			

//			ListIterator<WebElement> lii = allwindow2.listIterator();
//			for (int i = 0; i <allwindow2.size(); i++) {
//				Cell cell3=row.createCell(i);
//				cell3.setCellStyle(style);
//				cell3.setCellValue(lii.next().getText());
//			//	row.createCell(i).setCellValue(lii.next().getText());	
//			}
			
			Thread.sleep(5000);
			
			WebDriverWait waits2=new WebDriverWait(driver,30);
			waits.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//table[@class='table']//tbody//tr//td[1]")));
			List<WebElement> allrow1 = driver.findElements(By.xpath("//table[@class='table']//tbody//tr//td[1]"));
			List<WebElement> allrow2 = driver.findElements(By.xpath("//table[@class='table']//tbody//tr//td[2]"));
			List<WebElement> allrow3 = driver.findElements(By.xpath("//table[@class='table']//tbody//tr//td[3]"));
			List<WebElement> allrow4 = driver.findElements(By.xpath("//table[@class='table']//tbody//tr//td[4]"));
			List<WebElement> allrow5 = driver.findElements(By.xpath("//table[@class='table']//tbody//tr//td[5]"));
			List<WebElement> allrow6 = driver.findElements(By.xpath("//table[@class='table']//tbody//tr//td[6]"));
			List<WebElement> allrow7 = driver.findElements(By.xpath("//table[@class='table']//tbody//tr//td[7]"));
			List<WebElement> allrow8 = driver.findElements(By.xpath("//table[@class='table']//tbody//tr//td[8]"));
			List<WebElement> allrow9 = driver.findElements(By.xpath("//table[@class='table']//tbody//tr//td[9]"));
			List<WebElement> allrow10 = driver.findElements(By.xpath("//table[@class='table']//tbody//tr//td[10]"));

		    ListIterator<WebElement> li1=allrow1.listIterator();
		    ListIterator<WebElement> li2=allrow2.listIterator();
		    ListIterator<WebElement> li3=allrow3.listIterator();
		    ListIterator<WebElement> li4=allrow4.listIterator();
		    ListIterator<WebElement> li5=allrow5.listIterator();
		    ListIterator<WebElement> li6=allrow6.listIterator();
		    ListIterator<WebElement> li7=allrow7.listIterator();
		    ListIterator<WebElement> li8=allrow8.listIterator();
		    ListIterator<WebElement> li9=allrow9.listIterator();
		    ListIterator<WebElement> li10=allrow10.listIterator();
	    
			for (int j = 1; j < allrow1.size(); j++) {
				 Row row1=a.createRow(j);
			row1.createCell(0).setCellValue(li1.next().getText());
		    row1.createCell(1).setCellValue(li2.next().getText());
			row1.createCell(2).setCellValue(li3.next().getText());
			row1.createCell(3).setCellValue(li4.next().getText());
			row1.createCell(4).setCellValue(li5.next().getText());
			row1.createCell(5).setCellValue(li6.next().getText());
			row1.createCell(6).setCellValue(li7.next().getText());
			row1.createCell(7).setCellValue(li8.next().getText());
			row1.createCell(8).setCellValue(li9.next().getText());					
			row1.createCell(9).setCellValue(li10.next().getText());	
			}
			Thread.sleep(5000);
			FileOutputStream f = new FileOutputStream("C:\\Users\\arun.kumar\\OneDrive - SrinSoft Technologies Pvt Ltd\\Documents\\Excel\\punchdetails2.xlsx");
			w.write(f);
			f.close();}

			String Email_Id = "smtpsstech@srinsofttech.com";

			String host_name = "smtp.office365.com";

			// String password = "srinsoft@2022";
			String filename = ("C:\\Users\\arun.kumar\\OneDrive - SrinSoft Technologies Pvt Ltd\\Documents\\Excel\\punchdetails2.xlsx");
			String password = "$rin@$oft#098";

			String from = "smtpsstech@srinsofttech.com";

			//String Username = "arun.kumar";
	
			@Test
			public void sendmail() throws Exception, MessagingException {

				StringBuilder email = new StringBuilder();

				email.append("<html><body><table><tr><td>Bubu<td>Lala</tr></table></body></html>");

				email.append("Hi All,");

				// email.append("<br />" + "<br />");

				// email.append("<font face=\"Trebuchet MS\"><font size=\"3\">" +
				// "Shapiro Velocity Report:" + "</face></font>");

				// email.append("<br />");

				////////////////////// Table
				email.append("<html><body><table><tr><td>Bubu<td>Lala</tr></table></body></html>");

				email.append("Velocity Table:");

				// email.append("")

				// mail_part(fid);
				email.append("<html><body><table><tr><td>Bubu<td>Lala</tr></table></body></html>");

				email.append("Thanks,");

				email.append("<html><body><table><tr><td>Bubu<td2>Lala</tr></table></body></html>");

				// email.append("<br />");

				email.append("Automation team");

				Properties properties = System.getProperties();

				properties.put("mail.smtp.host", host_name);

				properties.put("mail.smtp.user", Email_Id);

				properties.put("mail.smtp.password", password);

				properties.put("mail.smtp.ssl.protocols", "TLSv1.2");

				properties.put("mail.smtp.port", "587");

				properties.put("mail.smtp.auth", "true");

				properties.put("mail.smtp.starttls.enable", "true");

				Session session = Session.getDefaultInstance(properties);

				MimeMessage message = new MimeMessage(session);

				message.setFrom(new InternetAddress(from));

				message.setContent(email.toString(), "text/html");

				// message.addRecipient(RecipientType.TO, new
				// InternetAddress(naveenkumar.k@srinsofttech.com));

				message.addRecipients(javax.mail.Message.RecipientType.TO, "2504arunkumar@gmail.com");

				// message.addRecipients(Message.RecipientType.CC, Email_1);

				message.setSubject("Shapiro Velocity Report");
				MimeBodyPart messagebodypart2 = new MimeBodyPart();
				messagebodypart2.setText(email.toString());

				message.setContent(email.toString(), "text/html");

				Multipart multipart = new MimeMultipart();
				// BodyPart messagebodypart = new MimeBodyPart();
				multipart.addBodyPart(messagebodypart2);
				message.setContent(multipart);
				message.setContent(email.toString(), "text/html");

				//messagebodypart2  = new MimeBodyPart();
				String filenames = "C:\\Users\\arun.kumar\\OneDrive - SrinSoft Technologies Pvt Ltd\\Documents\\Excel\\punchdetails2.xlsx";
				DataSource source = new FileDataSource(filenames);
				messagebodypart2.setDataHandler(new DataHandler(source));
				messagebodypart2.setFileName(filenames);
				multipart.addBodyPart(messagebodypart2);

				// Send the complete message parts
				message.setContent(multipart);

				Transport transport = session.getTransport("smtp");

				transport.connect(host_name, Email_Id, password);

				transport.sendMessage(message, message.getAllRecipients());

				transport.close();
			}
}