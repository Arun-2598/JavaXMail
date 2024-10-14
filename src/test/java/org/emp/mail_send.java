package org.emp;

import java.sql.PreparedStatement;
import java.util.Properties;

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
import javax.swing.JTable;
import javax.swing.table.DefaultTableModel;

import org.testng.annotations.Test;

public class mail_send {
	// String Email_Id = "arun.kumar@srinsofttech.com";

	String Email_Id = "smtpsstech@srinsofttech.com";

	String host_name = "smtp.office365.com";

	// String password = "srinsoft@2022";
	String filename = ("C:\\Users\\arun.kumar\\OneDrive - SrinSoft Technologies Pvt Ltdmail\\Documents\\Java\\This.txt");
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
		String filename = "C:\\Users\\arun.kumar\\Downloads\\punchdetails.xlsx";
		DataSource source = new FileDataSource(filename);
		messagebodypart2.setDataHandler(new DataHandler(source));
		messagebodypart2.setFileName(filename);
		multipart.addBodyPart(messagebodypart2);

		// Send the complete message parts
		message.setContent(multipart);

		Transport transport = session.getTransport("smtp");

		transport.connect(host_name, Email_Id, password);

		transport.sendMessage(message, message.getAllRecipients());

		transport.close();
	}

	public static void main(String[] args) throws MessagingException, Exception {
		mail_send m = new mail_send();
		m.sendmail();
	}

	// public static void main(String[] args) throws MessagingException,
	// Exception {
	// mail_send m = new mail_send();
	// m.sendmail();
	// }

}