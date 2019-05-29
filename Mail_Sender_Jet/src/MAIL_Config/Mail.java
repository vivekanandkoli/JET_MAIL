package MAIL_Config;

import java.io.IOException;
import java.lang.management.ManagementFactory;
import java.lang.management.OperatingSystemMXBean;
import java.lang.reflect.Method;
import java.net.InetAddress;
import java.net.UnknownHostException;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.Properties;

import javax.activation.DataHandler;
import javax.activation.DataSource;
import javax.activation.FileDataSource;
import javax.mail.BodyPart;
import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.Transport;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;
import javax.management.AttributeNotFoundException;
import javax.management.InstanceNotFoundException;
import javax.management.MBeanException;
import javax.management.MBeanServer;
import javax.management.MalformedObjectNameException;
import javax.management.ObjectName;
import javax.management.ReflectionException;


public class Mail 
{
	private javax.mail.Session session;  
	
	WriteExcel rw=new WriteExcel();
	String Status=null;
	public static void main(String[] args) 
	{
		
	}
	

	 public void SendMail(String ExcelPath,String Receiver_ID,int r,String ExcelSavePath) throws MessagingException, IOException, InstanceNotFoundException, AttributeNotFoundException, MalformedObjectNameException, ReflectionException, MBeanException, ParseException, NoSuchMethodException, SecurityException
     {
		// ExcelPath="F:\\TiffaProject\\TestExecution_Report\\Test_Execution Report.xls";
		 String ReceivedTime =null, ID = null, MailBody = null;
		 String Subject = null; String Body = null;
         ID = "vivekanand.koli@kalelogistics.in";
     //    Sender = "vivekanand.koli@kalelogistics.in";
         Subject = "Jet Airways-Data Transmission_FWB-FHL";

         Mail pgm = new Mail();
         String mailbody = pgm.MailBody_HTML();

         pgm.SendReportMail(Receiver_ID, mailbody, Subject, ID,ExcelPath,r,ExcelSavePath);
     }

     public String MailBody_HTML() throws UnknownHostException, InstanceNotFoundException, AttributeNotFoundException, MalformedObjectNameException, ReflectionException, MBeanException, ParseException, NoSuchMethodException, SecurityException
     {
    	 Date date =new Date();
    	 
		 String s = new SimpleDateFormat("ddMMyyyy").format(Calendar.getInstance().getTime());
		 Calendar cal = Calendar.getInstance();
		 DateFormat dateFormat = new SimpleDateFormat("dd-MM-yyyy");
		 cal.add(Calendar.DATE, -1);
		 String YesDate=dateFormat.format(cal.getTime());
		 
           
        // String Timestamp;
         StringBuilder Emaibody = new StringBuilder();
         String MachineConfiguration;

     
       
         Emaibody.append("<tr><td>Dear Customer,</td><td style='color: #000'>");
         Emaibody.append("</td></tr>");
         
         Emaibody.append("<tr><td>                                <br/>                                                                                                          </td><td style='color: #000'>");
         Emaibody.append("<tr><td>                                <br/>                                                                                                          </td><td style='color: #000'>");
         
         
         Emaibody.append("<tr><td>As per attached Jet Airways (9W/589) circular requesting you to start processing Jet Airways MAWB/HAWB for all destinations.</td><td style='color: #000'>");
         Emaibody.append("<tr><td>                                <br/>                                                                                                          </td><td style='color: #000'>");
         Emaibody.append("<tr><td>As per instruction from Jet Airways there is no need for a Forwarder to wait for final weight. </td><td style='color: #000'>");
        
         Emaibody.append("<tr><td>                                <br/>                                                                                                          </td><td style='color: #000'>");
         Emaibody.append("<tr><td>                                <br/>                                                                                                          </td><td style='color: #000'>");
         
         Emaibody.append("<tr><td><i><b>Note: Effective 01-Feb-19 the charges are applicable for all Destinations.</i></b></td><td style='color: #000'>"); 
         
         Emaibody.append("<tr><td>                                <br/>                                                                                                          </td><td style='color: #000'>");
         Emaibody.append("<tr><td>                                <br/>                                                                                                          </td><td style='color: #000'>");
         
         
         Emaibody.append("<tr><td>Regards,</td><td style='color: #000'>");
         Emaibody.append("<tr><td>                                <br/>                                                                                                          </td><td style='color: #000'>");
         Emaibody.append("<tr><td>UPLIFT Support</td><td style='color: #000'>");
        
         Emaibody.append("<tr><td> </td><td>");
         String alterateBgColor = "#E2F8FE";
              
        Emaibody.append("</tr>");
         if (alterateBgColor == "#E2F8FE")
         {
             alterateBgColor = "#D0F3FD";
         }
         else
         {
             alterateBgColor = "#E2F8FE";
         }
         String e1=Emaibody.toString();
        return e1;
     }
  
   //  @SuppressWarnings("null")
	public int SendReportMail(String To, String body, String SubjectFileName, String ID,String ExcelPath,int r,String ExcelSavePath)throws MessagingException, IOException
     {

    	   final String username ="vivekanand.koli@kalelogistics.in";
    	   final String password ="kale_1725";

        	 boolean sessionDebug=false;
        	 Properties props1 = System.getProperties();
     		 props1.put("mail.smtp.auth", "true");
     		 props1.put("mail.smtp.starttls.enable", "true");            
             props1.put("mail.smtp.host", "smtp.office365.com");
             props1.put("mail.smtp.port", "587");
            
             props1.put("mail.smtp.starttls.required", "true");
     	     	     		
             javax.mail.Session session= javax.mail.Session.getDefaultInstance(props1, new javax.mail.Authenticator() {
                 protected javax.mail.PasswordAuthentication getPasswordAuthentication()
                 {
                     return new javax.mail.PasswordAuthentication(username, password);	                
                 }
             	 });
             String from = "uplift.support@kalelogistics.in";
             try
             {
            	 
            	 MimeMessage message = new MimeMessage(session);

             	   // Set From: header field of the header.
          	   message.setFrom(new InternetAddress(from));

          	   // Set To: header field of the header.
          	   message.setRecipients(Message.RecipientType.TO,
                        InternetAddress.parse(To));

          	   // Set Subject: header field
          	   message.setSubject(SubjectFileName);	          	 

          	  BodyPart messageBodyPart = new MimeBodyPart();
         	
          	messageBodyPart.setContent(body,"text/html");  
     	  MimeMultipart multipart = new MimeMultipart();   
         	  multipart.addBodyPart(messageBodyPart);
         	  
     	 // Attachment //
         	 
         	 messageBodyPart = new MimeBodyPart(); 
            String filename= ExcelPath;
               DataSource source=new FileDataSource(filename);
             messageBodyPart.setDataHandler(new DataHandler(source));  	 
         messageBodyPart.setFileName("Data Transmission - FHL-FWB.pdf");
     	multipart.addBodyPart(messageBodyPart);
         	message.setContent(multipart);
            // Send message
         	  Transport.send(message);

          	   System.out.println("Email Sent successfully....");	
          String Time = new SimpleDateFormat("HH:mm:ss").format(Calendar.getInstance().getTime());	   
          Status="Pass";
          message = null;
          rw.WriteToExcel(To, Status, r++, ExcelSavePath,Time);
            return 1;		
         }
         catch (Exception ex)
         {
            System.out.println(ex.getMessage());
            Status="Fail";
            String Time = new SimpleDateFormat("HH:mm:ss").format(Calendar.getInstance().getTime());	
            rw.WriteToExcel(To, Status, r++, ExcelSavePath,Time);
                 //Write in Notepad
         }
             return 0;
         }
		

}
