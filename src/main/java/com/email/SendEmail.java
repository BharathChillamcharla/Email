package com.email;


import org.apache.commons.mail.DefaultAuthenticator;
import org.apache.commons.mail.Email;
import org.apache.commons.mail.EmailException;
import org.apache.commons.mail.SimpleEmail;

import javax.activation.MailcapCommandMap;
import javax.mail.Message;
public class SendEmail {

	public static void main(String[] args) throws EmailException {
		
//		StringBuilder sb = new StringBuilder();
//		sb.append("<table><tr><td>your content here</td></tr></table>");

		String text= "<table><tr><td>EmpId</td><td>Emp name</td><td>age</td></tr><tr><td>value</td><td>value</td><td>value</td></tr></table>";
		Email email=new SimpleEmail();
		//MailMessage mail = new MailMessage();
		
		email.setHostName("smtp.googlemail.com");
		email.setSmtpPort(465);
		email.setAuthenticator(new DefaultAuthenticator("chillamcharla.bharath@gmail.com", ""));
		email.setSSL(true);
		email.setFrom("chillamcharla.bharath@gmail.com");
		email.setSubject("TestMail");
		email.setMsg(text);
		
		email.addTo("chillamcharla.bharath@gmail.com");
		email.send();
		
		System.out.println("Email Sent Sucessfully");
		
}

}
