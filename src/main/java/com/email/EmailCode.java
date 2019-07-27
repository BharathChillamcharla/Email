package com.email;

import java.awt.Desktop;
import java.io.File;
import java.io.FileWriter;

public class EmailCode {
	public static void main(String args[]) throws Exception {
		
		
		String attach = System.getProperty("user.dir") + "\\src\\test\\resources\\Submission Report\\MetricSubmissionReport.xlsx";
		
		sentMail(false, "chillamcharla.bharath@gmail.com", "chillamcharla.bharath@gmail.com", "chillamcharla.bharath@gmail.com", "Test", "Test\n\nEmailingReport", attach);
	}

	public static void sentMail(Boolean shownOnly, String toAddressList, String replyAddressList, String ccAddressList,
			String subject, String body, String attach) throws Exception {
		if (toAddressList == null && ccAddressList == null) {
			throw new Exception("Address not found");
		}
		StringBuilder script = new StringBuilder();
		script.append("Dim objOutlook\n").append("set objOutlook = CreateObject(\"Outlook.Application\")\n")
				.append("Dim objOutlookMsg\n").append("Set objOutlookMsg = objOutlook.CreateItem(olMailItem)\n")
				.append("On Error resume next\n").append("objOutlookMsg.ReplyRecipients.Count\n")
				.append("If Err.Number <> 0 Then\n")
				.append("  MsgBox \"Please start your Outlook client and retry.\", 0,\"Failed to sent mail\"\n")
				.append("  Err.clear \n").append("Else \n").append("On Error goto 0 \n");

		if (replyAddressList != null && replyAddressList.length() > 0) {
			String[] replyToS = replyAddressList.split(";");
			for (int i = 1; i <= replyToS.length; i++) {
				script.append("objOutlookMsg.ReplyRecipients.Add(\"").append(replyToS[i - 1]).append("\")\n");
			}
		}
		if (toAddressList != null && toAddressList.length() > 0)
			script.append("objOutlookMsg.To= \"").append(toAddressList).append("\"\n");
		if (ccAddressList != null && ccAddressList.length() > 0)
			script.append("objOutlookMsg.Cc= \"").append(ccAddressList).append("\"\n");
		if (subject != null) {
			script.append("objOutlookMsg.Subject = \"").append(subject.replace("\"", "\"\"").replace("\n", " "))
					.append("\"\n");
		}
		if (body != null) {
			script.append("objOutlookMsg.Body = \"")
					.append(body.replace("\"", "\"\"").replace("\n", "\"&vbCr&vbLf&\"").replace("\r", ""))
					.append("\"\n");
		}
		if (attach != null) {
			// for (String fileName : attach) {
			File f = new File(attach);
			if (!(f.exists() && f.isFile())) {
				throw new Exception("Invalid File");
			}
			script.append("Set myAttachments = objOutlookMsg.Attachments\n");
			script.append("myAttachments.Add \"").append(f.getAbsolutePath()).append("\"\n");
			// }
		}
		script.append("objOutlookMsg.").append(shownOnly ? "display\n" : "send\n")
				.append("set objOutlookMsg = Nothing \n").append("set objOutlook = Nothing \n").append("end if\n");
		String s = script.toString();
		File temp = File.createTempFile("OutMail", ".vbs");
		//
		temp.deleteOnExit();
		FileWriter writer = new FileWriter(temp);
		writer.write(s);
		writer.close();
		Desktop d = Desktop.getDesktop();
		d.open(temp);
	}

}
