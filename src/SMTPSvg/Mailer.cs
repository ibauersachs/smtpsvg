/*
SMTPSvg - A Mailer Component compatible to ServerObjects ASPMail
Copyright (C) 2010, Ingo Bauersachs

This program is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License
along with this program.  If not, see <http://www.gnu.org/licenses/>.
*/
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Net.Mail;
using System.Net.Mime;
using System.Runtime.InteropServices;

namespace SMTPSvg
{
	[
		ProgId("SMTPSvg.Mailer"),
		ClassInterface(ClassInterfaceType.None),
		Guid("FEFCE318-4CB9-11D2-A2D9-080009AB4447")
	]
	public class Mailer
	{
		#region Fields
		private int _priority;
		private MailAddressCollection recipients = new MailAddressCollection();
		private MailAddressCollection ccs = new MailAddressCollection();
		private MailAddressCollection bccs = new MailAddressCollection();
		private IList<string> attachments = new List<string>();
		private IList<string> headers = new List<string>();
		#endregion

		#region Properties
		/// <summary>
		/// The message body text. To clear the text once you have set it use the ClearBodyText Method.
		/// Example:
		/// Mailer.BodyText = "Your order for 15 widgets has been processed"
		/// </summary>
		public string BodyText { get; set; }

		/// <summary>
		/// The character set. By default the char set is US Ascii
		/// 
		/// Valid values:
		/// * 1 = US ASCII
		/// * 2 = ISO-8859-1
		/// 
		/// Example:
		/// Mailer.CharSet = 2
		/// </summary>
		public int CharSet { get; set; }

		/// <summary>
		/// The ConfirmReading flag. If this is set to true AND the recipients email program supports
		/// this feature (and it is enabled) the recipients email program will send a notice back to
		/// the FromAddress confirming that this email has been read.
		/// 
		/// Example:
		/// 
		/// Mailer.ConfirmRead = true
		/// </summary>
		public bool ConfirmRead { get; set; }

		/// <summary>
		/// The ContentType property allows you to set the ContentType header of the message's BodyText. 
		/// If, for example, you wanted to send HTML as the messages's body, you could set
		/// ContentType = "text/html" and EMail programs that support HTML content could properly display the HTML text.
		/// 
		/// Note: The ContentType property is ignored if you have file attachments.
		/// 
		/// Example:
		/// 
		/// Mailer.ContentType = "text/html"
		/// </summary>
		public string ContentType { get; set; }

		/// <summary>
		/// If you wish to use a character set besides the included types you can set CustomCharSet to a character set string.
		/// 
		/// Example:
		/// 
		/// Mailer.CustomCharSet = "ISO-2022" or
		/// Mailer.CustomCharSet = "big5"
		/// </summary>
		public string CustomCharSet {get; set; }

		/// <summary>
		/// AspMail will, by default, create a Date/Time header for your local system using GMT. 
		/// If you would like to override the date/time calculation set the DateTime property to
		/// a valid date/time string in the format defined by RFC 822 & RFC 1123.
		/// 
		/// Example:
		/// 
		/// Mailer.DateTime = "Fri, 02 May 1997 10:53:49 -0500"
		/// </summary>
		public string DateTime { get; set; }

		/// <summary>
		/// The encoding type for attachments. The default setting is MIME.
		/// 
		/// Valid values:
		/// 
		/// * 1 = UUEncoded
		/// * 2 = MIME
		/// 
		/// Example:
		/// 
		/// Mailer.Encoding = 1
		/// </summary>
		public int Encoding { get; set; }

		/// <summary>
		/// If the component is an eval version the expires property will return the date that the component quits functioning.
		/// Example:
		/// Response.Write "Component Expires: " & Mailer.Expires
		/// </summary>
		public DateTime Expires
		{
			get
			{
				return System.DateTime.MaxValue;
			}
		}

		/// <summary>
		/// The message originator’s name.
		/// 
		/// Example:
		/// 
		/// Mailer.FromName = "Joe’s Widget Shop"
		/// </summary>
		public string FromName { get; set; }

		/// <summary>
		/// The message originator’s email address.
		/// 
		/// Example:
		/// 
		/// Mailer.FromAddress = "joe@widgets.com"
		/// </summary>
		public string FromAddress { get; set; }

		/// <summary>
		/// Defaults to false. When false AspMail will check for '@' in the email address for calls
		///  to AddRecipient, AddCC and AddBCC. An error would be returned in the Response property.
		///  When this property is set to true AspMail will not perform any address syntax validation.
		///  If you are using AspMail to send a message through an SMS gateway or fax system
		///  you may need to set this property to true.
		/// </summary>
		public bool IgnoreMalformedAddress { get; set; }

		/// <summary>
		/// Defaults to true. If true AspMail will ignore error messages returned by the SMTP
		///  server for invalid addresses. This is useful when a mailing is addressed to
		///  a number of recipients.
		/// </summary>
		public bool IgnoreRecipientErrors { get; set; }

		/// <summary>
		/// Live allows you to test the component without an SMTP server. If Live is set
		///  to false then the NET SEND message will be executed with the FromName property
		///  being used as the recipient of the message. Only the subject is sent to the recipient.
		/// 
		/// Example:
		/// 
		/// Mailer.FromName = "ASPProgrammer"
		/// Mailer.Live = false
		/// </summary>
		public bool Live { get; set; }

		/// <summary>
		/// Sets the Organization header in the message.
		/// 
		/// Example:
		/// 
		/// Mailer.Organization = "Your Company Name"
		/// </summary>
		public string Organization { get; set; }

		/// <summary>
		/// The path where PGP is located.
		/// </summary>
		public string PGPPath { get; set; }

		/// <summary>
		/// Parameters that PGP will use to process message.
		/// </summary>
		public string PGPParams { get; set; }

		/// <summary>
		/// Sets the message priority. Priorities are 1-5 and are reflected in the X-Priority
		/// 
		/// Valid values:
		/// 
		/// * 1 – High
		/// * 3 – Normal
		/// * 5 – Low
		/// 
		/// Example:
		/// 
		/// Mailer.Priority = 1
		/// </summary>
		public int Priority
		{
			get { return _priority; }
			set
			{
				if(value != 1 && value != 3 && value != 5)
					throw new COMException("must be 1, 3 or 5");
				_priority = value;
			}
		}

		/// <summary>
		/// The remote SMTP host that the message will be sent through. This is typically
		///  an SMTP server located at your local ISP or it could be an internal SMTP server
		///  on your companies premises. Up to 3 server addresses can be specified, separated
		///  by a semicolon. If the primary server is down the component will attempt to
		///  send the mail using the secondary server and so on.
		/// 
		/// If your RemoteHost uses another port besides 25 then append a colon and port number to the RemoteHost value.
		/// 
		/// Example:
		/// 
		/// Mailer.RemoteHost = "mailhost.myisp.net" or
		/// 
		/// Mailer.RemoteHost = "mailhost.myisp.net;mailhost.myotherisp.net" or
		/// 
		/// Mailer.RemoteHost = "mailhost.myisp.net:160"
		/// </summary>
		public string RemoteHost { get; set; }

		/// <summary>
		/// The ReplyTo property allows you to specify a different email address that replies
		///  should be sent to. By default mail programs should use the Reply-To: header for 
		/// responses if this header is specified.
		/// </summary>
		public string ReplyTo { get; set; }

		/// <summary>
		/// The Response property returns any error messages that may occur.
		/// </summary>
		public string Response { get; set; }

		/// <summary>
		/// The ReturnReceipt flag. If this is set to true AND the recipients SMTP server supports
		/// this feature (and it is enabled) the recipients SMTP server will send a notice back to 
		/// the FromAddress confirming that this email has been delivered.
		/// 
		/// Example:
		/// 
		/// Mailer.ReturnReceipt = false
		/// </summary>
		public bool ReturnReceipt { get; set; }

		/// <summary>
		/// If you need to debug the session give a log file name here. Make sure the IUSR_XYZ 
		/// IIS user has security that allows the component to write to this file. Warning: 
		/// Do not use this setting in situations where multiple users can access this component
		/// at the same time. This is for single user debugging ONLY!
		/// 
		/// Example:
		/// 
		/// Mailer.SMTPLog = "c:\smtplog.txt"
		/// </summary>
		public string SMTPLog { get; set; }

		/// <summary>
		/// The message subject.
		/// 
		/// Example:
		/// 
		/// Mailer.Subject = "Stock split announced!"
		/// </summary>
		public string Subject { get; set; }

		/// <summary>
		/// The SuppressMsgBody property is true by default and is used in conjunction with
		///  the SMTPLog property. When SMTPLog is set to a file and SuppressMsgBody is true 
		/// the log file receives a copy of the message text. If SuppressMsgBody is false the 
		/// message text is not sent to the log.
		/// </summary>
		public bool SuppressMsgBody { get; set; }

		/// <summary>
		/// Timeout is the maximum time that AspMail should wait for a response from the 
		/// remote server. The default is 30 seconds.
		/// 
		/// Example:
		/// 
		/// Mailer.Timeout = 15
		/// </summary>
		public int TimeOut { get; set; }

		/// <summary>
		/// The urgent flag sets the X-Urgent header in the outgoing message. Not all mail readers support this flag.
		/// 
		/// Example:
		/// 
		/// Mailer.Urgent = true
		/// </summary>
		public bool Urgent { get; set; }

		/// <summary>
		/// MS-Mail priority headers, by default, are sent in addition to the standard 
		/// SMTP priority headers. You can turn MS-Mail headers off with this property
		/// 
		/// Example:
		/// 
		/// Mailer.UseMSMailHeaders = false
		/// </summary>
		public bool UseMSMailHeaders { get; set; }

		/// <summary>
		/// Gets the internal component version number.
		/// 
		/// Example:
		/// 
		/// Response.Write "Component Version: " & Mailer.Version
		/// </summary>
		public string Version
		{
			get
			{
				return "4.1.0.0";
			}
		}

		/// <summary>
		/// The WordWrap property is off by default. Setting WordWrap to true causes 
		/// the message body to wordwrap at the position specified by the WordWrapLen property.
		/// </summary>
		public bool WordWrap { get; set; }

		/// <summary>
		/// The WordWrapLen property is set to 70 by default. You can modify the position
		///  that wordwrap occurs by changing this value.
		/// </summary>
		public int WordWrapLen { get; set; }
		#endregion

		#region Constructor
		/// <summary>
		/// Creates a new instance of this class.
		/// </summary>
		public Mailer()
		{
			CharSet = 1;
			IgnoreRecipientErrors = true;
			Live = true;
			SuppressMsgBody = true;
			Priority = 3;
			TimeOut = 30;
			WordWrapLen = 70;
		}
		#endregion

		#region Methods
		/// <summary>
		/// The SendMail method attempts to send the email. 
		/// </summary>
		/// <returns>True or False
		/// 
		/// Example:
		/// 
		/// if Mailer.SendMail then
		/// 
		/// Response.Write "Mail sent..."
		/// 
		/// else
		/// 
		/// Response.Write "Mail failure. Check mail host server name and tcp/ip connection..."
		/// 
		/// end if</returns>
		public bool SendMail()
		{
			try
			{
				if(PGPPath != null)
				{
					Process p = new Process
					            	{
					            		StartInfo =
					            			{
					            				CreateNoWindow = true,
					            				FileName = PGPPath,
					            				Arguments = PGPParams
					            			}
					            	};
					p.Start();
				}

				MailMessage msg = new MailMessage();

				msg.Body = BodyText;
				if(CustomCharSet != null)
				{
					msg.BodyEncoding = System.Text.Encoding.GetEncoding(CustomCharSet);
				}
				else
				{
					switch(CharSet)
					{
						case 1: //us ascci
							msg.BodyEncoding = System.Text.Encoding.ASCII;
							break;
						case 2: //iso-8859-1
							msg.BodyEncoding = System.Text.Encoding.GetEncoding("iso-8859-1");
							break;
					}
				}
				msg.IsBodyHtml = ContentType == "text/html";
				if(ConfirmRead)
					msg.Headers.Add("Disposition-Notification-To", "<" + FromAddress + ">");

				if(this.DateTime != null)
					msg.Headers.Add("Date", this.DateTime);

				if(Organization != null)
					msg.Headers.Add("Organization", Organization);

				msg.Sender = new MailAddress(FromAddress, FromName);
				msg.From = msg.Sender;

				foreach(var addr in recipients)
					msg.To.Add(addr);
				foreach(var addr in ccs)
					msg.CC.Add(addr);
				foreach(var addr in bccs)
					msg.Bcc.Add(addr);
				foreach(var filename in attachments)
				{
					Attachment data = new Attachment(filename);
					// Add time stamp information for the file.
					ContentDisposition disposition = data.ContentDisposition;
					disposition.CreationDate = File.GetCreationTime(filename);
					disposition.ModificationDate = File.GetLastWriteTime(filename);
					disposition.ReadDate = File.GetLastAccessTime(filename);

					msg.Attachments.Add(data);
				}
				foreach(var header in headers)
					msg.Headers.Add(header.Split(new string[] {": "}, StringSplitOptions.None)[0], header.Split(new string[] {": "}, StringSplitOptions.None)[1]);

				switch(Priority)
				{
					case 1:
						msg.Priority = MailPriority.High;
						break;
					case 3:
						msg.Priority = MailPriority.Normal;
						break;
					case 5:
						msg.Priority = MailPriority.Low;
						break;
				}

				if (ReplyTo != null)
					msg.ReplyTo = new MailAddress(ReplyTo);

				if(ReturnReceipt)
					msg.DeliveryNotificationOptions = DeliveryNotificationOptions.OnFailure | DeliveryNotificationOptions.OnSuccess | DeliveryNotificationOptions.Delay;

				msg.Subject = Subject;

				SmtpClient smtp = new SmtpClient();
				smtp.Timeout = TimeOut;
				foreach(string server in RemoteHost.Split(';'))
				{
					if(server.IndexOf(':') != -1)
					{
						smtp.Host = server.Split(':')[0];
						smtp.Port = Int32.Parse(server.Split(':')[1]);
					}
					else
						smtp.Host = server;

					try
					{
						smtp.Send(msg);
						return true;
					}
					catch(Exception ex)
					{
						Response = ex.Message;
					}
				}
			}
			catch(Exception ex)
			{
				Response = ex.Message;
			}
			return false;
		}

		/// <summary>
		/// Adds a new recipient, as shown in the message's To:  list.
		/// Mailer.AddRecipient "Jay Jones", "jayj@somehost.net"
		/// </summary>
		/// <param name="name">"Jay Jones"</param>
		/// <param name="address">"jayj@somehost.net"</param>
		/// <returns>True/False based upon success or failure.</returns>
		public bool AddRecipient(string name, string address)
		{
			if(String.IsNullOrEmpty(address) || (!IgnoreMalformedAddress && address.IndexOf('@') == -1))
				return false;
			recipients.Add(new MailAddress(address, name));
			return true;
		}

		/// <summary>
		/// Clears any recipients assigned to the To list.
		/// </summary>
		public void ClearRecipients()
		{
			recipients.Clear();
		}

		/// <summary>
		/// Adds a new recipient, as shown in the message's CC list.
		/// Mailer.AddCC "Jay Jones", "jayj@somehost.net"
		/// </summary>
		/// <param name="name">"Jay Jones"</param>
		/// <param name="address">"jayj@somehost.net"</param>
		/// <returns>True/False based upon success or failure.</returns>
		public bool AddCC(string name, string address)
		{
			if(String.IsNullOrEmpty(address) || (!IgnoreMalformedAddress && address.IndexOf('@') == -1))
				return false;
			ccs.Add(new MailAddress(address, name));
			return true;
		}

		/// <summary>
		/// Clears any recipients assigned to the CC list.
		/// </summary>
		public void ClearCCs()
		{
			ccs.Clear();
		}

		/// <summary>
		/// Adds a new Blind Carbon Copy recipient. BCC recipients are not shown in any message recipient list.
		/// Mailer.AddBCC "Jay Jones", "jayj@somehost.net"
		/// </summary>
		/// <param name="name">"Jay Jones"</param>
		/// <param name="address">"jayj@somehost.net"</param>
		/// <returns>True/False based upon success or failure.</returns>
		public bool AddBCC(string name, string address)
		{
			if(String.IsNullOrEmpty(address) || (!IgnoreMalformedAddress && address.IndexOf('@') == -1))
				return false;
			bccs.Add(new MailAddress(address, name));
			return true;
		}

		/// <summary>
		/// Clears any recipients assigned to the BCC list.
		/// </summary>
		public void ClearBCCs()
		{
			bccs.Clear();
		}

		/// <summary>
		/// Clears all recipients assigned to the To, CC  and BCC lists.
		/// </summary>
		public void ClearAllRecipients()
		{
			ClearRecipients();
			ClearCCs();
			ClearBCCs();
		}

		/// <summary>
		/// Adds attachments to current mailing. You must use an explicit path to attach files. UNC's are not legal.
		/// 
		/// Make sure that the IUSR_XYZ IIS user, or the authenticated user has security rights
		/// that allow the component to read the necessary files!
		/// Note: Attachments may not be supported in some eval versions.
		/// 
		/// Example:
		/// 
		/// Mailer.AddAttachment "d:\shipping\proddsk1.zip"
		/// </summary>
		/// <param name="filename">Filename to attach to message.</param>
		public void AddAttachment(string filename)
		{
			attachments.Add(filename);
		}

		/// <summary>
		/// Clears any attachments that were previously set.
		/// 
		/// Example:
		/// 
		/// Mailer.ClearAttachments
		/// </summary>
		public void ClearAttachments()
		{
			attachments.Clear();
		}

		/// <summary>
		/// Clears any text assigned to the message’s body which may have been set previously by using the BodyText property.
		/// </summary>
		public void ClearBodyText()
		{
			BodyText = null;
		}

		/// <summary>
		/// Clears any X-Headers that were set by use of AddExtraHeader.
		/// </summary>
		public void ClearExtraHeaders()
		{
			headers.Clear();
		}

		/// <summary>
		/// Adds extra X-Headers to the mail envelope.
		/// 
		/// Example:
		/// 
		/// Mailer.AddExtraHeader("X-HeaderName: XHdrValue")
		/// </summary>
		/// <param name="header">A string value that forms a proper SMTP X-Header</param>
		/// <returns>True or false. Returns true if X-Header was added.</returns>
		public bool AddExtraHeader(string header)
		{
			headers.Add(header);
			return true;
		}

		/// <summary>
		/// This method will execute PGP (if you've set the PGP parameters up)
		/// and then load the specified file into the message body text. If EraseFile 
		/// is true then the file is erased once the file is loaded. ShowWindow can
		/// be set to true for debugging purposes but it is suggested that you turn
		/// it off once the component has been configured properly.
		/// </summary>
		/// <param name="filename"></param>
		/// <param name="deleteFile"></param>
		/// <param name="showWindow"></param>
		/// <returns>See pgpmail.asp for more information.</returns>
		public bool GetBodyTextFromFile(string filename, bool deleteFile, bool showWindow)
		{
			BodyText = File.ReadAllText(filename);
			if(deleteFile)
				File.Delete(filename);
			return true;
		}

		/// <summary>
		/// Encodes a string in RFC1522 format to provide support for 8bit mail
		///  headers such as 8bit subject headers.
		/// 
		/// Example:
		/// 
		/// Mailer.Subject = Mailer.EncodeHeader("Résponse de Service à la clientèle")
		/// </summary>
		/// <param name="header">string to encode</param>
		/// <returns>Encoded string</returns>
		public string EncodeHeader(string header)
		{
			return header;
		}

		/// <summary>
		/// Returns the path set up by the OS for temporary mail files. 
		/// See the discussion on TMP env variables for more information.
		/// </summary>
		/// <returns>OS's temp path</returns>
		public string GetTempPath()
		{
			return Environment.GetEnvironmentVariable("TMP");
		}
		#endregion
	}
}