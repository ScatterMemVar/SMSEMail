using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.SqlTypes;
using System.Data.SqlClient;
using System.Configuration;
using System.Runtime.CompilerServices;
using MailKit;
using MimeKit;
using MailKit.Net.Smtp;

namespace SMSNotification
{
    /// <summary>
    /// This class handles the creation and sending of E-Mails using either the SQL Server "sp_send_dbmail" system stored procedure
    /// or the ".NET System.Net.Mail" classes.
    /// </summary>
    public class SMSEMail
    {
         public enum Platform
        {
            SQLServer,
            DotNetMail
        }

        List<EMailData> EMailList = null;

        // Properties.

        private string ConnectionString { get; set; }
        private string FromAddress { get; set; }

        private string SMTPServerNameOrIP { get; set; }

        private string SMTPPortNumber { get; set; }

        private string SMTPServerUserName { get; set; }

        private string SMTPServerPword { get; set; }

        // Constructor(s).

        /// <summary>
        /// This class handles the creation and sending of EMail messages using elements stored in the SMS database.  An E-Mail can be sent using SQL Server Database Mail functionality, or
        /// .NET funcitonality.
        /// </summary>
        /// <param name="connectionstring"></param>
        /// <param name="smtpservernameorip"></param>
        /// <param name="smtpportnumber"></param>
        /// <param name="fromaddress"></param>
        /// <param name="smtpserverusername"></param>
        /// <param name="smtpserverpword"></param>
        public SMSEMail(string connectionstring, string smtpservernameorip, string smtpportnumber, string fromaddress, string smtpserverusername = null, string smtpserverpword = null)
        {
            this.ConnectionString = connectionstring;
            this.FromAddress = fromaddress;
            this.SMTPServerNameOrIP = smtpservernameorip;
            this.SMTPPortNumber = smtpportnumber;
            this.SMTPServerUserName = smtpserverusername;
            this.SMTPServerPword = smtpserverpword;
        }

        /// <summary>
        /// This method creates and sends an E-Mail.
        /// </summary>
        /// <param name="platform"></param>
        /// <param name="emailtypeid"></param>
        /// <param name="userid"></param>
        /// <param name="includealternateaddress"></param>
        /// <param name="bbcourseid"></param>
        /// <param name="newstudentaccountpassword"></param>
        /// <param name="ccuserid"></param>
        public async Task CreateAndSendEMail(Platform platform, string sqlmailprofilename, int emailtypeid, int userid, bool includealternateaddress, bool logemail, int? courseid = null, string newstudentaccountpassword = null, int? ccuserid = null)
        {
            try
            {
                if (string.IsNullOrEmpty(this.ConnectionString) || string.IsNullOrEmpty(this.FromAddress) || string.IsNullOrEmpty(this.SMTPServerNameOrIP) || string.IsNullOrEmpty(this.SMTPPortNumber))
                {
                    Exception missingParameterException = new Exception(@"Cannot create E-Mail ... one or more invalid parameters passed to constructor.");
                    throw missingParameterException;
                }

                bool _buildEMailListSuccess = await BuildEMailListAsync(userid, emailtypeid, includealternateaddress, courseid, newstudentaccountpassword, ccuserid);
                
                if (_buildEMailListSuccess)
                {
                    if (EMailList.Count > 0)
                    {
                        foreach (EMailData emailItem in EMailList)
                        {
                            if (platform == Platform.SQLServer)
                            {
                                bool _sendSQLEMailSuccess = await SendSQLEMail(emailItem, sqlmailprofilename);
                                if (_sendSQLEMailSuccess)
                                {
                                    if (logemail == true)
                                    {
                                        await LogEMail(userid, emailtypeid);
                                    }
                                }
                            }
                            else
                            {
                                bool _sendDotNetEMailSuccess = SendDotNetEMail(emailItem, this.SMTPServerUserName, this.SMTPServerPword);
                                if (_sendDotNetEMailSuccess)
                                {
                                    if (logemail == true)
                                    {
                                        await LogEMail(userid, emailtypeid);
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="userid"></param>
        /// <param name="emailtypeid"></param>
        /// <param name="includealternateaddress"></param>
        /// <param name="courseid"></param>
        /// <param name="newstudentaccountpassword"></param>
        /// <param name="ccuserid"></param>
        /// <returns></returns>
        private async Task<bool> BuildEMailListAsync(int userid, int emailtypeid, bool includealternateaddress, int? courseid = null, string newstudentaccountpassword = null, int? ccuserid = null)
        {
            bool _retVal = false;
            try
            {
                using (SqlCommand command = new SqlCommand())
                {
                    command.Connection = new SqlConnection(this.ConnectionString);
                    command.CommandType = CommandType.StoredProcedure;
                    command.CommandText = "usp_GetEMailElements";
                    command.Parameters.AddWithValue("@UserID", userid);
                    command.Parameters.AddWithValue("@EMailTypeID", emailtypeid);
                    command.Parameters.AddWithValue("@GetAlternateAddress", includealternateaddress);
                    command.Parameters.AddWithValue("@CourseID", courseid);
                    command.Parameters.AddWithValue("@NewStudentAccountPWord", newstudentaccountpassword);
                    command.Parameters.AddWithValue("@CCUserID", ccuserid);
                    command.Connection.Open();
                    SqlDataReader rdr = await command.ExecuteReaderAsync();
                    if (rdr.HasRows)
                    {
                        while (rdr.Read())
                        {
                            // Create an EMail for the user's primary address, and
                            // add it to the "EMailList" list.

                            EMailData primaryemaildata = new EMailData();
                            primaryemaildata.FromAddress = this.FromAddress;
                            primaryemaildata.RecipientAddress = rdr.GetString(0);
                            primaryemaildata.Subject_Text = rdr.GetString(2);
                            primaryemaildata.Body_Text = rdr.GetString(3);
                            if (rdr.GetString(4) != null)
                            {
                                primaryemaildata.CCAddress = rdr.GetString(4);
                            }

                            EMailList.Add(primaryemaildata);

                            if (rdr.GetString(1) != null)
                            {
                                // Create an EMail for the user's secondary address, and
                                // add it to the "EMailList" list.

                                EMailData altemaildata = new EMailData();
                                altemaildata.FromAddress = this.FromAddress;
                                altemaildata.RecipientAddress = rdr.GetString(1);
                                altemaildata.Subject_Text = rdr.GetString(2);
                                altemaildata.Body_Text = rdr.GetString(3);
                                if (rdr.GetString(4) != null)
                                {
                                    altemaildata.CCAddress = rdr.GetString(4);
                                }

                                EMailList.Add(altemaildata);
                            }
                        }
                    }
                    rdr.Close();
                    command.Connection.Close();
                }
                _retVal = true;
                return _retVal;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="emaildata"></param>
        /// <param name="profilename"></param>
        /// <returns></returns>
        private async Task<bool> SendSQLEMail(EMailData emaildata, string profilename)
        {
            bool _retVal = false;
            try
            {
                using (SqlCommand command = new SqlCommand())
                {
                    command.Connection = new SqlConnection(this.ConnectionString);
                    command.CommandType = CommandType.StoredProcedure;
                    command.CommandText = "usp_SendDBMail";
                    command.Parameters.AddWithValue("@ProfileName", profilename);
                    command.Parameters.AddWithValue("@RecipientAddress", emaildata.RecipientAddress);
                    command.Parameters.AddWithValue("@FromAddress", this.FromAddress);
                    command.Parameters.AddWithValue("@CCAddress", emaildata.CCAddress);
                    command.Parameters.AddWithValue("@SubjectText", emaildata.Subject_Text);
                    command.Parameters.AddWithValue("@BodyText", emaildata.Body_Text);
                    command.Connection.Open();
                    await command.ExecuteNonQueryAsync();
                    command.Connection.Close();
                }
                 _retVal = true;
                 return _retVal;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="emaildata"></param>
        /// <param name="username"></param>
        /// <param name="password"></param>
        /// <returns></returns>
        private bool SendDotNetEMail(EMailData emaildata, string username = null, string password = null)
        {
            bool _retVal = false;
            try
            {
                var mailMessage = new MimeMessage();
                mailMessage.From.Add(new MailboxAddress("From Name", this.FromAddress));
                mailMessage.To.Add(new MailboxAddress("To Name", emaildata.RecipientAddress));
                mailMessage.Subject = emaildata.Subject_Text;
                mailMessage.Body = new TextPart(MimeKit.Text.TextFormat.Plain)
                {
                    Text = emaildata.Body_Text
                };

                using (var smtpClient = new SmtpClient())
                {
                    smtpClient.Connect(this.SMTPServerNameOrIP, int.Parse(this.SMTPPortNumber));
                    if (username != null && password != null)
                    {
                        smtpClient.Authenticate(username, password);
                    }
                    smtpClient.Send(mailMessage);
                    smtpClient.Disconnect(true);
                }
                _retVal = true;
                return _retVal;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="userid"></param>
        /// <param name="emailtypeid"></param>
        /// <returns></returns>
        private async Task<bool> LogEMail(int userid, int emailtypeid)
        {
            bool _retVal = false;
            try
            {
                using (SqlCommand command = new SqlCommand())
                {
                    command.Connection = new SqlConnection(this.ConnectionString);
                    command.CommandType = CommandType.StoredProcedure;
                    command.CommandText = "usp_InsertEMailLogEntry";
                    command.Parameters.AddWithValue("@UserID", userid);
                    command.Parameters.AddWithValue("@EMailTypeID", emailtypeid);
                    command.Connection.Open();
                    await command.ExecuteNonQueryAsync();
                    command.Connection.Close();
                }
                _retVal = true;
                return _retVal;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
                

        /// <summary>
        /// Each instance of this class represents all of the data needed to create one E-Mail.
        /// </summary>
        internal class EMailData
        {
            // Properties

            internal string RecipientAddress { get; set; }
            internal string FromAddress { get; set; }
            internal string CCAddress { get; set; }
            internal string Subject_Text { get; set; }
            internal string Body_Text { get; set; }

        }
    }
}


