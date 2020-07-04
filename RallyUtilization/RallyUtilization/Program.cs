using System;
using System.Configuration;
using System.Net.Mail;

namespace RallyUtilization
{
    internal class Program
    {
        public Program()
        {
        }

        private static void Main(string[] args)
        {
            //data to be manually added
            string emailSubject = "";
            string fromEmailID = "report@report.com";
            string smptpServer = "smptpserver";

            string startDate = ConfigurationManager.AppSettings["IterationStartDate"];
            string endDate = ConfigurationManager.AppSettings["IterationEndDate"];
            string[] toEmailIDs = ConfigurationManager.AppSettings["ToEmail"].Split(new char[] { ';' });
            string[] strArrays = ConfigurationManager.AppSettings["Members"].Split(new char[] { ',' });

            Owner owner;
            string iteration = string.Empty;
            string emailBody = "<table style=\"border-collapse:collapse;border-spacing:0\">   <tr>      " +
                               "<th style=\"font-family:Arial, sans-serif;font-size:14px;font-weight:normal;padding:10px 5px;border-style:solid;border-width:1px;" +
                               "overflow:hidden;word-break:normal;border-color:black\">Iteration</th>      <th style=\"font-family:Arial, sans-serif;font-size:14px;" +
                               "font-weight:normal;padding:10px 5px;border-style:solid;border-width:1px;overflow:hidden;word-break:normal;border-color:black\">" +
                               "Tasks</th>      <th style=\"font-family:Arial, sans-serif;font-size:14px;font-weight:normal;padding:10px 5px;border-style:solid;" +
                               "border-width:1px;overflow:hidden;word-break:normal;border-color:black\">Capacity</th>      " +
                               "<th style=\"font-family:Arial, sans-serif;font-size:14px;font-weight:normal;padding:10px 5px;border-style:solid;border-width:1px;" +
                               "overflow:hidden;word-break:normal;border-color:black;vertical-align:top\">Estimate</th>      " +
                               "<th style=\"font-family:Arial, sans-serif;font-size:14px;font-weight:normal;padding:10px 5px;border-style:solid;border-width:1px;" +
                               "overflow:hidden;word-break:normal;border-color:black;vertical-align:top\">ToDo</th>      <th style=\"font-family:Arial, " +
                               "sans-serif;font-size:14px;font-weight:normal;padding:10px 5px;border-style:solid;border-width:1px;overflow:hidden;" +
                               "word-break:normal;border-color:black;vertical-align:top\">Actuals</th>      <th style=\"font-family:Arial, sans-serif;font-size:14px;" +
                               "font-weight:normal;padding:10px 5px;border-style:solid;border-width:1px;overflow:hidden;word-break:normal;border-color:black;" +
                               "vertical-align:top\">TimeSpent</th>    <th style=\"font-family:Arial, sans-serif;font-size:14px;font-weight:normal;padding:10px 5px;" +
                               "border-style:solid;border-width:1px;overflow:hidden;word-break:normal;border-color:black;vertical-align:top\">Time/Cap<br>%</th> " +
                               "     <th style=\"font-family:Arial, sans-serif;font-size:14px;font-weight:normal;padding:10px 5px;border-style:solid;border-width:1px;" +
                               "overflow:hidden;word-break:normal;border-color:black;vertical-align:top\">Est/Cap<br>%</th>  <th style=\"font-family:Arial, " +
                               "sans-serif;font-size:14px;font-weight:normal;padding:10px 5px;border-style:solid;border-width:1px;overflow:hidden;word-break:normal;" +
                               "border-color:black;vertical-align:top\">Avg. time to finish 1 story point<br>%</th> </tr>";
            string[] strArrays1 = strArrays;
            for (int i = 0; i < (int)strArrays1.Length; i++)
            {
                string member = strArrays1[i];
                owner = new RallyClass(startDate, endDate, member).LoadWorkSpaceDetails();
                if (owner != null)
                {
                    emailBody = string.Concat(new object[] { emailBody,
                        "   <tr>      <td style=\"font-family:Arial, sans-serif;font-size:14px;padding:10px 5px;border-style:solid;border-width:1px;" +
                        "overflow:hidden;word-break:normal;border-color:black\">", owner.name, "</td>      <td style=\"font-family:Arial, sans-serif;font-size:14px;" +
                        "padding:10px 5px;border-style:solid;border-width:1px;overflow:hidden;word-break:normal;border-color:black\">", owner.taskCount, "</td>      " +
                        "<td style=\"font-family:Arial, sans-serif;font-size:14px;padding:10px 5px;border-style:solid;border-width:1px;overflow:hidden;" +
                        "word-break:normal;border-color:black\">", owner.capacity, "</td>      <td style=\"font-family:Arial, sans-serif;font-size:14px;" +
                        "padding:10px 5px;border-style:solid;border-width:1px;overflow:hidden;word-break:normal;border-color:black;vertical-align:top\">",
                        owner.estimate, "</td>      <td style=\"font-family:Arial, sans-serif;font-size:14px;padding:10px 5px;border-style:solid;border-width:1px;" +
                        "overflow:hidden;word-break:normal;border-color:black;vertical-align:top\">", owner.toDo, "</td>      <td style=\"font-family:Arial, " +
                        "sans-serif;font-size:14px;padding:10px 5px;border-style:solid;border-width:1px;overflow:hidden;word-break:normal;border-color:black;" +
                        "vertical-align:top\">", owner.actuals, "</td>	<td style=\"font-family:Arial, sans-serif;font-size:14px;padding:10px 5px;border-style:solid;" +
                        "border-width:1px;overflow:hidden;word-break:normal;border-color:black;vertical-align:top\">", owner.timespent, "</td>      " +
                        "<td style=\"font-family:Arial, sans-serif;font-size:14px;padding:10px 5px;border-style:solid;border-width:1px;overflow:hidden;" +
                        "word-break:normal;border-color:black;vertical-align:top\">", owner.timespentPercentage, "</td>      <td style=\"font-family:Arial, " +
                        "sans-serif;font-size:14px;padding:10px 5px;border-style:solid;border-width:1px;overflow:hidden;word-break:normal;border-color:black;" +
                        "vertical-align:top\">", owner.estimatePercentage, "</td><td style=\"font-family:Arial, sans-serif;font-size:14px;padding:10px 5px;" +
                        "border-style:solid;border-width:1px;overflow:hidden;word-break:normal;border-color:black;" +
                        "vertical-align:top\">", owner.totalPoints != 0 ? Math.Round(owner.timespent / owner.totalPoints, 2) : 0, "</td>   </tr>" });
                    iteration = owner.iteration;
                }
            }
            emailBody = string.Concat(emailBody, "</table>");
            SmtpClient smtpClient = new SmtpClient(smptpServer);//SMTP server needs to go here
            MailMessage mailMessage = new MailMessage()
            {
                From = new MailAddress(fromEmailID)//from mail ID
            };
            for (int i = 0; i < (int)toEmailIDs.Length; i++)
            {
                mailMessage.To.Add(toEmailIDs[i]);
            }
            mailMessage.IsBodyHtml = true;
            mailMessage.Body = emailBody;
            mailMessage.Subject = string.Concat(emailSubject, iteration);
            smtpClient.Send(mailMessage);
        }
    }
}