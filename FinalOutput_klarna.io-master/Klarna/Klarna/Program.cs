using System;
using System.Configuration;
using System.Runtime.InteropServices;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using System.Net.Mail;
using System.Collections.Generic;

namespace Klarna
{
    class Program
    {
        static Excel.Application xlApp;
        static Excel.Workbook xlWorkBook;
        static Excel.Worksheet xlWorkSheet;
        static Excel.Range range;

        static int columnCnt = 0;
        static int rowCnt = 0;
      
        

        private static string CsvFileName = ConfigurationManager.AppSettings["csvFileName"];
        private static string TsvFileName = ConfigurationManager.AppSettings["tsvFileName"];
        private static string ExcelFileName = ConfigurationManager.AppSettings["excelFileName"];

        static List<string> ccEmailAddresses = new List<string>() { "yogesh.jadhav@asos.com", "keyur.patel@asos.com" };
       

        static void Main(string[] args)
        {


            FileInfo fileInfo1 = new FileInfo(CsvFileName);
            Delete(TsvFileName);
            ConvertCSVtoTabDelimited(fileInfo1);



            FileInfo fileInfo2 = new FileInfo(TsvFileName);
            Delete(ExcelFileName);
            ConvertTSVtoEXCEL(fileInfo2);




            if (File.Exists(ExcelFileName))
            {
                string klarnaFileName = ConfigurationManager.AppSettings["OutputFile"] + DateTime.Now.ToString(Constants.KlarnaDateFormat) + ".xlsx";

                Delete(klarnaFileName);

                xlApp = new Excel.Application();
                xlWorkBook = xlApp.Workbooks.Open(ExcelFileName);
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                range = xlWorkSheet.UsedRange;
                columnCnt = range.Columns.Count;
                rowCnt = range.Rows.Count;



                int targetTeamIndex = GetColumnIndex("TargetTeam");
                int batchNumberIndex = GetColumnIndex("batchNumber");



                if (targetTeamIndex > 0 && batchNumberIndex > 0)
                {
                    ((Excel.Range)xlWorkSheet.Columns[targetTeamIndex]).EntireColumn.Delete(null);

                    ((Excel.Range)xlWorkSheet.Columns[batchNumberIndex - 1]).EntireColumn.Delete(null);

                    int invoiceNumberIndex = GetColumnIndex("InvoiceNumber");
                    int receiptIdIndex = GetColumnIndex("ReceiptId");
                    int dateEnteredIndex = GetColumnIndex("dateentered");
                    int voidHeaderIdIndex = GetColumnIndex("VoidHeaderId");
                    int PaymentAmountIndex = GetColumnIndex("PaymentAmount");


                    if (invoiceNumberIndex > 0 && receiptIdIndex > 0 && dateEnteredIndex > 0 && voidHeaderIdIndex > 0)
                    {
                     

                        //Formatting

                        xlWorkSheet.Columns[invoiceNumberIndex].NumberFormat = Constants.Format;
                        xlWorkSheet.Columns[receiptIdIndex].NumberFormat = Constants.Format;
                        xlWorkSheet.Columns[dateEnteredIndex].NumberFormat = Constants.FormatDate;
                        xlWorkSheet.Columns[voidHeaderIdIndex].NumberFormat = Constants.Format;
                        xlWorkSheet.Columns[PaymentAmountIndex].NumberFormat = Constants.FormatPaymentAmount;



                        string filename = ConfigurationManager.AppSettings["OutputFile"] + DateTime.Now.ToString(Constants.KlarnaDateFormat);
                        xlWorkBook.SaveAs(filename + ".xlsx");


                        Marshal.ReleaseComObject(range);

                        Marshal.ReleaseComObject(xlWorkSheet);

                        xlWorkBook.Close();

                        Marshal.ReleaseComObject(xlWorkBook);

                        xlApp.Quit();

                        Marshal.ReleaseComObject(xlApp);

                      // SendMail(Constants.MailSender, ccEmailAddresses, filename + ".xlsx");
                        
                    }
                    else
                    {

                        Console.WriteLine(Constants.FormattingColumnNotExists);
                    }
                }
                else
                {

                    Console.WriteLine(Constants.DeletingColumnNotexists);

                }
            }
            else
            {

                Console.WriteLine(Constants.FileNotExists);

            }

            
        }


        private static void ConvertCSVtoTabDelimited(FileInfo fi)// csv-tsv
        {
            try
            {
                string NewFileName = Path.Combine(Path.GetDirectoryName(fi.FullName), Path.GetFileNameWithoutExtension(fi.FullName) + ".tsv");
                File.WriteAllText(NewFileName, System.IO.File.ReadAllText(fi.FullName).Replace(",", "\t"));
            }
            catch (Exception ex)
            {
                Console.WriteLine("File: " + fi.FullName + Environment.NewLine + Environment.NewLine + ex.ToString(), "Error Converting csv File to tsv");
            }


        }
        public static void ConvertTSVtoEXCEL(FileInfo fi1)
        {

            string worksheetsName = "sheet1";
            var format = new ExcelTextFormat();
            format.Delimiter = '\t';
            format.EOL = "\r";
            format.DataTypes = new eDataTypes[] { eDataTypes.String, eDataTypes.String,
                eDataTypes.String, eDataTypes.String, eDataTypes.String,
                eDataTypes.String, eDataTypes.String, eDataTypes.String,
                eDataTypes.String, eDataTypes.String
                };

            using (ExcelPackage package = new ExcelPackage(new FileInfo(ExcelFileName)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(worksheetsName);
                
             
                worksheet.Cells["A1"].LoadFromText(new FileInfo(TsvFileName), format);
                package.Save();
            }     
  


        }


 


        public static void Delete(string filename)
        {
            if (File.Exists(filename))
            {
                File.Delete(filename);
            }
        }

       
       public static int GetColumnIndex(string str)
        {
            int output = 0;

            for (int col = 1; col <= columnCnt; col++)
            {

                if (xlWorkSheet.Cells[1, col].value == str)
                {
                    output = col;
                    return output;
                }

            }

            return 0;
        }
       
        private static void SendMail(string Toemail,List<string> ccEmails, string filePath)
        {

            string subject = DateTime.Now.ToString(Constants.MailDateFormat) + Constants.MailSubject1 + rowCnt + Constants.MailSubject2;

            var SmtpServer = new SmtpClient(Constants.MailServerAddress);

            var froMailAddress = new MailAddress(Constants.MailSender);
            var toMailAddress = new MailAddress(Toemail);

            var mailMessage = new MailMessage(froMailAddress, toMailAddress)
            {
                Subject = subject,

            };
            Attachment attachment;
            attachment = new Attachment(filePath);
            mailMessage.Attachments.Add(attachment);

            foreach (var emailAddress in ccEmails)
            {
                mailMessage.CC.Add(emailAddress);
            }
            
            mailMessage.Body = Constants.MailBody;
            mailMessage.IsBodyHtml = true;

            SmtpServer.Credentials = new System.Net.NetworkCredential(Constants.MailSender,Constants.MailCredential);
            SmtpServer.EnableSsl = true;
            SmtpServer.Send(mailMessage);

        }

    }
}

       
