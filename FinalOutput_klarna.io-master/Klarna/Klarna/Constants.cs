using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Klarna
{
    class Constants
    {
        public const string FormattingColumnNotExists = "<<----COLUMNS TO BE FORMATTED DOES NOT EXIST---->>";

        public const string DeletingColumnNotexists = "<<----COLUMNS FOR DELETION DOES NOT EXIST---->>";

        public const string FileNotExists = "<<---FILE DOES NOT EXISTS--->>";

        public const string KlarnaDateFormat="dd_MM_yyyy";

        public const string Format = "0";

        public const string FormatPaymentAmount = "0.00";

        public const string FormatDate = "yyyy-mm-dd hh:mm:s.000";

        public const string MailBody = @"Hi Klarna Support, <br /><br />In the attached spreadsheet please find today's list of possibly failed refunds.<br />As usual please could you review each of these and let us know if they are at a failed or success status so we can action accordingly.<br />Any refunds which have failed we will retry on our side and any refunds which are successful we will set to complete within Back Office.<br />Please endeavor to reply back to this email within 24 hours so we can action accordingly.<br /><br /><br />Regards,<br />Global Opsbridge";

        public const string MailDateFormat= "dd/MM/yyyy HH:mm:ss";

        public const string MailSubject1= " Daily Klarna Refunds Report for Klarna (automated) (";

        public const string MailSubject2 = " Results Found )  ";

        public const string MailServerAddress= "smtp.office365.com";

        public const string MailSender = "sagar.andre@asos.com";

        public const string MailCredential = "";


    }
}
