using System;
using System.Net.Mail;
using ClosedXML.Excel;
using System.Data.SqlClient;
using ExcelCreator.Services;
using ExcelCreator.ExcelModel;

namespace ExcelCreator
{
    public class ReportMerchantOne
    {
        public static void MerchantOne()
        {
            Connection connection = new Connection();
            Sendmail sendmail = new Sendmail();
            Transactions excelTransactions = new Transactions();
            Refunds excelRefunds = new Refunds();

            string excel = "Merchant One.xlsx";
            var workBook = new XLWorkbook();

            MailMessage message = new MailMessage();
            message.To.Add(new MailAddress($"emailteste@email.com"));

            try
            {
                SqlDataReader sqlTransactions = connection.SqlConnection("EXEC dbo.stp_Report @merchantId='FACFE5B2-F770-4FD6-BCB0-7946C3A9305D', @relatorio='Transactions';");
                excelTransactions.ExcelTransactions(sqlTransactions, workBook);
                connection.ConnectionClose();

                SqlDataReader sqlRefunds = connection.SqlConnection("EXEC dbo.stp_Report @merchantId='FACFE5B2-F770-4FD6-BCB0-7946C3A9305D', @relatorio='Refunds';");
                excelRefunds.ExcelRefunds(sqlRefunds, workBook);
                connection.ConnectionClose();

                workBook.SaveAs(excel);
                Attachment attachment = new Attachment(excel, "application/vnd.ms-excel");

                sendmail.Email(attachment, message, "Relatório MerchantOne", "Olá,\n\nSegue em anexo relatório.\n\nAtenciosamente,");
            }
            catch (Exception e)
            {
                Guid guid = Guid.NewGuid();
                SaveLog.Logs(guid, "MerchantOne", e);
                sendmail.Email(message, "Falha ao enviar relatório", $"Atenção!\n\nOcorreu uma falha ao enviar o relatório da loja MerchantOne, para detalhes utilize o CurrentId: {guid} para consultar os logs.");
            }
        }
    }
}