using System;
using System.Net.Mail;
using ClosedXML.Excel;
using System.Data.SqlClient;
using ExcelCreator.Services;
using ExcelCreator.ExcelModel;

namespace ExcelCreator
{
    public class ReportMerchantTwo
    {
        public static void MerchantTwo()
        {
            Connection connection = new Connection();
            Sendmail sendmail = new Sendmail();
            Transactions excelTransactions = new Transactions();
            Refunds excelRefunds = new Refunds();

            string excel = "Merchant Two.xlsx";
            var workBook = new XLWorkbook();

            MailMessage message = new MailMessage();
            message.To.Add(new MailAddress($"email1@email.com"));
            message.To.Add(new MailAddress($"email2@email.com.br"));

            try
            {
                SqlDataReader sqlTransactions = connection.SqlConnection("EXEC dbo.stp_Report @merchantId='C6915847-6627-4CA0-AA02-7A0C4BA274D3', @relatorio='Transactions';");
                excelTransactions.ExcelTransactions(sqlTransactions, workBook);
                connection.ConnectionClose();

                SqlDataReader sqlRefunds = connection.SqlConnection("EXEC dbo.stp_Report @merchantId='C6915847-6627-4CA0-AA02-7A0C4BA274D3', @relatorio='Refunds';");
                excelRefunds.ExcelRefunds(sqlRefunds, workBook);
                connection.ConnectionClose();

                workBook.SaveAs(excel);
                Attachment attachment = new Attachment(excel, "application/vnd.ms-excel");

                sendmail.Email(attachment, message, "Relatório MerchantTwo", "Olá,\n\nSegue em anexo relatório.\n\nAtenciosamente,");
            }
            catch (Exception e)
            {
                Guid guid = Guid.NewGuid();
                SaveLog.Logs(guid, "MerchantTwo", e);
                sendmail.Email(message, "Falha ao enviar relatório", $"Atenção!\n\nOcorreu uma falha ao enviar o relatório da loja MerchantTwo, para detalhes utilize o CurrentId: {guid} para consultar os logs.");
            }
        }
    }
}