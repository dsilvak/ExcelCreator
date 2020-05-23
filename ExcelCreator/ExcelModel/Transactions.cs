using ClosedXML.Excel;
using System.Data.SqlClient;

namespace ExcelCreator.ExcelModel
{
    class Transactions
    {
        public XLWorkbook ExcelTransactions(SqlDataReader sqlDataReader, XLWorkbook workbook)
        {
            var transactionWs = workbook.Worksheets.Add("Transações");

            transactionWs.Range("A1:H1").Style.Fill.BackgroundColor = XLColor.CelestialBlue;
            transactionWs.Range("A1:H1").Style.Font.Bold = true;

            transactionWs.Cell("A1").Value = "PaymentId";
            transactionWs.Cell("B1").Value = "Cliente";
            transactionWs.Cell("C1").Value = "Valor";
            transactionWs.Cell("D1").Value = "Numero do cartão";
            transactionWs.Cell("E1").Value = "Banceira";
            transactionWs.Cell("F1").Value = "Data de criação";
            transactionWs.Cell("G1").Value = "MerchantId";
            transactionWs.Cell("H1").Value = "Nome da loja";
            transactionWs.Cell("A2").Value = sqlDataReader;

            transactionWs.Columns(1, 9).AdjustToContents();
            transactionWs.Column(3).Style.NumberFormat.Format = "R$#,##";

            return workbook;
        }
    }
}