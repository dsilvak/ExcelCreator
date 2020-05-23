using ClosedXML.Excel;
using System.Data.SqlClient;

namespace ExcelCreator.ExcelModel
{
    class Refunds
    {
        public XLWorkbook ExcelRefunds(SqlDataReader sqlDataReader, XLWorkbook workbook)
        {
            var refundsWs = workbook.Worksheets.Add("Reembolsos");

            refundsWs.Range("A1:J1").Style.Fill.BackgroundColor = XLColor.CelestialBlue;
            refundsWs.Range("A1:J1").Style.Font.Bold = true;

            refundsWs.Cell("A1").Value = "PaymentId";
            refundsWs.Cell("B1").Value = "Cliente";
            refundsWs.Cell("C1").Value = "Valor da venda";
            refundsWs.Cell("D1").Value = "Valor do Reembolso";
            refundsWs.Cell("E1").Value = "Numero do cartão";
            refundsWs.Cell("F1").Value = "Bandeira";
            refundsWs.Cell("G1").Value = "Data de criação";
            refundsWs.Cell("H1").Value = "Data de Reembolso";
            refundsWs.Cell("I1").Value = "MerchantId";
            refundsWs.Cell("J1").Value = "Nome da loja";

            refundsWs.Cell("A2").Value = sqlDataReader;

            refundsWs.Columns(1, 10).AdjustToContents();
            refundsWs.Columns(3, 4).Style.NumberFormat.Format = "R$#,##";

            return workbook;
        }
    }
}