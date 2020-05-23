using System;
using ExcelCreator.Services;

namespace ExcelCreator
{
    class SaveLog
    {
        public static void Logs(Guid currentId, string merchant, Exception e)
        {
            var connection = new Connection();
            DateTime dateTime = DateTime.UtcNow;

            string exception = e.ToString().Replace("'", "|");

            connection.SqlConnection($"INSERT INTO dbo.ExcelCreatorLogs VALUES ('{currentId}','{merchant}','{dateTime}','{exception}')");
        }
    }
}