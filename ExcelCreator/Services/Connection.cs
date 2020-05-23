using System.Configuration;
using System.Data.SqlClient;

namespace ExcelCreator.Services
{
    class Connection
    {
        SqlConnection connection = new SqlConnection();

        public SqlDataReader SqlConnection(string proc)
        {
            string connectionString = ConfigurationManager.ConnectionStrings["ExcelCreatorConnectionString"].ConnectionString;
            
            connection = new SqlConnection(connectionString);
            connection.Open();
            SqlCommand command = new SqlCommand(proc, connection);
            command.CommandTimeout = 60;
            SqlDataReader dataReader = command.ExecuteReader();

            return dataReader;
        }

        public void ConnectionClose()
        {
            connection.Close();
        }
    }
}