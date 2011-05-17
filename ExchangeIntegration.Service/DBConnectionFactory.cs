using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;

namespace ExchangeIntegration.Service
{
    public interface IDbConnectionFactory
    {
        IDbConnection CreateConnection();
    }

    public class SqlConnectionFactory : IDbConnectionFactory
    {
        public string ConnectionString { get; set; }

        public IDbConnection CreateConnection()
        {
            SqlConnection con = new SqlConnection(ConnectionString);
            return con;
        }
    }
}
