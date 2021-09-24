using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;

namespace MakeReport.Entities
{
    class SqlTools
    {
        public static DataTable SelectWithParameter(string query, List<string> parameters, string connectionString)
        {
            DataTable dataTable = new DataTable();
            List<string> paramNames = new List<string>();
            List<SqlParameter> sqlParameters = new List<SqlParameter>();

            for (int j = 0; j < parameters.Count; j++)
            {
                SqlParameter param = new SqlParameter();

                param.ParameterName = "@tag" + j.ToString();
                param.Value = parameters[j];

                sqlParameters.Add(param);
                paramNames.Add("@tag" + j.ToString());
            }

            // Formata o paramNames para o SQL
            string inClause = string.Join(", ", paramNames);

            string queryWithParameters = string.Format(query, inClause);

            // Conexão do SQL Server
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                // Abre conexão com SQL
                conn.Open();

                SqlCommand cmd = new SqlCommand(queryWithParameters, conn);

                foreach (SqlParameter p in sqlParameters)
                {
                    cmd.Parameters.Add(p);
                }
                
                using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                {
                    da.Fill(dataTable);
                }

                // Fecha Conexão com SQL
                conn.Close();
            }

            return dataTable;
        }
    }
}
