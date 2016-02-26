using System.Data;
using System.Data.SqlClient;
using System.Text;

namespace OpenXml.Excel.Data.Test
{
    class DataHelper
    {
        private const int MaxLength = 8000;

        public static void CreateTableIfNotExists(string connection, string tableName, string[] cols)
        {
            const string pattern = @"IF NOT EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID('{0}') AND [type] in ('U'))
BEGIN
{1}
END";

            var createScript = string.Format(pattern, tableName, GetCreateTableScript(tableName, cols));
            ExecuteNonQuery(connection, createScript);
        }

        private static void ExecuteNonQuery(string connection, string sql)
        {
            using (var conn = new SqlConnection(connection))
            {
                conn.Open();
                using (var cmd = new SqlCommand(sql, conn) { CommandTimeout = 0 })
                    cmd.ExecuteNonQuery();
            }
        }

        private static string GetCreateTableScript(string tableName, string[] cols)
        {
            var sql = new StringBuilder("CREATE TABLE " + tableName);
            sql.AppendLine("(");
            sql.AppendLine("\tLINK int IDENTITY(1,1) NOT NULL,");

            foreach (var col in cols)
                sql.AppendLine("\t[" + col + "] varchar(" + MaxLength + ") NULL,");

            sql.AppendLine("CONSTRAINT PK_" + AdjustTableName(tableName) + " PRIMARY KEY CLUSTERED");
            sql.AppendLine("(LINK ASC)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON)");
            sql.AppendLine(")");

            return sql.ToString();
        }

        private static string AdjustTableName(string tableName)
        {
            return tableName.Replace(".", "_");
        }
    }
}
