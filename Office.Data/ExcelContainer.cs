namespace Ideal.Office.Data
{
    using System.Data;
    using System.Data.SqlClient;

    public class ExcelContainer
    {
        public static DataTable ToDataTable(string query, params object[] parameters) {
            using (var helper = DbHelper.Create())
            {
                return helper.GetDataTable(query, parameters);
            }
        }

    }
}
